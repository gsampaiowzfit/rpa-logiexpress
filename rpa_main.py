"""
RPA LogiExpress Brasil — Robô de Validação Cadastral, Enriquecimento via API e Cálculo de Frete.

Requisitos atendidos de forma simples e didática:
1. Leitura e geração de arquivos Excel (pandas + openpyxl)
2. Expressões Regulares (re) para validação e formatação de CPF, E-mail e CEP
3. Integração com API REST pública (ViaCEP via requests)
4. Web Scraping de página HTML da Intranet (BeautifulSoup4)
5. Tratamento resiliente de erros por registro (try-except + exceção customizada)
6. Geração detalhada de logs passo a passo (logging em arquivo e console)
7. Persistência e auditoria em Banco de Dados Relacional (SQLite)
"""

from datetime import datetime
import logging
from pathlib import Path
import re
import sqlite3
import time
import uuid

from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import requests


# ==============================================================================
# CONFIGURAÇÃO DE DIRETÓRIOS E ARQUIVOS
# ==============================================================================
BASE_DIR = Path(__file__).resolve().parent
ARQUIVO_ENTRADA = BASE_DIR / "dados" / "clientes_pendentes.xlsx"
ARQUIVO_SAIDA = BASE_DIR / "dados" / "relatorio_processamento.xlsx"
ARQUIVO_PORTAL_HTML = BASE_DIR / "portal_intranet" / "tabela_frete_portal.html"
ARQUIVO_PAINEL_HTML = BASE_DIR / "portal_intranet" / "painel_resultado.html"
ARQUIVO_BANCO = BASE_DIR / "banco" / "rpa_logiexpress.db"
ARQUIVO_LOG = BASE_DIR / "logs" / "rpa_execucao.log"


# ==============================================================================
# EXCEÇÃO CUSTOMIZADA PARA REGRAS DE NEGÓCIO
# ==============================================================================
class ErroRegraNegocio(Exception):
    """Exceção disparada quando um registro viola uma regra de negócio."""


# ==============================================================================
# CONTINGÊNCIA OFFLINE PARA API VIACEP (USADA APENAS SE NÃO HOUVER INTERNET)
# ==============================================================================
CONTINGENCIA_VIACEP = {
    "01001000": {
        "cep": "01001-000",
        "logradouro": "Praça da Sé",
        "bairro": "Sé",
        "localidade": "São Paulo",
        "uf": "SP",
    },
    "20040002": {
        "cep": "20040-002",
        "logradouro": "Avenida Rio Branco",
        "bairro": "Centro",
        "localidade": "Rio de Janeiro",
        "uf": "RJ",
    },
    "30140071": {
        "cep": "30140-071",
        "logradouro": "Rua dos Inconfidentes",
        "bairro": "Savassi",
        "localidade": "Belo Horizonte",
        "uf": "MG",
    },
    "80020310": {
        "cep": "80020-310",
        "logradouro": "Rua Marechal Deodoro",
        "bairro": "Centro",
        "localidade": "Curitiba",
        "uf": "PR",
    },
    "40020000": {
        "cep": "40020-000",
        "logradouro": "Praça da Piedade",
        "bairro": "Centro",
        "localidade": "Salvador",
        "uf": "BA",
    },
    "70040010": {
        "cep": "70040-010",
        "logradouro": "SBN Quadra 1",
        "bairro": "Asa Norte",
        "localidade": "Brasília",
        "uf": "DF",
    },
    "99999999": {"erro": True},
}


# ==============================================================================
# CONFIGURAÇÃO DO SISTEMA DE LOGS
# ==============================================================================
def configurar_logger() -> logging.Logger:
    """Configura o logger para gravar cada passo em arquivo e exibir no console."""
    ARQUIVO_LOG.parent.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("RPA_LogiExpress")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formato = logging.Formatter(
        "%(asctime)s | %(levelname)-7s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    arquivo_handler = logging.FileHandler(ARQUIVO_LOG, mode="w", encoding="utf-8")
    arquivo_handler.setFormatter(formato)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formato)

    logger.addHandler(arquivo_handler)
    logger.addHandler(console_handler)
    return logger


# ==============================================================================
# CAMADA DE BANCO DE DADOS (SQLITE)
# ==============================================================================
def inicializar_banco(caminho_banco: Path) -> sqlite3.Connection:
    """Cria o banco SQLite e a tabela de auditoria, caso ainda não existam."""
    caminho_banco.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(caminho_banco)
    cursor = conn.cursor()
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS historico_processamento (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            id_execucao TEXT NOT NULL,
            id_cliente TEXT NOT NULL,
            nome TEXT NOT NULL,
            cpf_formatado TEXT,
            email TEXT,
            cep_formatado TEXT,
            endereco_completo TEXT,
            uf TEXT,
            valor_pedido REAL,
            valor_frete REAL,
            valor_total REAL,
            status TEXT NOT NULL,
            mensagem_detalhe TEXT NOT NULL,
            data_hora TEXT NOT NULL
        )
        """
    )
    conn.commit()
    return conn


def salvar_registro_banco(conn: sqlite3.Connection, registro: dict) -> None:
    """Insere um registro de processamento (SUCESSO ou INSUCESSO) no SQLite."""
    cursor = conn.cursor()
    cursor.execute(
        """
        INSERT INTO historico_processamento (
            id_execucao, id_cliente, nome, cpf_formatado, email,
            cep_formatado, endereco_completo, uf, valor_pedido,
            valor_frete, valor_total, status, mensagem_detalhe, data_hora
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            registro["id_execucao"],
            registro["id_cliente"],
            registro["nome"],
            registro["cpf_formatado"],
            registro["email"],
            registro["cep_formatado"],
            registro["endereco_completo"],
            registro["uf"],
            registro["valor_pedido"],
            registro["valor_frete"],
            registro["valor_total"],
            registro["status"],
            registro["mensagem_detalhe"],
            registro["data_hora"],
        ),
    )
    conn.commit()


# ==============================================================================
# WEB SCRAPING COM BEAUTIFULSOUP (INTRANET)
# ==============================================================================
def carregar_tabela_frete_html(caminho_html: Path, logger: logging.Logger) -> dict[str, float]:
    """Lê a tabela HTML de fretes da intranet usando BeautifulSoup4."""
    logger.info("Lendo tabela oficial de frete via BeautifulSoup: %s", caminho_html.name)
    conteudo = caminho_html.read_text(encoding="utf-8")
    soup = BeautifulSoup(conteudo, "html.parser")

    tabela = soup.find("table", id="tabela-frete")
    if not tabela:
        raise RuntimeError("Tabela 'tabela-frete' não encontrada no portal HTML.")

    tabela_fretes: dict[str, float] = {}
    for linha in tabela.select("tbody tr.linha-uf"):
        uf = linha.get("data-uf", "").strip().upper()
        celula_valor = linha.select_one("td.valor-frete")
        if uf and celula_valor:
            tabela_fretes[uf] = float(celula_valor.text.strip())

    logger.info("Tabela de frete carregada com %d UFs: %s", len(tabela_fretes), tabela_fretes)
    return tabela_fretes


# ==============================================================================
# VALIDAÇÕES E FORMATAÇÕES COM EXPRESSÕES REGULARES (REGEX)
# ==============================================================================
def validar_e_formatar_cpf(cpf_bruto: str) -> str:
    """Valida se o CPF possui 11 dígitos válidos e retorna no padrão 000.000.000-00."""
    digitos = re.sub(r"\D", "", str(cpf_bruto or ""))
    if not re.match(r"^\d{11}$", digitos):
        raise ErroRegraNegocio(
            f"CPF inválido ('{cpf_bruto}'): deve conter exatamente 11 dígitos numéricos."
        )
    if len(set(digitos)) == 1:
        raise ErroRegraNegocio(
            f"CPF inválido ('{cpf_bruto}'): sequência repetida não permitida."
        )
    return re.sub(r"(\d{3})(\d{3})(\d{3})(\d{2})", r"\1.\2.\3-\4", digitos)


def validar_email(email_bruto: str) -> str:
    """Valida o formato do e-mail utilizando Expressão Regular."""
    email_limpo = str(email_bruto or "").strip().lower()
    padrao_email = r"^[\w\.-]+@[\w\.-]+\.\w{2,}$"
    if not re.match(padrao_email, email_limpo):
        raise ErroRegraNegocio(
            f"E-mail inválido ('{email_bruto}'): formato não atende ao padrão usuário@domínio.ext."
        )
    return email_limpo


def validar_e_formatar_cep(cep_bruto: str) -> tuple[str, str]:
    """Valida o CEP (8 dígitos) e retorna tupla (somente_digitos, formatado 00000-000)."""
    digitos = re.sub(r"\D", "", str(cep_bruto or ""))
    if not re.match(r"^\d{8}$", digitos):
        raise ErroRegraNegocio(
            f"CEP em formato inválido ('{cep_bruto}'): deve possuir exatamente 8 dígitos."
        )
    cep_formatado = re.sub(r"(\d{5})(\d{3})", r"\1-\2", digitos)
    return digitos, cep_formatado


# ==============================================================================
# INTEGRAÇÃO COM API REST PÚBLICA (VIACEP)
# ==============================================================================
def consultar_api_viacep(cep_limpo: str, logger: logging.Logger) -> dict:
    """Consulta endereço completo na API ViaCEP com fallback de contingência se offline."""
    url = f"https://viacep.com.br/ws/{cep_limpo}/json/"
    try:
        resposta = requests.get(url, timeout=6)
        resposta.raise_for_status()
        dados = resposta.json()
    except requests.RequestException as exc:
        logger.warning(
            "Sem resposta online da API ViaCEP (%s). Usando contingência local para CEP %s.",
            exc.__class__.__name__,
            cep_limpo,
        )
        dados = CONTINGENCIA_VIACEP.get(cep_limpo)
        if dados is None:
            raise ErroRegraNegocio(
                f"Falha ao consultar API ViaCEP para CEP {cep_limpo}: {exc}"
            ) from exc

    if dados.get("erro") in (True, "true", "True"):
        raise ErroRegraNegocio(
            f"CEP {cep_limpo} não encontrado na base dos Correios/ViaCEP."
        )

    return dados


# ==============================================================================
# GERAÇÃO DE RELATÓRIO EXCEL E PAINEL HTML DE SAÍDA
# ==============================================================================
def gerar_relatorio_excel(
    resultados: list[dict], id_execucao: str, tempo_total_seg: float, caminho_saida: Path
) -> None:
    """Gera a planilha Excel final formatada com status coloridos e aba de resumo."""
    caminho_saida.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()

    # Aba 1: Detalhamento por cliente
    ws = wb.active
    ws.title = "Resultado_Processamento"

    colunas = [
        ("id_cliente", "ID Cliente"),
        ("nome", "Nome"),
        ("cpf_formatado", "CPF Padronizado"),
        ("email", "E-mail"),
        ("cep_formatado", "CEP"),
        ("endereco_completo", "Endereço (ViaCEP)"),
        ("uf", "UF"),
        ("valor_pedido", "Valor Pedido (R$)"),
        ("valor_frete", "Frete (R$)"),
        ("valor_total", "Total (R$)"),
        ("status", "Status"),
        ("mensagem_detalhe", "Detalhe / Regra Aplicada"),
        ("data_hora", "Data/Hora"),
    ]

    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=10)
    sucesso_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    sucesso_font = Font(color="166534", bold=True)
    erro_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    erro_font = Font(color="991B1B", bold=True)
    borda = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB"),
    )

    ws.append([titulo for _, titulo in colunas])
    for col_idx in range(1, len(colunas) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = borda
    ws.row_dimensions[1].height = 24

    for row_idx, item in enumerate(resultados, start=2):
        for col_idx, (chave, _) in enumerate(colunas, start=1):
            valor = item.get(chave)
            cell = ws.cell(row=row_idx, column=col_idx, value=valor)
            cell.border = borda
            if chave in ("valor_pedido", "valor_frete", "valor_total") and valor is not None:
                cell.number_format = "R$ #,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif chave in ("id_cliente", "cpf_formatado", "cep_formatado", "uf", "data_hora"):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif chave == "status":
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if valor == "SUCESSO":
                    cell.fill = sucesso_fill
                    cell.font = sucesso_font
                else:
                    cell.fill = erro_fill
                    cell.font = erro_font
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    larguras = {
        "A": 12,
        "B": 20,
        "C": 17,
        "D": 28,
        "E": 13,
        "F": 44,
        "G": 7,
        "H": 17,
        "I": 14,
        "J": 15,
        "K": 14,
        "L": 54,
        "M": 20,
    }
    for col, largura in larguras.items():
        ws.column_dimensions[col].width = largura

    # Aba 2: Resumo Executivo
    ws_resumo = wb.create_sheet(title="Resumo_Executivo")
    total = len(resultados)
    sucessos = sum(1 for r in resultados if r["status"] == "SUCESSO")
    insucessos = total - sucessos
    soma_total = sum(r["valor_total"] or 0.0 for r in resultados if r["status"] == "SUCESSO")

    linhas_resumo = [
        ("Indicador", "Valor"),
        ("ID da Execução", id_execucao),
        ("Data/Hora do Processamento", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Total de Registros Lidos", total),
        ("Processados com SUCESSO", sucessos),
        ("Bloqueados por Regra de Negócio (INSUCESSO)", insucessos),
        ("Taxa de Sucesso (%)", round((sucessos / total) * 100, 1) if total else 0),
        ("Valor Total Faturado (Pedidos + Fretes)", soma_total),
        ("Tempo Total de Execução do Robô (s)", round(tempo_total_seg, 2)),
    ]

    for idx, (indicador, valor) in enumerate(linhas_resumo, start=1):
        c1 = ws_resumo.cell(row=idx, column=1, value=indicador)
        c2 = ws_resumo.cell(row=idx, column=2, value=valor)
        c1.border = borda
        c2.border = borda
        if idx == 1:
            c1.fill = header_fill
            c1.font = header_font
            c2.fill = header_fill
            c2.font = header_font
        elif indicador == "Valor Total Faturado (Pedidos + Fretes)":
            c2.number_format = "R$ #,##0.00"

    ws_resumo.column_dimensions["A"].width = 45
    ws_resumo.column_dimensions["B"].width = 25

    wb.save(caminho_saida)


def gerar_painel_resultado_html(
    resultados: list[dict], id_execucao: str, tempo_total_seg: float, caminho_html: Path
) -> None:
    """Atualiza o painel HTML da Intranet com o resultado da última execução do RPA."""
    total = len(resultados)
    sucessos = sum(1 for r in resultados if r["status"] == "SUCESSO")
    insucessos = total - sucessos

    linhas_html = []
    for r in resultados:
        classe_status = "badge-ok" if r["status"] == "SUCESSO" else "badge-erro"
        total_fmt = f"R$ {r['valor_total']:.2f}" if r["valor_total"] is not None else "-"
        frete_fmt = f"R$ {r['valor_frete']:.2f}" if r["valor_frete"] is not None else "-"
        linhas_html.append(
            f"""
            <tr>
                <td><strong>{r['id_cliente']}</strong></td>
                <td>{r['nome']}</td>
                <td>{r['cpf_formatado'] or '-'}</td>
                <td>{r['cep_formatado'] or '-'}</td>
                <td>{r['uf'] or '-'}</td>
                <td>{frete_fmt}</td>
                <td>{total_fmt}</td>
                <td><span class="{classe_status}">{r['status']}</span></td>
                <td>{r['mensagem_detalhe']}</td>
            </tr>
            """
        )

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>Painel RPA LogiExpress - Resultado da Execução</title>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #f1f5f9; margin: 0; padding: 28px; color: #0f172a; }}
        .wrap {{ max-width: 1180px; margin: 0 auto; background: #fff; padding: 24px; border-radius: 12px; box-shadow: 0 4px 16px rgba(15,23,42,.08); }}
        h1 {{ margin: 0 0 6px; color: #1e3a8a; font-size: 24px; }}
        .meta {{ color: #475569; font-size: 13px; margin-bottom: 18px; }}
        .cards {{ display: flex; gap: 16px; margin-bottom: 22px; }}
        .card {{ flex: 1; padding: 14px 18px; border-radius: 8px; background: #f8fafc; border-left: 5px solid #1d4ed8; }}
        .card.ok {{ border-left-color: #16a34a; background: #f0fdf4; }}
        .card.err {{ border-left-color: #dc2626; background: #fef2f2; }}
        .card .num {{ font-size: 26px; font-weight: 700; }}
        .card .lbl {{ font-size: 12px; color: #475569; text-transform: uppercase; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
        th {{ background: #1e3a8a; color: #fff; padding: 10px; text-align: left; }}
        td {{ padding: 9px 10px; border-bottom: 1px solid #e2e8f0; }}
        .badge-ok {{ background: #dcfce7; color: #166534; padding: 3px 8px; border-radius: 99px; font-weight: 700; font-size: 11px; }}
        .badge-erro {{ background: #fee2e2; color: #991b1b; padding: 3px 8px; border-radius: 99px; font-weight: 700; font-size: 11px; }}
    </style>
</head>
<body>
    <div class="wrap">
        <h1>Painel de Auditoria — RPA LogiExpress Brasil</h1>
        <div class="meta">Execução: <strong>{id_execucao}</strong> | Tempo total: <strong>{tempo_total_seg:.2f}s</strong> | Banco SQLite sincronizado</div>
        <div class="cards">
            <div class="card"><div class="num">{total}</div><div class="lbl">Total Lidos</div></div>
            <div class="card ok"><div class="num">{sucessos}</div><div class="lbl">Sucesso</div></div>
            <div class="card err"><div class="num">{insucessos}</div><div class="lbl">Insucesso (Regra de Negócio)</div></div>
        </div>
        <table>
            <thead>
                <tr>
                    <th>ID</th><th>Nome</th><th>CPF</th><th>CEP</th><th>UF</th>
                    <th>Frete</th><th>Total</th><th>Status</th><th>Detalhe do Processamento</th>
                </tr>
            </thead>
            <tbody>
                {''.join(linhas_html)}
            </tbody>
        </table>
    </div>
</body>
</html>
"""
    caminho_html.write_text(html, encoding="utf-8")


# ==============================================================================
# FLUXO PRINCIPAL DO ROBÔ RPA
# ==============================================================================
def executar_rpa() -> None:
    """Executa o pipeline completo de automação do RPA LogiExpress."""
    inicio = time.perf_counter()
    logger = configurar_logger()
    id_execucao = f"EXEC-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:4].upper()}"

    logger.info("=================================================================")
    logger.info("INICIANDO ROBÔ RPA LOGIEXPRESS | Execução: %s", id_execucao)
    logger.info("=================================================================")

    if not ARQUIVO_ENTRADA.exists():
        logger.error("Planilha de entrada não encontrada: %s", ARQUIVO_ENTRADA)
        return
    if not ARQUIVO_PORTAL_HTML.exists():
        logger.error("Portal HTML de fretes não encontrado: %s", ARQUIVO_PORTAL_HTML)
        return

    # 1. Inicializar Banco de Dados SQLite
    conn = inicializar_banco(ARQUIVO_BANCO)
    logger.info("Conexão com banco SQLite inicializada: %s", ARQUIVO_BANCO.name)

    # 2. Ler tabela oficial de fretes na Intranet via BeautifulSoup
    tabela_fretes = carregar_tabela_frete_html(ARQUIVO_PORTAL_HTML, logger)

    # 3. Ler planilha de entrada garantindo preservação de zeros à esquerda
    df_clientes = pd.read_excel(
        ARQUIVO_ENTRADA,
        dtype={"ID_Cliente": str, "Nome": str, "CPF": str, "Email": str, "CEP": str},
    )
    logger.info("Planilha '%s' carregada: %d registros pendentes.", ARQUIVO_ENTRADA.name, len(df_clientes))

    resultados: list[dict] = []

    for indice, linha in df_clientes.iterrows():
        id_cliente = str(linha.get("ID_Cliente", f"REG-{indice+1}")).strip()
        nome = str(linha.get("Nome", "")).strip()
        cpf_bruto = str(linha.get("CPF", "")).strip()
        email_bruto = str(linha.get("Email", "")).strip()
        cep_bruto = str(linha.get("CEP", "")).strip()
        valor_pedido_bruto = linha.get("Valor_Pedido", 0)
        agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        logger.info("-----------------------------------------------------------------")
        logger.info("[%s] Iniciando processamento de '%s'...", id_cliente, nome)

        # Estrutura padrão do registro para auditoria
        registro = {
            "id_execucao": id_execucao,
            "id_cliente": id_cliente,
            "nome": nome,
            "cpf_formatado": cpf_bruto,
            "email": email_bruto,
            "cep_formatado": cep_bruto,
            "endereco_completo": None,
            "uf": None,
            "valor_pedido": None,
            "valor_frete": None,
            "valor_total": None,
            "status": "INSUCESSO",
            "mensagem_detalhe": "",
            "data_hora": agora,
        }

        try:
            # Passo A: Validar valor do pedido (> 0)
            valor_pedido = float(valor_pedido_bruto)
            registro["valor_pedido"] = round(valor_pedido, 2)
            if valor_pedido <= 0:
                raise ErroRegraNegocio(
                    f"Valor do pedido inválido (R$ {valor_pedido:.2f}): deve ser maior que zero."
                )

            # Passo B: Validar e padronizar CPF, E-mail e CEP usando Regex
            cpf_formatado = validar_e_formatar_cpf(cpf_bruto)
            registro["cpf_formatado"] = cpf_formatado
            logger.info("[%s] CPF validado e padronizado via Regex: %s", id_cliente, cpf_formatado)

            email_valido = validar_email(email_bruto)
            registro["email"] = email_valido
            logger.info("[%s] E-mail validado via Regex: %s", id_cliente, email_valido)

            cep_limpo, cep_formatado = validar_e_formatar_cep(cep_bruto)
            registro["cep_formatado"] = cep_formatado
            logger.info("[%s] CEP validado e formatado via Regex: %s", id_cliente, cep_formatado)

            # Passo C: Consultar endereço completo na API ViaCEP
            logger.info("[%s] Consultando API ViaCEP para CEP %s...", id_cliente, cep_limpo)
            dados_cep = consultar_api_viacep(cep_limpo, logger)

            uf = dados_cep.get("uf", "").strip().upper()
            logradouro = dados_cep.get("logradouro", "") or "Sem logradouro"
            bairro = dados_cep.get("bairro", "") or "Centro"
            localidade = dados_cep.get("localidade", "")
            endereco_completo = f"{logradouro}, {bairro} - {localidade}/{uf}"

            registro["uf"] = uf
            registro["endereco_completo"] = endereco_completo
            logger.info("[%s] Endereço obtido da API: %s", id_cliente, endereco_completo)

            # Passo D: Buscar frete na tabela extraída via BeautifulSoup e calcular total
            valor_frete = tabela_fretes.get(uf, tabela_fretes.get("OUTROS", 65.0))
            valor_total = round(valor_pedido + valor_frete, 2)

            registro["valor_frete"] = round(valor_frete, 2)
            registro["valor_total"] = valor_total
            registro["status"] = "SUCESSO"
            registro["mensagem_detalhe"] = (
                f"Processado com sucesso. Frete {uf}: R$ {valor_frete:.2f} | Total: R$ {valor_total:.2f}"
            )

            logger.info(
                "[%s] SUCESSO! Frete %s = R$ %.2f | Valor Total = R$ %.2f",
                id_cliente,
                uf,
                valor_frete,
                valor_total,
            )

        except ErroRegraNegocio as erro_negocio:
            # Captura violações de regra de negócio sem interromper os demais clientes
            registro["status"] = "INSUCESSO"
            registro["mensagem_detalhe"] = f"Regra de Negócio: {erro_negocio}"
            logger.error("[%s] INSUCESSO (Regra de Negócio): %s", id_cliente, erro_negocio)

        except Exception as erro_inesperado:
            # Captura falhas técnicas imprevistas garantindo resiliência do robô
            registro["status"] = "INSUCESSO"
            registro["mensagem_detalhe"] = f"Erro Técnico Inesperado: {erro_inesperado}"
            logger.exception("[%s] INSUCESSO (Erro Técnico): %s", id_cliente, erro_inesperado)

        # Passo E: Persistir resultado (sucesso ou insucesso) no SQLite para auditoria
        salvar_registro_banco(conn, registro)
        resultados.append(registro)

    conn.close()

    tempo_total = time.perf_counter() - inicio
    gerar_relatorio_excel(resultados, id_execucao, tempo_total, ARQUIVO_SAIDA)
    gerar_painel_resultado_html(resultados, id_execucao, tempo_total, ARQUIVO_PAINEL_HTML)

    qtd_sucesso = sum(1 for r in resultados if r["status"] == "SUCESSO")
    qtd_erro = len(resultados) - qtd_sucesso

    logger.info("=================================================================")
    logger.info(
        "EXECUÇÃO FINALIZADA em %.2fs | Total: %d | Sucesso: %d | Insucesso: %d",
        tempo_total,
        len(resultados),
        qtd_sucesso,
        qtd_erro,
    )
    logger.info("Relatório Excel gerado em: %s", ARQUIVO_SAIDA)
    logger.info("Painel HTML atualizado em: %s", ARQUIVO_PAINEL_HTML)
    logger.info("=================================================================")


if __name__ == "__main__":
    executar_rpa()
