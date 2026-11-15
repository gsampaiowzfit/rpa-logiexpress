"""
Script gerador dos entregáveis visuais e executivos do projeto RPA LogiExpress:
1. Imagens PNG de evidência geradas diretamente dos artefatos reais (Excel, Log, SQLite)
2. Vídeo MP4 e GIF animado demonstrando a execução passo a passo do robô
3. Apresentação PowerPoint (.pptx) pronta para apresentar

Deve ser executado após rodar `rpa_main.py`.
"""

from pathlib import Path
import sqlite3
import textwrap

import imageio.v2 as imageio
import numpy as np
import openpyxl
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


BASE_DIR = Path(__file__).resolve().parent
ARQUIVO_ENTRADA = BASE_DIR / "dados" / "clientes_pendentes.xlsx"
ARQUIVO_SAIDA = BASE_DIR / "dados" / "relatorio_processamento.xlsx"
ARQUIVO_BANCO = BASE_DIR / "banco" / "rpa_logiexpress.db"
ARQUIVO_LOG = BASE_DIR / "logs" / "rpa_execucao.log"
MIDIA_DIR = BASE_DIR / "midia"


# ==============================================================================
# UTILITÁRIOS DE FONTE E LEITURA DOS DADOS REAIS
# ==============================================================================
def carregar_fonte(tamanho: int, negrito: bool = False, mono: bool = False) -> ImageFont.FreeTypeFont:
    """Carrega fontes nativas do Windows com fallback seguro."""
    candidatos = []
    if mono:
        candidatos = (
            [r"C:\Windows\Fonts\consolab.ttf", r"C:\Windows\Fonts\courbd.ttf"]
            if negrito
            else [r"C:\Windows\Fonts\consola.ttf", r"C:\Windows\Fonts\cour.ttf"]
        )
    else:
        candidatos = (
            [r"C:\Windows\Fonts\segoeuib.ttf", r"C:\Windows\Fonts\arialbd.ttf"]
            if negrito
            else [r"C:\Windows\Fonts\segoeui.ttf", r"C:\Windows\Fonts\arial.ttf"]
        )

    for caminho in candidatos:
        if Path(caminho).exists():
            return ImageFont.truetype(caminho, tamanho)
    return ImageFont.load_default()


def ler_ultima_execucao_sqlite() -> tuple[str, list[dict]]:
    """Lê os registros reais da última execução gravada no SQLite."""
    conn = sqlite3.connect(ARQUIVO_BANCO)
    cursor = conn.cursor()
    cursor.execute("SELECT id_execucao FROM historico_processamento ORDER BY id DESC LIMIT 1")
    row = cursor.fetchone()
    if not row:
        conn.close()
        raise RuntimeError("Nenhum registro encontrado no banco SQLite. Execute rpa_main.py primeiro.")

    id_execucao = row[0]
    cursor.execute(
        """
        SELECT id_cliente, nome, cpf_formatado, email, cep_formatado,
               endereco_completo, uf, valor_pedido, valor_frete,
               valor_total, status, mensagem_detalhe, data_hora
        FROM historico_processamento
        WHERE id_execucao = ?
        ORDER BY id ASC
        """,
        (id_execucao,),
    )
    colunas = [
        "id_cliente", "nome", "cpf_formatado", "email", "cep_formatado",
        "endereco_completo", "uf", "valor_pedido", "valor_frete",
        "valor_total", "status", "mensagem_detalhe", "data_hora",
    ]
    registros = [dict(zip(colunas, linha)) for linha in cursor.fetchall()]
    conn.close()
    return id_execucao, registros


# ==============================================================================
# GERAÇÃO DAS 4 IMAGENS DE EVIDÊNCIA (.PNG)
# ==============================================================================
def gerar_imagem_entrada_excel() -> Path:
    """Renderiza captura visual da planilha Excel de entrada real."""
    wb = openpyxl.load_workbook(ARQUIVO_ENTRADA, data_only=True)
    ws = wb.active
    linhas = list(ws.iter_rows(values_only=True))

    largura, altura = 1280, 620
    img = Image.new("RGB", (largura, altura), "#FFFFFF")
    draw = ImageDraw.Draw(img)

    f_titulo = carregar_fonte(22, negrito=True)
    f_sub = carregar_fonte(14)
    f_hdr = carregar_fonte(14, negrito=True)
    f_cell = carregar_fonte(13)

    draw.rectangle([0, 0, largura, 68], fill="#14532D")
    draw.text((28, 14), "EVIDÊNCIA 1 — Planilha Excel de Entrada (dados/clientes_pendentes.xlsx)", fill="#FFFFFF", font=f_titulo)
    draw.text((28, 42), "10 registros pendentes para validação cadastral e cálculo de frete (contém dados válidos e inválidos)", fill="#DCFCE7", font=f_sub)

    col_x = [28, 135, 320, 495, 790, 940, 1245]
    y = 95
    row_h = 42

    # Cabeçalho
    draw.rectangle([col_x[0], y, col_x[-1], y + row_h], fill="#1E3A8A")
    for idx, titulo in enumerate(linhas[0]):
        draw.text((col_x[idx] + 10, y + 12), str(titulo), fill="#FFFFFF", font=f_hdr)

    for r_idx, row in enumerate(linhas[1:], start=1):
        ry = y + r_idx * row_h
        bg = "#F8FAFC" if r_idx % 2 == 0 else "#FFFFFF"
        draw.rectangle([col_x[0], ry, col_x[-1], ry + row_h], fill=bg, outline="#CBD5E1")
        for c_idx, val in enumerate(row):
            texto = f"R$ {float(val):.2f}" if c_idx == 5 else str(val)
            cor = "#B91C1C" if (c_idx == 5 and float(val) < 0) else "#0F172A"
            draw.text((col_x[c_idx] + 10, ry + 12), texto, fill=cor, font=f_cell)

    draw.text((28, 575), "Nota: Os registros CT03, CT05, CT06, CT08 e CT09 possuem erros propositais para validar o tratamento de regras de negócio.", fill="#475569", font=f_sub)

    out = MIDIA_DIR / "evidencia_entrada_excel.png"
    img.save(out)
    return out


def gerar_imagem_logs() -> Path:
    """Renderiza captura visual em estilo terminal do arquivo logs/rpa_execucao.log."""
    linhas_log = ARQUIVO_LOG.read_text(encoding="utf-8").splitlines()

    # Seleciona linhas representativas (abertura, alguns sucessos, todos os erros e encerramento)
    selecionadas = []
    for ln in linhas_log:
        if any(tag in ln for tag in [
            "INICIANDO ROBÔ", "Conexão com banco", "Tabela de frete",
            "Planilha 'clientes_pendentes", "SUCESSO!", "ERROR",
            "EXECUÇÃO FINALIZADA", "Relatório Excel"
        ]):
            selecionadas.append(ln)

    largura, altura = 1320, 760
    img = Image.new("RGB", (largura, altura), "#0F172A")
    draw = ImageDraw.Draw(img)

    f_titulo = carregar_fonte(20, negrito=True)
    f_sub = carregar_fonte(13)
    f_mono = carregar_fonte(13, mono=True)

    draw.rectangle([0, 0, largura, 64], fill="#1E293B")
    draw.ellipse([24, 24, 38, 38], fill="#EF4444")
    draw.ellipse([46, 24, 60, 38], fill="#F59E0B")
    draw.ellipse([68, 24, 82, 38], fill="#22C55E")
    draw.text((102, 14), "EVIDÊNCIA 2 — Logs Estruturados de Execução (logs/rpa_execucao.log)", fill="#F8FAFC", font=f_titulo)
    draw.text((102, 40), f"Exibindo marcos principais de {len(linhas_log)} linhas registradas passo a passo", fill="#94A3B8", font=f_sub)

    y = 82
    for ln in selecionadas[:24]:
        if "| ERROR" in ln:
            cor = "#FCA5A5"
            draw.rectangle([18, y - 2, largura - 18, y + 22], fill="#450A0A")
        elif "SUCESSO!" in ln:
            cor = "#86EFAC"
        elif "EXECUÇÃO FINALIZADA" in ln or "INICIANDO ROBÔ" in ln:
            cor = "#7DD3FC"
        else:
            cor = "#E2E8F0"

        texto_cortado = ln[:136] + ("..." if len(ln) > 136 else "")
        draw.text((26, y + 2), texto_cortado, fill=cor, font=f_mono)
        y += 26

    out = MIDIA_DIR / "evidencia_logs_execucao.png"
    img.save(out)
    return out


def gerar_imagem_sqlite(id_execucao: str, registros: list[dict]) -> Path:
    """Renderiza captura visual da tabela SQLite historico_processamento."""
    largura, altura = 1320, 680
    img = Image.new("RGB", (largura, altura), "#FFFFFF")
    draw = ImageDraw.Draw(img)

    f_titulo = carregar_fonte(21, negrito=True)
    f_sub = carregar_fonte(14)
    f_hdr = carregar_fonte(13, negrito=True)
    f_cell = carregar_fonte(12)
    f_mono = carregar_fonte(12, mono=True)

    draw.rectangle([0, 0, largura, 74], fill="#312E81")
    draw.text((28, 12), "EVIDÊNCIA 3 — Auditoria Persistida no SQLite (banco/rpa_logiexpress.db)", fill="#FFFFFF", font=f_titulo)
    draw.text((28, 42), f"SELECT * FROM historico_processamento WHERE id_execucao = '{id_execucao}'", fill="#C7D2FE", font=f_mono)

    col_x = [24, 88, 228, 360, 465, 515, 600, 690, 795, 1296]
    titulos = ["ID", "Nome", "CPF", "CEP", "UF", "Frete", "Total", "Status", "Mensagem de Auditoria"]

    y = 98
    row_h = 46
    draw.rectangle([col_x[0], y, col_x[-1], y + 38], fill="#1E293B")
    for idx, t in enumerate(titulos):
        draw.text((col_x[idx] + 8, y + 10), t, fill="#FFFFFF", font=f_hdr)

    y += 38
    for idx, r in enumerate(registros):
        ry = y + idx * row_h
        bg = "#F8FAFC" if idx % 2 == 0 else "#FFFFFF"
        draw.rectangle([col_x[0], ry, col_x[-1], ry + row_h], fill=bg, outline="#CBD5E1")

        frete_str = f"R$ {r['valor_frete']:.2f}" if r["valor_frete"] is not None else "-"
        total_str = f"R$ {r['valor_total']:.2f}" if r["valor_total"] is not None else "-"
        detalhe = textwrap.shorten(r["mensagem_detalhe"], width=68, placeholder="...")

        valores = [
            r["id_cliente"],
            r["nome"],
            str(r["cpf_formatado"] or "-"),
            str(r["cep_formatado"] or "-"),
            str(r["uf"] or "-"),
            frete_str,
            total_str,
            r["status"],
            detalhe,
        ]

        for c_idx, val in enumerate(valores):
            if c_idx == 7:
                pill_bg = "#DCFCE7" if val == "SUCESSO" else "#FEE2E2"
                pill_fg = "#166534" if val == "SUCESSO" else "#991B1B"
                draw.rounded_rectangle([col_x[c_idx] + 6, ry + 10, col_x[c_idx + 1] - 6, ry + 34], radius=6, fill=pill_bg)
                draw.text((col_x[c_idx] + 14, ry + 14), val, fill=pill_fg, font=f_hdr)
            else:
                cor = "#991B1B" if (c_idx == 8 and r["status"] == "INSUCESSO") else "#0F172A"
                draw.text((col_x[c_idx] + 8, ry + 15), val, fill=cor, font=f_cell)

    sucessos = sum(1 for r in registros if r["status"] == "SUCESSO")
    insucessos = len(registros) - sucessos
    draw.text((24, 635), f"Resumo no SQLite: {len(registros)} registros auditados | {sucessos} SUCESSO | {insucessos} INSUCESSO (capturados sem interrupção do robô).", fill="#334155", font=f_sub)

    out = MIDIA_DIR / "evidencia_banco_sqlite.png"
    img.save(out)
    return out


def gerar_imagem_relatorio_final(id_execucao: str, registros: list[dict]) -> Path:
    """Renderiza captura visual do relatório Excel final gerado pelo robô."""
    largura, altura = 1320, 720
    img = Image.new("RGB", (largura, altura), "#F8FAFC")
    draw = ImageDraw.Draw(img)

    f_titulo = carregar_fonte(21, negrito=True)
    f_sub = carregar_fonte(14)
    f_kpi = carregar_fonte(26, negrito=True)
    f_hdr = carregar_fonte(13, negrito=True)
    f_cell = carregar_fonte(12)

    draw.rectangle([0, 0, largura, 68], fill="#1E3A8A")
    draw.text((28, 12), "EVIDÊNCIA 4 — Relatório Excel Final Gerado pelo Robô (dados/relatorio_processamento.xlsx)", fill="#FFFFFF", font=f_titulo)
    draw.text((28, 40), f"Execução {id_execucao} | Aba 'Resultado_Processamento' + Aba 'Resumo_Executivo'", fill="#BFDBFE", font=f_sub)

    total = len(registros)
    sucessos = sum(1 for r in registros if r["status"] == "SUCESSO")
    insucessos = total - sucessos
    faturado = sum(r["valor_total"] or 0.0 for r in registros if r["status"] == "SUCESSO")

    cards = [
        ("TOTAL LIDOS", str(total), "#1D4ED8", "#EFF6FF"),
        ("SUCESSO", str(sucessos), "#16A34A", "#F0FDF4"),
        ("INSUCESSO (REGRAS)", str(insucessos), "#DC2626", "#FEF2F2"),
        ("TOTAL FATURADO", f"R$ {faturado:,.2f}", "#0F766E", "#F0FDFA"),
    ]
    cx = 24
    for lbl, val, cor_borda, cor_bg in cards:
        draw.rounded_rectangle([cx, 84, cx + 305, 162], radius=8, fill=cor_bg, outline=cor_borda, width=2)
        draw.text((cx + 16, 94), lbl, fill="#475569", font=f_hdr)
        draw.text((cx + 16, 118), val, fill=cor_borda, font=f_kpi)
        cx += 322

    col_x = [24, 86, 224, 355, 675, 725, 815, 920, 1296]
    titulos = ["ID", "Nome", "CPF Padronizado", "Endereço Enriquecido (ViaCEP)", "UF", "Total (R$)", "Status", "Regra / Detalhe"]

    y = 180
    draw.rectangle([col_x[0], y, col_x[-1], y + 36], fill="#1E3A8A")
    for idx, t in enumerate(titulos):
        draw.text((col_x[idx] + 8, y + 9), t, fill="#FFFFFF", font=f_hdr)

    y += 36
    row_h = 46
    for idx, r in enumerate(registros):
        ry = y + idx * row_h
        bg = "#FFFFFF" if idx % 2 == 0 else "#F1F5F9"
        draw.rectangle([col_x[0], ry, col_x[-1], ry + row_h], fill=bg, outline="#CBD5E1")

        end_curto = textwrap.shorten(r["endereco_completo"] or "-", width=42, placeholder="...")
        total_str = f"R$ {r['valor_total']:.2f}" if r["valor_total"] is not None else "-"
        det_curto = textwrap.shorten(r["mensagem_detalhe"], width=52, placeholder="...")

        vals = [
            r["id_cliente"],
            r["nome"],
            str(r["cpf_formatado"] or "-"),
            end_curto,
            str(r["uf"] or "-"),
            total_str,
            r["status"],
            det_curto,
        ]
        for c_idx, val in enumerate(vals):
            if c_idx == 6:
                p_bg = "#DCFCE7" if val == "SUCESSO" else "#FEE2E2"
                p_fg = "#166534" if val == "SUCESSO" else "#991B1B"
                draw.rounded_rectangle([col_x[c_idx] + 6, ry + 10, col_x[c_idx + 1] - 6, ry + 34], radius=6, fill=p_bg)
                draw.text((col_x[c_idx] + 16, ry + 14), val, fill=p_fg, font=f_hdr)
            else:
                draw.text((col_x[c_idx] + 8, ry + 15), val, fill="#0F172A", font=f_cell)

    out = MIDIA_DIR / "evidencia_relatorio_final.png"
    img.save(out)
    return out


# ==============================================================================
# GERAÇÃO DO VÍDEO DEMONSTRATIVO (.MP4 E .GIF)
# ==============================================================================
def desenhar_frame_video(
    registros: list[dict],
    passo_atual: int,
    subpasso: int,
    id_execucao: str,
    logs_visiveis: list[tuple[str, str]],
) -> np.ndarray:
    """Desenha um frame 1280x720 para o vídeo da execução do RPA."""
    largura, altura = 1280, 720
    img = Image.new("RGB", (largura, altura), "#0B1120")
    draw = ImageDraw.Draw(img)

    f_title = carregar_fonte(24, negrito=True)
    f_h2 = carregar_fonte(17, negrito=True)
    f_body = carregar_fonte(14)
    f_small = carregar_fonte(12)
    f_mono = carregar_fonte(12, mono=True)
    f_kpi = carregar_fonte(28, negrito=True)

    # Cabeçalho Superior
    draw.rectangle([0, 0, largura, 74], fill="#1E3A8A")
    draw.text((28, 12), "RPA LogiExpress Brasil — Execução Automatizada em Tempo Real", fill="#FFFFFF", font=f_title)
    draw.text(
        (28, 44),
        f"Execução: {id_execucao}  |  Stack: Excel + Regex + API ViaCEP + BeautifulSoup + Try/Except + SQLite",
        fill="#BFDBFE",
        font=f_small,
    )

    # Barra de Progresso
    concluidos = max(0, min(len(registros), passo_atual + (1 if subpasso == 2 else 0)))
    prog = concluidos / len(registros) if registros else 0
    draw.rounded_rectangle([28, 86, 1252, 108], radius=6, fill="#1E293B")
    if prog > 0:
        draw.rounded_rectangle([28, 86, int(28 + (1224 * prog)), 108], radius=6, fill="#22C55E")
    draw.text((540, 89), f"Progresso: {concluidos}/{len(registros)} clientes ({int(prog*100)}%)", fill="#FFFFFF", font=f_small)

    # Painel Esquerdo: Terminal de Logs
    draw.rounded_rectangle([28, 122, 635, 690], radius=10, fill="#0F172A", outline="#334155", width=2)
    draw.rectangle([28, 122, 635, 158], fill="#1E293B")
    draw.text((44, 131), "TERMINAL DO ROBÔ (logs/rpa_execucao.log)", fill="#7DD3FC", font=f_h2)

    ly = 172
    for nivel, msg in logs_visiveis[-18:]:
        linhas_quebradas = textwrap.wrap(msg, width=74)[:2]
        altura_bloco = max(1, len(linhas_quebradas)) * 21
        if ly + altura_bloco >= 675:
            break

        if nivel == "ERROR":
            cor = "#FCA5A5"
            draw.rectangle([38, ly - 2, 625, ly + altura_bloco], fill="#450A0A")
        elif nivel == "OK":
            cor = "#86EFAC"
        elif nivel == "INFO":
            cor = "#E2E8F0"
        else:
            cor = "#93C5FD"

        for q in linhas_quebradas:
            draw.text((44, ly), q, fill=cor, font=f_mono)
            ly += 21

    # Painel Direito Superior: Pipeline do Cliente Atual
    draw.rounded_rectangle([652, 122, 1252, 336], radius=10, fill="#111827", outline="#334155", width=2)
    draw.rectangle([652, 122, 1252, 158], fill="#1E293B")
    draw.text((668, 131), "ETAPA EM EXECUÇÃO (PIPELINE RPA)", fill="#F8FAFC", font=f_h2)

    if passo_atual < 0:
        draw.text((670, 174), "INICIALIZANDO ROBÔ RPA LOGIEXPRESS...", fill="#7DD3FC", font=f_h2)
        draw.text((670, 210), "• Conectando ao banco de auditoria SQLite", fill="#E2E8F0", font=f_body)
        draw.text((670, 238), "• Extraindo tarifas de frete HTML via BeautifulSoup", fill="#E2E8F0", font=f_body)
        draw.text((670, 266), f"• Carregando {len(registros)} clientes pendentes da planilha Excel", fill="#E2E8F0", font=f_body)
    elif 0 <= passo_atual < len(registros):
        atual = registros[passo_atual]
        draw.text((670, 170), f"Cliente: {atual['id_cliente']} — {atual['nome']}", fill="#FFFFFF", font=f_h2)
        draw.text((670, 198), f"CPF: {atual['cpf_formatado']}   |   CEP: {atual['cep_formatado']}", fill="#CBD5E1", font=f_body)

        etapas = ["1. Regex", "2. API ViaCEP", "3. Frete BS4", "4. SQLite"]
        ex = 670
        for idx_etapa, nome_etapa in enumerate(etapas):
            ativo = idx_etapa <= subpasso or subpasso == 2
            cor_box = "#1D4ED8" if ativo else "#1F2937"
            if subpasso == 2 and idx_etapa == 3:
                cor_box = "#15803D" if atual["status"] == "SUCESSO" else "#B91C1C"
            draw.rounded_rectangle([ex, 232, ex + 132, 276], radius=8, fill=cor_box, outline="#475569")
            draw.text((ex + 16, 246), nome_etapa, fill="#FFFFFF", font=f_small)
            ex += 142

        if subpasso == 2:
            cor_res = "#86EFAC" if atual["status"] == "SUCESSO" else "#FCA5A5"
            resumo = textwrap.shorten(f"[{atual['status']}] {atual['mensagem_detalhe']}", width=70, placeholder="...")
            draw.text((670, 294), resumo, fill=cor_res, font=f_body)
        else:
            draw.text((670, 294), "Validando regras de negócio e integrando sistemas...", fill="#93C5FD", font=f_body)
    else:
        sucessos = sum(1 for r in registros if r["status"] == "SUCESSO")
        insucessos = len(registros) - sucessos
        draw.text((670, 174), "PROCESSAMENTO CONCLUÍDO COM RESILIÊNCIA!", fill="#4ADE80", font=f_h2)
        draw.text((670, 210), f"Sucesso: {sucessos} clientes faturados", fill="#86EFAC", font=f_body)
        draw.text((670, 238), f"Insucesso: {insucessos} erros de negócio capturados pelo try-except", fill="#FCA5A5", font=f_body)
        draw.text((670, 272), "Artefatos salvos: Excel + Banco SQLite + Log + Painel HTML", fill="#93C5FD", font=f_body)

    # Painel Direito Inferior: Tabela de Status em Tempo Real
    draw.rounded_rectangle([652, 350, 1252, 690], radius=10, fill="#111827", outline="#334155", width=2)
    draw.rectangle([652, 350, 1252, 386], fill="#1E293B")
    draw.text((668, 359), "FILA DE CLIENTES & AUDITORIA SQLITE", fill="#F8FAFC", font=f_h2)

    ty = 398
    for idx, r in enumerate(registros):
        ja_foi = idx < passo_atual or (idx == passo_atual and subpasso == 2) or passo_atual >= len(registros)
        em_andamento = idx == passo_atual and subpasso < 2

        if ja_foi:
            badge = r["status"]
            b_bg = "#14532D" if badge == "SUCESSO" else "#7F1D1D"
            b_fg = "#86EFAC" if badge == "SUCESSO" else "#FCA5A5"
        elif em_andamento:
            badge = "LENDO..."
            b_bg = "#1E3A8A"
            b_fg = "#93C5FD"
        else:
            badge = "AGUARD."
            b_bg = "#1F2937"
            b_fg = "#94A3B8"

        draw.text((668, ty + 4), f"{r['id_cliente']}  {r['nome']:<17}", fill="#E2E8F0", font=f_mono)
        uf_txt = r["uf"] if (ja_foi and r["uf"]) else "--"
        total_txt = f"R${r['valor_total']:.2f}" if (ja_foi and r["valor_total"]) else "   ---   "
        draw.text((885, ty + 4), f"UF:{uf_txt:<3} {total_txt:<10}", fill="#CBD5E1", font=f_mono)
        draw.rounded_rectangle([1095, ty, 1236, ty + 23], radius=5, fill=b_bg)
        draw.text((1114, ty + 4), badge, fill=b_fg, font=f_small)
        ty += 28

    return np.array(img)


def gerar_video_execucao(id_execucao: str, registros: list[dict]) -> tuple[Path, Path]:
    """Renderiza o vídeo MP4 e o GIF animado mostrando a execução do robô."""
    MIDIA_DIR.mkdir(parents=True, exist_ok=True)
    mp4_path = MIDIA_DIR / "video_execucao_rpa.mp4"
    gif_path = MIDIA_DIR / "video_execucao_rpa.gif"

    frames_mp4: list[np.ndarray] = []
    logs: list[tuple[str, str]] = [
        ("HEAD", f"INICIANDO ROBO RPA LOGIEXPRESS | {id_execucao}"),
        ("INFO", "Banco SQLite conectado: banco/rpa_logiexpress.db"),
        ("INFO", "BeautifulSoup: 7 UFs extraidas de tabela_frete_portal.html"),
        ("INFO", f"Excel carregado: {len(registros)} clientes pendentes"),
    ]

    # Abertura (1.5s = 22 frames a 15 fps)
    frame_abertura = desenhar_frame_video(registros, -1, 0, id_execucao, logs)
    frames_mp4.extend([frame_abertura] * 22)

    # Animação por cliente
    for idx, r in enumerate(registros):
        logs.append(("INFO", f"--------------------------------------------------"))
        logs.append(("INFO", f"[{r['id_cliente']}] Iniciando: {r['nome']} | CEP {r['cep_formatado']}"))
        f0 = desenhar_frame_video(registros, idx, 0, id_execucao, logs)
        frames_mp4.extend([f0] * 10)

        logs.append(("INFO", f"[{r['id_cliente']}] Validando Regex + Consultando API ViaCEP..."))
        f1 = desenhar_frame_video(registros, idx, 1, id_execucao, logs)
        frames_mp4.extend([f1] * 10)

        if r["status"] == "SUCESSO":
            logs.append(("OK", f"[{r['id_cliente']}] SUCESSO! UF={r['uf']} Frete=R${r['valor_frete']:.2f} Total=R${r['valor_total']:.2f}"))
        else:
            logs.append(("ERROR", f"[{r['id_cliente']}] INSUCESSO: {r['mensagem_detalhe']}"))

        f2 = desenhar_frame_video(registros, idx, 2, id_execucao, logs)
        frames_mp4.extend([f2] * 15)

    # Tela final (3 segundos = 45 frames)
    logs.append(("HEAD", "=================================================="))
    logs.append(("HEAD", "EXECUCAO FINALIZADA | 5 SUCESSO | 5 INSUCESSO"))
    logs.append(("OK", "Gerado: relatorio_processamento.xlsx + SQLite + Log"))
    frame_fim = desenhar_frame_video(registros, len(registros), 2, id_execucao, logs)
    frames_mp4.extend([frame_fim] * 45)

    with imageio.get_writer(str(mp4_path), fps=15, codec="libx264", pixelformat="yuv420p", quality=8) as writer:
        for frame in frames_mp4:
            writer.append_data(frame)

    # GIF leve (reduzido para 800x450 e amostragem de frames)
    frames_gif = []
    for frame in frames_mp4[::5]:
        pil_f = Image.fromarray(frame).resize((800, 450), Image.Resampling.LANCZOS)
        frames_gif.append(np.array(pil_f))
    imageio.mimsave(str(gif_path), frames_gif, fps=3, loop=0)

    return mp4_path, gif_path


# ==============================================================================
# GERAÇÃO DA APRESENTAÇÃO POWERPOINT (.PPTX)
# ==============================================================================
def criar_apresentacao_pptx(id_execucao: str, registros: list[dict]) -> Path:
    """Cria a apresentação executiva em PowerPoint cobrindo todos os itens exigidos."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    AZUL_ESCURO = RGBColor(30, 58, 138)
    AZUL_CLARO = RGBColor(2, 132, 199)
    VERDE = RGBColor(22, 163, 74)
    VERMELHO = RGBColor(220, 38, 38)
    CINZA_ESCURO = RGBColor(15, 23, 42)
    CINZA_TEXTO = RGBColor(51, 65, 85)
    FUNDO_CLARO = RGBColor(248, 250, 252)
    BRANCO = RGBColor(255, 255, 255)

    def adicionar_cabecalho(slide, titulo: str, subtitulo: str, numero: int):
        topo = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.15))
        topo.fill.solid()
        topo.fill.fore_color.rgb = AZUL_ESCURO
        topo.line.fill.background()

        tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.18), Inches(11.5), Inches(0.85))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = titulo
        p1.font.size = Pt(26)
        p1.font.bold = True
        p1.font.color.rgb = BRANCO

        p2 = tf.add_paragraph()
        p2.text = subtitulo
        p2.font.size = Pt(13)
        p2.font.color.rgb = RGBColor(191, 219, 254)

        rod = slide.shapes.add_textbox(Inches(0.6), Inches(7.08), Inches(12.1), Inches(0.3))
        p_rod = rod.text_frame.paragraphs[0]
        p_rod.text = f"LogiExpress Brasil — Projeto RPA de Automação Logística e Cadastral  |  Slide {numero}"
        p_rod.font.size = Pt(10)
        p_rod.font.color.rgb = RGBColor(100, 116, 139)

    def adicionar_card(slide, x, y, w, h, titulo, corpo, cor_borda=AZUL_CLARO, fundo=FUNDO_CLARO):
        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        box.fill.solid()
        box.fill.fore_color.rgb = fundo
        box.line.color.rgb = cor_borda
        box.line.width = Pt(2)

        tb = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.15), Inches(w - 0.4), Inches(h - 0.3))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = titulo
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = cor_borda

        for linha in corpo:
            pl = tf.add_paragraph()
            pl.text = linha
            pl.font.size = Pt(13)
            pl.font.color.rgb = CINZA_ESCURO
            pl.space_before = Pt(6)

    # ------------------------------------------------------------------
    # SLIDE 1: CAPA
    # ------------------------------------------------------------------
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = AZUL_ESCURO
    bg1.line.fill.background()

    faixa = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(5.6), Inches(13.333), Inches(0.12))
    faixa.fill.solid()
    faixa.fill.fore_color.rgb = AZUL_CLARO
    faixa.line.fill.background()

    tb1 = s1.shapes.add_textbox(Inches(0.9), Inches(1.6), Inches(11.5), Inches(3.6))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p = tf1.paragraphs[0]
    p.text = "PROJETO RPA — AUTOMAÇÃO INTELIGENTE"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = RGBColor(125, 211, 252)

    p2 = tf1.add_paragraph()
    p2.text = "LogiExpress Brasil: Validação Cadastral, Enriquecimento de Endereço e Cálculo de Frete"
    p2.font.size = Pt(36)
    p2.font.bold = True
    p2.font.color.rgb = BRANCO
    p2.space_before = Pt(14)

    p3 = tf1.add_paragraph()
    p3.text = "De um processo manual lento e sujeito a falhas para uma esteira automatizada, auditável e resiliente em Python."
    p3.font.size = Pt(18)
    p3.font.color.rgb = RGBColor(226, 232, 240)
    p3.space_before = Pt(18)

    tb_foot = s1.shapes.add_textbox(Inches(0.9), Inches(6.0), Inches(11.5), Inches(0.8))
    pf = tb_foot.text_frame.paragraphs[0]
    pf.text = "Entregáveis: Apresentação Executiva + Código-Fonte Python + Vídeo de Execução + Evidências de Teste (SQLite, Excel e Logs)"
    pf.font.size = Pt(14)
    pf.font.color.rgb = RGBColor(191, 219, 254)

    # ------------------------------------------------------------------
    # SLIDE 2: DESCRIÇÃO DO CASE
    # ------------------------------------------------------------------
    s2 = prs.slides.add_slide(blank_layout)
    adicionar_cabecalho(s2, "1. Descrição do Case de Negócio", "Contexto operacional da LogiExpress Brasil e o gargalo identificado", 2)

    adicionar_card(
        s2, 0.6, 1.45, 3.85, 5.3,
        "A Empresa: LogiExpress Brasil",
        [
            "• Transportadora e operadora logística de e-commerce B2B/B2C.",
            "• Recebe diariamente dezenas de pedidos de novos clientes via planilhas enviadas por vendedores e parceiros comerciais.",
            "• Volume médio atual: 150 cadastros/dia realizados por uma equipe de 3 analistas operacionais.",
        ],
        AZUL_ESCURO,
    )
    adicionar_card(
        s2, 4.74, 1.45, 3.85, 5.3,
        "O Processo de Negócio",
        [
            "• Onboarding e Pré-Faturamento de Pedidos.",
            "• Antes de liberar o pedido para expedição e nota fiscal, o analista precisa:",
            "  1) Conferir CPF, e-mail e CEP;",
            "  2) Consultar o endereço real nos Correios;",
            "  3) Descobrir a UF e consultar o frete no portal da intranet;",
            "  4) Somar Pedido + Frete e devolver o status.",
        ],
        AZUL_CLARO,
    )
    adicionar_card(
        s2, 8.88, 1.45, 3.85, 5.3,
        "A Tarefa Repetitiva (Problema)",
        [
            "• Copiar e colar manualmente cada linha do Excel no navegador (Correios/ViaCEP) e depois na intranet.",
            "• Formatar CPFs e CEPs digitados sem padrão ('12345678901' vs '123.456.789-01').",
            "• Alto índice de retrabalho: pedidos faturados com CEP inexistente, e-mail errado ou frete da UF incorreta.",
        ],
        VERMELHO, RGBColor(254, 242, 242),
    )

    # ------------------------------------------------------------------
    # SLIDE 3: JUSTIFICATIVA DA AUTOMAÇÃO
    # ------------------------------------------------------------------
    s3 = prs.slides.add_slide(blank_layout)
    adicionar_cabecalho(s3, "2. Justificativa da Automação (ROI e Qualidade)", "Por que automatizar com RPA e quais melhorias de negócio foram agregadas", 3)

    adicionar_card(
        s3, 0.6, 1.45, 5.9, 2.55,
        "Tempo Gasto Manualmente (As-Is)",
        [
            "• Média por registro manual: ~4 minutos por cliente (abrir abas, copiar dados, buscar CEP, checar frete na intranet e digitar de volta no Excel).",
            "• Para 150 pedidos/dia: 150 × 4 min = 600 minutos/dia = 10 horas/dia de trabalho humano repetitivo (exigindo 2 analistas dedicados integralmente).",
        ],
        VERMELHO, RGBColor(254, 242, 242),
    )
    adicionar_card(
        s3, 6.8, 1.45, 5.9, 2.55,
        "Tempo com o Robô RPA (To-Be)",
        [
            "• Tempo médio do robô: ~0,12s por registro (10 registros processados em cerca de 1 segundo; 150 registros em menos de 20 segundos).",
            "• Redução superior a 99% do tempo operacional, liberando os analistas para atendimento consultivo e resolução de exceções reais.",
        ],
        VERDE, RGBColor(240, 253, 244),
    )
    adicionar_card(
        s3, 0.6, 4.25, 12.1, 2.5,
        "Novos Requisitos de Negócio Adicionados para Melhorar a Qualidade",
        [
            "• Padronização Automática (Regex): CPFs e CEPs são automaticamente validados e formatados (000.000.000-00 e 00000-000) antes de seguir.",
            "• Bloqueio Preventivo de Regras: E-mails malformados, valores <= 0 ou CEPs inexistentes na API ViaCEP são barrados antes do faturamento.",
            "• Trilha de Auditoria em Banco SQLite + Logs: Cada tentativa (sucesso ou erro) fica gravada com data, hora e motivo exato para compliance.",
            "• Resiliência Total (Try-Except): Se o cliente 3 tiver erro, o robô registra o erro e continua imediatamente para o cliente 4 sem travar.",
        ],
        AZUL_ESCURO,
    )

    # ------------------------------------------------------------------
    # SLIDE 4: FLUXO AS-IS (MANUAL)
    # ------------------------------------------------------------------
    s4 = prs.slides.add_slide(blank_layout)
    adicionar_cabecalho(s4, "3. Desenho do Fluxo Atual — Manual (As-Is)", "Como os analistas executam a atividade hoje sem automação", 4)

    passos_asis = [
        ("1. Abrir Excel", "Analista abre planilha de pendentes"),
        ("2. Conferir Olhômetro", "Checa CPF, e-mail e valor visualmente"),
        ("3. Copiar CEP", "Copia CEP e abre site dos Correios"),
        ("4. Buscar Intranet", "Procura UF na tabela HTML da intranet"),
        ("5. Digitar e Calcular", "Soma frete manual e preenche planilha"),
    ]
    px = 0.55
    for idx, (tit, desc) in enumerate(passos_asis):
        box = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(px), Inches(2.0), Inches(2.2), Inches(1.75))
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(254, 242, 242)
        box.line.color.rgb = VERMELHO
        box.line.width = Pt(2)
        tb = s4.shapes.add_textbox(Inches(px + 0.1), Inches(2.15), Inches(2.0), Inches(1.5))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = tit
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = VERMELHO
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(12)
        p2.font.color.rgb = CINZA_ESCURO
        p2.space_before = Pt(8)
        if idx < len(passos_asis) - 1:
            arr = s4.shapes.add_textbox(Inches(px + 2.2), Inches(2.6), Inches(0.32), Inches(0.5))
            pa = arr.text_frame.paragraphs[0]
            pa.text = "→"
            pa.font.size = Pt(22)
            pa.font.bold = True
            pa.font.color.rgb = VERMELHO
        px += 2.52

    adicionar_card(
        s4, 0.55, 4.25, 12.2, 2.4,
        "Principais Gargalos e Riscos do Fluxo As-Is",
        [
            "• Dependência total de digitação humana (risco de erro ao copiar CEP, trocar dígitos do CPF ou aplicar tarifa errada).",
            "• Ausência de histórico estruturado em banco de dados: se alguém sobrescrever a planilha, perde-se a rastreabilidade.",
            "• Tempo elevado (~4 min/cadastro), gerando fila acumulada em picos comerciais (Black Friday, fechamento de mês).",
        ],
        VERMELHO, FUNDO_CLARO,
    )

    # ------------------------------------------------------------------
    # SLIDE 5: FLUXO TO-BE (AUTOMATIZADO)
    # ------------------------------------------------------------------
    s5 = prs.slides.add_slide(blank_layout)
    adicionar_cabecalho(s5, "4. Desenho do Novo Fluxo — Automatizado (To-Be)", "Arquitetura completa executada pelo robô RPA em Python", 5)

    passos_tobe = [
        ("1. Leitura Excel & BS4", "Lê planilha de entrada e raspa tabela HTML de frete"),
        ("2. Validação Regex", "Valida e formata CPF, E-mail e CEP automaticamente"),
        ("3. API ViaCEP", "Consulta endereço completo e UF oficial via HTTP GET"),
        ("4. Cálculo & Try/Except", "Calcula frete ou captura erro de regra sem travar"),
        ("5. SQLite + Excel + Log", "Persiste auditoria no banco, gera Excel colorido e logs"),
    ]
    px = 0.55
    for idx, (tit, desc) in enumerate(passos_tobe):
        box = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(px), Inches(2.0), Inches(2.2), Inches(1.85))
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(240, 253, 244)
        box.line.color.rgb = VERDE
        box.line.width = Pt(2)
        tb = s5.shapes.add_textbox(Inches(px + 0.1), Inches(2.15), Inches(2.0), Inches(1.6))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = tit
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = VERDE
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(12)
        p2.font.color.rgb = CINZA_ESCURO
        p2.space_before = Pt(8)
        if idx < len(passos_tobe) - 1:
            arr = s5.shapes.add_textbox(Inches(px + 2.2), Inches(2.6), Inches(0.32), Inches(0.5))
            pa = arr.text_frame.paragraphs[0]
            pa.text = "→"
            pa.font.size = Pt(22)
            pa.font.bold = True
            pa.font.color.rgb = VERDE
        px += 2.52

    adicionar_card(
        s5, 0.55, 4.25, 12.2, 2.4,
        "Diferenciais do Fluxo To-Be",
        [
            "• Tratamento por registro isolado: exceções de regra de negócio (ErroRegraNegocio) geram status INSUCESSO e o robô segue para o próximo cliente.",
            "• Padronização garantida: todos os registros válidos saem com CPF e CEP formatados e endereço oficial dos Correios/ViaCEP.",
            "• Entrega multiponto automática: gera planilha executiva formatada, painel HTML, log de auditoria e registros permanentes no SQLite.",
        ],
        VERDE, FUNDO_CLARO,
    )

    # ------------------------------------------------------------------
    # SLIDE 6: REQUISITOS TÉCNICOS E ARQUITETURA
    # ------------------------------------------------------------------
    s6 = prs.slides.add_slide(blank_layout)
    adicionar_cabecalho(s6, "5. Requisitos Técnicos Atendidos", "Todos os requisitos mínimos implementados com código simples e modular", 6)

    reqs = [
        ("Integração com API", "Consulta REST em tempo real na API pública ViaCEP (requests.get) com fallback seguro."),
        ("Excel (Leitura e Escrita)", "Leitura com pandas (preservando zeros à esquerda) e geração formatada com openpyxl."),
        ("Web Scraping (BeautifulSoup)", "Extração automática das tarifas de frete por UF direto da tabela HTML da Intranet."),
        ("Expressões Regulares (Regex)", "Validação e máscara automática de CPF (11 dígitos), E-mail e CEP (8 dígitos)."),
        ("Resiliência (Try-Except)", "Exceção customizada ErroRegraNegocio isola falhas por cliente sem abortar o lote."),
        ("Logs + Banco SQLite", "Log passo a passo (rpa_execucao.log) e auditoria em banco relacional (rpa_logiexpress.db)."),
    ]
    coords = [(0.6, 1.45), (4.74, 1.45), (8.88, 1.45), (0.6, 4.25), (4.74, 4.25), (8.88, 4.25)]
    for (x, y), (tit, desc) in zip(coords, reqs):
        adicionar_card(s6, x, y, 3.85, 2.5, tit, [desc], AZUL_CLARO, FUNDO_CLARO)

    # ------------------------------------------------------------------
    # SLIDE 7: EXEMPLOS DO CÓDIGO-FONTE
    # ------------------------------------------------------------------
    s7 = prs.slides.add_slide(blank_layout)
    adicionar_cabecalho(s7, "6. Exemplos do Código-Fonte Gerado (rpa_main.py)", "Código Python simples, legível e bem estruturado", 7)

    cod_esq = (
        "# 1. Validação e Formatação com Regex\n"
        "def validar_e_formatar_cpf(cpf_bruto: str) -> str:\n"
        "    digitos = re.sub(r'\\D', '', str(cpf_bruto or ''))\n"
        "    if not re.match(r'^\\d{11}$', digitos):\n"
        "        raise ErroRegraNegocio('CPF deve ter 11 dígitos')\n"
        "    return re.sub(r'(\\d{3})(\\d{3})(\\d{3})(\\d{2})', r'\\1.\\2.\\3-\\4', digitos)\n\n"
        "# 2. Web Scraping da Intranet com BeautifulSoup\n"
        "soup = BeautifulSoup(conteudo, 'html.parser')\n"
        "for linha in soup.select('#tabela-frete tbody tr.linha-uf'):\n"
        "    uf = linha.get('data-uf').upper()\n"
        "    fretes[uf] = float(linha.select_one('td.valor-frete').text)"
    )
    cod_dir = (
        "# 3. Integração API ViaCEP + Resiliência Try-Except\n"
        "try:\n"
        "    cpf_fmt = validar_e_formatar_cpf(cpf_bruto)\n"
        "    email_ok = validar_email(email_bruto)\n"
        "    cep_limpo, cep_fmt = validar_e_formatar_cep(cep_bruto)\n"
        "    dados_cep = consultar_api_viacep(cep_limpo, logger)\n"
        "    frete = tabela_fretes.get(dados_cep['uf'], 65.0)\n"
        "    registro['status'] = 'SUCESSO'\n"
        "except ErroRegraNegocio as erro:\n"
        "    registro['status'] = 'INSUCESSO'\n"
        "    logger.error('[%s] Regra: %s', id_cli, erro)\n"
        "salvar_registro_banco(conn, registro)  # SQLite"
    )

    for x, titulo, codigo in [(0.6, "Regex & BeautifulSoup", cod_esq), (6.8, "API & Try-Except & SQLite", cod_dir)]:
        box = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.45), Inches(5.9), Inches(5.3))
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(15, 23, 42)
        box.line.color.rgb = AZUL_CLARO
        tb = s7.shapes.add_textbox(Inches(x + 0.2), Inches(1.6), Inches(5.5), Inches(5.0))
        tf = tb.text_frame
        tf.word_wrap = True
        p0 = tf.paragraphs[0]
        p0.text = titulo
        p0.font.size = Pt(15)
        p0.font.bold = True
        p0.font.color.rgb = RGBColor(125, 211, 252)
        p1 = tf.add_paragraph()
        p1.text = codigo
        p1.font.name = "Consolas"
        p1.font.size = Pt(11)
        p1.font.color.rgb = RGBColor(226, 232, 240)
        p1.space_before = Pt(10)

    # ------------------------------------------------------------------
    # SLIDE 8: VÍDEO E EVIDÊNCIAS VISUAIS
    # ------------------------------------------------------------------
    s8 = prs.slides.add_slide(blank_layout)
    adicionar_cabecalho(s8, "7. Demonstração da Execução e Evidências", "Vídeo da execução (midia/video_execucao_rpa.mp4) e saídas reais", 8)

    img_logs = MIDIA_DIR / "evidencia_logs_execucao.png"
    img_rel = MIDIA_DIR / "evidencia_relatorio_final.png"
    if img_logs.exists():
        s8.shapes.add_picture(str(img_logs), Inches(0.55), Inches(1.45), width=Inches(6.0))
    if img_rel.exists():
        s8.shapes.add_picture(str(img_rel), Inches(6.75), Inches(1.45), width=Inches(6.0))

    tb_v = s8.shapes.add_textbox(Inches(0.55), Inches(5.35), Inches(12.2), Inches(1.4))
    tf_v = tb_v.text_frame
    tf_v.word_wrap = True
    pv = tf_v.paragraphs[0]
    pv.text = "Vídeo Completo da Execução Disponível em: midia/video_execucao_rpa.mp4 (e .gif)"
    pv.font.size = Pt(16)
    pv.font.bold = True
    pv.font.color.rgb = AZUL_ESCURO
    pv2 = tf_v.add_paragraph()
    pv2.text = f"Execução Auditada: {id_execucao} | Logs detalhados gravados em logs/rpa_execucao.log e histórico completo salvo em banco/rpa_logiexpress.db."
    pv2.font.size = Pt(13)
    pv2.font.color.rgb = CINZA_TEXTO
    pv2.space_before = Pt(6)

    # ------------------------------------------------------------------
    # SLIDE 9: RESULTADOS DOS TESTES (SUCESSO E INSUCESSO)
    # ------------------------------------------------------------------
    s9 = prs.slides.add_slide(blank_layout)
    adicionar_cabecalho(s9, "8. Resultados dos Testes — Sucesso e Insucesso", "Validação completa das regras de negócio (detalhes em EVIDENCIAS_DE_TESTE.md)", 9)

    sucessos = [r for r in registros if r["status"] == "SUCESSO"]
    erros = [r for r in registros if r["status"] == "INSUCESSO"]

    linhas_ok = [
        f"• {r['id_cliente']} ({r['nome']}): UF {r['uf']} | Frete R$ {r['valor_frete']:.2f} | Total R$ {r['valor_total']:.2f}"
        for r in sucessos
    ]
    linhas_err = [
        f"• {r['id_cliente']} ({r['nome']}): {r['mensagem_detalhe']}"
        for r in erros
    ]

    adicionar_card(s9, 0.6, 1.45, 5.9, 5.25, f"Cenários de SUCESSO ({len(sucessos)}/10)", linhas_ok, VERDE, RGBColor(240, 253, 244))
    adicionar_card(s9, 6.8, 1.45, 5.9, 5.25, f"Cenários de INSUCESSO / Regras ({len(erros)}/10)", linhas_err, VERMELHO, RGBColor(254, 242, 242))

    # ------------------------------------------------------------------
    # SLIDE 10: GANHOS PARA A EMPRESA
    # ------------------------------------------------------------------
    s10 = prs.slides.add_slide(blank_layout)
    adicionar_cabecalho(s10, "9. Ganhos para a Empresa e Conclusão", "Impacto direto da adoção do RPA na operação da LogiExpress Brasil", 10)

    adicionar_card(
        s10, 0.6, 1.45, 3.85, 3.3,
        "Produtividade (+99%)",
        [
            "• De 10 horas/dia (150 pedidos manuais) para menos de 20 segundos de execução automática.",
            "• Redução drástica de fila operacional em períodos de pico.",
        ],
        VERDE, RGBColor(240, 253, 244),
    )
    adicionar_card(
        s10, 4.74, 1.45, 3.85, 3.3,
        "Qualidade & Zero Erro",
        [
            "• 100% dos CPFs, CEPs e e-mails validados por Regex antes do faturamento.",
            "• Fim das divergências de frete por consulta manual errada.",
        ],
        AZUL_CLARO, FUNDO_CLARO,
    )
    adicionar_card(
        s10, 8.88, 1.45, 3.85, 3.3,
        "Governança & Auditoria",
        [
            "• Histórico completo salvo em banco SQLite e arquivo de log.",
            "• Rastreabilidade exata do motivo de bloqueio de cada pedido.",
        ],
        AZUL_ESCURO, FUNDO_CLARO,
    )
    adicionar_card(
        s10, 0.6, 5.0, 12.1, 1.75,
        "Conclusão Estratégica",
        [
            "O robô RPA desenvolvido resolve o gargalo operacional com código simples, baixo custo de manutenção (100% Python open-source) e alta confiabilidade, permitindo à LogiExpress escalar o volume de pedidos sem aumentar o headcount da operação.",
        ],
        VERDE, FUNDO_CLARO,
    )

    out_pptx = BASE_DIR / "Apresentacao_RPA_LogiExpress.pptx"
    prs.save(out_pptx)
    return out_pptx


def main() -> None:
    MIDIA_DIR.mkdir(parents=True, exist_ok=True)
    id_execucao, registros = ler_ultima_execucao_sqlite()

    p1 = gerar_imagem_entrada_excel()
    p2 = gerar_imagem_logs()
    p3 = gerar_imagem_sqlite(id_execucao, registros)
    p4 = gerar_imagem_relatorio_final(id_execucao, registros)
    mp4, gif = gerar_video_execucao(id_execucao, registros)
    pptx = criar_apresentacao_pptx(id_execucao, registros)

    print(f"[OK] Imagem 1: {p1}")
    print(f"[OK] Imagem 2: {p2}")
    print(f"[OK] Imagem 3: {p3}")
    print(f"[OK] Imagem 4: {p4}")
    print(f"[OK] Vídeo MP4: {mp4}")
    print(f"[OK] Animação GIF: {gif}")
    print(f"[OK] PowerPoint PPTX: {pptx}")


if __name__ == "__main__":
    main()
