"""
Script de preparação do ambiente de demonstração do RPA LogiExpress Brasil.

Gera:
1. O portal HTML simulado da Intranet (portal_intranet/tabela_frete_portal.html),
   que será lido pelo robô usando BeautifulSoup.
2. A planilha Excel de entrada (dados/clientes_pendentes.xlsx) com 10 cenários de teste
   cobrindo casos de sucesso e de erro de regra de negócio (insucesso).
"""

from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side


BASE_DIR = Path(__file__).resolve().parent
DADOS_DIR = BASE_DIR / "dados"
PORTAL_DIR = BASE_DIR / "portal_intranet"


def gerar_portal_frete_html() -> Path:
    """Cria a página HTML simulando a Intranet com a tabela oficial de frete por UF."""
    PORTAL_DIR.mkdir(parents=True, exist_ok=True)
    caminho_html = PORTAL_DIR / "tabela_frete_portal.html"

    html = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Intranet LogiExpress - Tabela Oficial de Frete</title>
    <style>
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            background-color: #f4f6f9;
            color: #1f2937;
            margin: 0;
            padding: 32px;
        }
        .container {
            max-width: 820px;
            margin: 0 auto;
            background: #ffffff;
            border-radius: 10px;
            padding: 28px;
            box-shadow: 0 4px 14px rgba(15, 23, 42, 0.08);
        }
        .header {
            border-bottom: 3px solid #1d4ed8;
            padding-bottom: 14px;
            margin-bottom: 20px;
        }
        .header h1 {
            margin: 0;
            color: #1e3a8a;
            font-size: 24px;
        }
        .header p {
            margin: 6px 0 0;
            color: #4b5563;
            font-size: 14px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 12px;
        }
        th {
            background-color: #1e3a8a;
            color: #ffffff;
            text-align: left;
            padding: 12px 14px;
            font-size: 14px;
        }
        td {
            padding: 11px 14px;
            border-bottom: 1px solid #e5e7eb;
            font-size: 14px;
        }
        tr:nth-child(even) {
            background-color: #f8fafc;
        }
        .valor-frete {
            font-weight: 700;
            color: #047857;
        }
        .rodape {
            margin-top: 18px;
            font-size: 12px;
            color: #6b7280;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Intranet LogiExpress Brasil</h1>
            <p>Tabela Oficial de Frete Fixo por Estado (UF) — Vigência 2026</p>
        </div>

        <table id="tabela-frete">
            <thead>
                <tr>
                    <th>UF</th>
                    <th>Região</th>
                    <th>Valor do Frete (R$)</th>
                    <th>Prazo Estimado</th>
                </tr>
            </thead>
            <tbody>
                <tr class="linha-uf" data-uf="SP">
                    <td>SP</td>
                    <td>Sudeste</td>
                    <td class="valor-frete">25.00</td>
                    <td>2 dias</td>
                </tr>
                <tr class="linha-uf" data-uf="RJ">
                    <td>RJ</td>
                    <td>Sudeste</td>
                    <td class="valor-frete">32.00</td>
                    <td>3 dias</td>
                </tr>
                <tr class="linha-uf" data-uf="MG">
                    <td>MG</td>
                    <td>Sudeste</td>
                    <td class="valor-frete">35.00</td>
                    <td>3 dias</td>
                </tr>
                <tr class="linha-uf" data-uf="PR">
                    <td>PR</td>
                    <td>Sul</td>
                    <td class="valor-frete">38.00</td>
                    <td>4 dias</td>
                </tr>
                <tr class="linha-uf" data-uf="BA">
                    <td>BA</td>
                    <td>Nordeste</td>
                    <td class="valor-frete">55.00</td>
                    <td>6 dias</td>
                </tr>
                <tr class="linha-uf" data-uf="DF">
                    <td>DF</td>
                    <td>Centro-Oeste</td>
                    <td class="valor-frete">45.00</td>
                    <td>4 dias</td>
                </tr>
                <tr class="linha-uf" data-uf="OUTROS">
                    <td>OUTROS</td>
                    <td>Demais Regiões</td>
                    <td class="valor-frete">65.00</td>
                    <td>7 dias</td>
                </tr>
            </tbody>
        </table>

        <div class="rodape">
            Sistema Interno de Tarifação Logística | Atualizado pela Diretoria de Operações
        </div>
    </div>
</body>
</html>
"""
    caminho_html.write_text(html, encoding="utf-8")
    return caminho_html


def gerar_planilha_entrada() -> Path:
    """Cria a planilha Excel de entrada com os 10 casos de teste (sucesso e insucesso)."""
    DADOS_DIR.mkdir(parents=True, exist_ok=True)
    caminho_excel = DADOS_DIR / "clientes_pendentes.xlsx"

    registros = [
        # ID, Nome, CPF, Email, CEP, Valor_Pedido
        ("CT01", "Ana Souza", "12345678901", "ana.souza@empresa.com.br", "01001-000", 350.00),
        ("CT02", "Bruno Oliveira", "234.567.890-12", "bruno.oliveira@logi.com", "20040002", 520.50),
        ("CT03", "Carlos Mendes", "123.456", "carlos@email.com", "30130-010", 200.00),
        ("CT04", "Fernanda Lima", "34567890123", "fernanda.lima@comercio.org.br", "30140-071", 890.00),
        ("CT05", "Diego Costa", "456.789.012-34", "diego.costa_sem_arroba.com", "80010-010", 150.00),
        ("CT06", "Eduarda Rocha", "567.890.123-45", "eduarda@rocha.com", "123", 410.00),
        ("CT07", "Gabriel Santos", "678.901.234-56", "gabriel.santos@tech.io", "80020-310", 1200.00),
        ("CT08", "Helena Martins", "789.012.345-67", "helena@martins.com.br", "99999-999", 300.00),
        ("CT09", "Igor Almeida", "890.123.456-78", "igor@almeida.com", "40020-000", -50.00),
        ("CT10", "Juliana Pereira", "90123456789", "juliana.p@brasilia.gov.br", "70040-010", 640.00),
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "Clientes_Pendentes"

    cabecalhos = ["ID_Cliente", "Nome", "CPF", "Email", "CEP", "Valor_Pedido"]
    ws.append(cabecalhos)

    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    borda_fina = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB"),
    )

    for col_idx in range(1, len(cabecalhos) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = borda_fina

    for row_idx, item in enumerate(registros, start=2):
        for col_idx, valor in enumerate(item, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=valor)
            cell.border = borda_fina
            if col_idx in (1, 3, 5):
                # Força texto para preservar zeros à esquerda em CPF e CEP
                cell.number_format = "@"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 6:
                cell.number_format = "R$ #,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    larguras = {"A": 13, "B": 22, "C": 18, "D": 32, "E": 14, "F": 16}
    for col, largura in larguras.items():
        ws.column_dimensions[col].width = largura

    ws.row_dimensions[1].height = 26
    wb.save(caminho_excel)
    return caminho_excel


def main() -> None:
    html_path = gerar_portal_frete_html()
    excel_path = gerar_planilha_entrada()
    print(f"[OK] Portal HTML gerado em: {html_path}")
    print(f"[OK] Planilha de entrada gerada em: {excel_path}")


if __name__ == "__main__":
    main()
