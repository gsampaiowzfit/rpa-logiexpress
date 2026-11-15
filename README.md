# RPA LogiExpress Brasil — Automação de Validação Cadastral e Cálculo de Frete

Projeto acadêmico/corporativo simples e completo de **Robotic Process Automation (RPA)** em Python para eliminar tarefas repetitivas no processo de onboarding de clientes e cálculo de frete pré-faturamento da empresa fictícia **LogiExpress Brasil**.

---

## Estrutura dos Entregáveis

| Arquivo / Pasta | Descrição |
|---|---|
| `rpa_main.py` | Código-fonte principal do robô RPA (modular, simples e comentado) |
| `gerar_ambiente.py` | Prepara a planilha de entrada (`dados/clientes_pendentes.xlsx`) e o portal HTML da intranet |
| `consultar_banco.py` | Consulta e exibe os registros auditados no banco SQLite |
| `gerar_entregaveis.py` | Gera o vídeo (`.mp4`/`.gif`), imagens de evidência (`.png`) e a apresentação `.pptx` |
| `APRESENTACAO_PROJETO.md` | Apresentação completa em Markdown (Case, Justificativa, Fluxos As-Is / To-Be, Código, Vídeo e Ganhos) |
| `apresentacao_slides.html` | Versão interativa em slides HTML (abra no navegador para apresentar) |
| `Apresentacao_RPA_LogiExpress.pptx` | Apresentação executiva em PowerPoint |
| `EVIDENCIAS_DE_TESTE.md` | Caderno completo de cenários de teste e evidências de sucesso e insucesso |
| `evidencias_teste.html` | Versão visual imprimível (PDF) das evidências de teste |
| `dados/clientes_pendentes.xlsx` | Planilha Excel de entrada (10 casos de teste) |
| `dados/relatorio_processamento.xlsx` | Planilha Excel de saída gerada pelo robô |
| `portal_intranet/tabela_frete_portal.html` | Página HTML simulando a intranet lida via `BeautifulSoup` |
| `portal_intranet/painel_resultado.html` | Painel HTML gerado após a execução |
| `banco/rpa_logiexpress.db` | Banco de dados SQLite com histórico de auditoria |
| `logs/rpa_execucao.log` | Log passo a passo de todas as etapas do RPA |
| `midia/video_execucao_rpa.mp4` | Vídeo demonstrando o robô em execução |

---

## Requisitos Técnicos Atendidos

- **Leitura e Geração de Excel**: `pandas` + `openpyxl`
- **Integração com API**: Consulta REST ao serviço público **ViaCEP** (`requests`)
- **Web Scraping HTML**: Extração da tabela oficial de frete por UF com **BeautifulSoup4**
- **Expressões Regulares (Regex)**: Validação e formatação automática de CPF, E-mail e CEP (`re`)
- **Tratamento de Exceções**: `try-except` com exceção customizada `ErroRegraNegocio` garantindo resiliência por registro
- **Geração de Logs**: `logging` gravando cada etapa em `logs/rpa_execucao.log` e console
- **Banco de Dados**: Persistência completa (sucessos e insucessos) em **SQLite** (`sqlite3`)

---

## Como Executar do Zero

1. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
2. Gere a massa de dados e o portal HTML da intranet:
   ```bash
   python gerar_ambiente.py
   ```
3. Execute o robô RPA:
   ```bash
   python rpa_main.py
   ```
4. Consulte os registros salvos no banco SQLite:
   ```bash
   python consultar_banco.py
   ```
5. (Opcional) Gere novamente os entregáveis visuais (vídeo, PNGs e PPTX):
   ```bash
   python gerar_entregaveis.py
   ```
