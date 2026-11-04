from pathlib import Path
import re
import openpyxl

# 1. Configurações
BASE_DIR = Path(__file__).resolve().parents[2]
ARQUIVO_ENTRADA = BASE_DIR / "docs" / "bundle_python.xlsx"
ARQUIVO_SAIDA = BASE_DIR / "docs" / "bundle_python_links.xlsx"
NOME_ABA = None  # Deixe None para usar a aba ativa, ou coloque ex: "Planilha1"
COLUNA_ORIGEM = "B"  # Coluna onde estão os textos com links
COLUNA_DESTINO = "C"  # Coluna onde os links serão colados
LINHA_INICIAL = 2  # 2 se a linha 1 tiver cabeçalho

# 2. Carrega o arquivo (data_only=False permite ler fórmulas =HIPERLINK se houver)
wb = openpyxl.load_workbook(ARQUIVO_ENTRADA, data_only=False)
ws = wb[NOME_ABA] if NOME_ABA else wb.active

# Cabeçalho opcional na nova copoetluna
if LINHA_INICIAL > 1:
    ws[f"{COLUNA_DESTINO}1"] = "Link Extraído"

# 3. Percorre as linhas e extrai o link
for linha in range(LINHA_INICIAL, ws.max_row + 1):
    celula_origem = ws[f"{COLUNA_ORIGEM}{linha}"]
    url = None

    # Caso 1: Hiperlink padrão do Excel (inserido via Ctrl+K / botão direito)
    if celula_origem.hyperlink and celula_origem.hyperlink.target:
        url = celula_origem.hyperlink.target

    # Caso 2: Link criado via fórmula =HYPERLINK("url", "texto") ou =HIPERLINK(...)
    elif isinstance(
        celula_origem.value, str
    ) and celula_origem.value.upper().startswith(("=HYPERLINK", "=HIPERLINK")):
        match = re.search(r'["\'](https?://[^"\']+)["\']', celula_origem.value)
        if match:
            url = match.group(1)

    # Escreve na nova coluna se encontrou um link
    if url:
        ws[f"{COLUNA_DESTINO}{linha}"].value = url

# 4. Salva o resultado
wb.save(ARQUIVO_SAIDA)
print(f"Concluído! Arquivo salvo em: {ARQUIVO_SAIDA}")
