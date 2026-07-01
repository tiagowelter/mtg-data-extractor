# Magic Extractor

Ferramenta local para colar uma imagem de carta de Magic, reconhecer a carta com OCR local e preencher uma linha pronta para colar no Google Sheets ou salvar em Excel.

## Instalação

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
```

O OCR usa Tesseract local. No Windows, instale pelo instalador oficial/UB Mannheim ou via `winget`:

```powershell
winget install UB-Mannheim.TesseractOCR
```

Depois rode pelo lançador:

```powershell
.\abrir_magic_extractor.bat
```

## Uso

1. Clique em `Atualizar base` na primeira execução. O app baixa os dados públicos do Scryfall para a pasta `data`.
2. Copie uma imagem da carta para a área de transferência.
3. Clique em `Colar imagem`.
4. Clique em `Extrair`. Enquanto o OCR estiver rodando, o app limpa os campos antigos, mostra uma barra de loading e bloqueia os botões até terminar.
5. Revise os campos, marque `Foil` se a carta física for foil, e clique em `Copiar linha`, `Copiar Google Docs` ou `Salvar Excel`.

O arquivo padrão é `magic_cards.xlsx`. A linha copiada é separada por TAB, pronta para colar em uma planilha do Google.
O botão `Copiar Google Docs` copia a mesma linha sem cabeçalho em texto simples, separada por ` | `.

Depois de atualizar o código, abra pelo `abrir_magic_extractor.bat`. Ele fecha somente processos antigos do `magic_extractor.py` e abre a versão atual.
No título da janela deve aparecer `Magic Extractor v2026-07-01.18`; se não aparecer, é uma janela antiga.
O botão `Extrair` usa exatamente a imagem que aparece no preview; depois de copiar uma nova imagem, clique em `Colar imagem` antes de extrair.
Se uma extração parecer suspeita, o app limpa os campos e grava o diagnóstico em `last_extraction_debug.txt`.
