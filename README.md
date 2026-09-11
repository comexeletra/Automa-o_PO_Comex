# Conversor de Pedido TOTVS para Excel

O PDF é a única fonte de dados. O sistema mantém o layout do template e deixa vazios NCM, Eletra code, finalidade, centro de custos, LI e projeto. Não faz lookup, substituição de código, MOQ nem inferência.

## Instalação (Windows / VS Code)

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## Arquivos de referência

Adicione, sem modificar os originais:

```text
samples/sample_po_027956.pdf
templates/reference_template.xlsx
```

Crie então o template de produção limpo:

```powershell
python scripts/build_clean_template.py `
  --reference templates/reference_template.xlsx `
  --output templates/po_template.xlsx
```

O utilitário remove conteúdo variável conhecido e fórmulas do bloco de materiais, preservando o layout, abas e bloco inferior do workbook.

## Testes

```powershell
pytest -q
```

Os testes de regressão do pedido 027956 são habilitados automaticamente quando `samples/sample_po_027956.pdf` existir.

## Servidor

```powershell
uvicorn app.main:app --reload
```

Abra `http://127.0.0.1:8000`.

## Watcher opcional

```powershell
$env:WATCH_ENABLED = "true"
python -m app.watcher
```

Ele usa as pastas `storage/incoming`, `storage/processing`, `storage/completed` e `storage/error`, sempre chamando o mesmo serviço usado pela web.

## Limitação atual

O parser é deliberadamente restritivo: espera PDF TOTVS textual com as colunas do layout especificado. Para layout diferente, ele falha com `UNSUPPORTED_TOTVS_PDF_LAYOUT` em vez de produzir planilha possivelmente incorreta. OCR não faz parte deste MVP.
