# Conversor de Pedido TOTVS para Excel

O PDF é a única fonte de dados. O sistema mantém o layout do template e deixa vazios NCM, Eletra code, finalidade, centro de custos, LI e projeto. Não faz lookup, substituição de código, MOQ nem inferência.

## Instalação (Windows / VS Code)

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements-dev.txt
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

Por padrão, a interface aceita PDFs de até 4 MB e entrega planilhas de até 4 MB. Esses limites deixam margem em relação ao teto de 4,5 MB por requisição ou resposta das Vercel Functions; arquivos maiores recebem uma mensagem de erro. As variáveis `MAX_UPLOAD_MB` e `MAX_OUTPUT_MB` permitem configurar os limites, mas valores acima de 4 MB exigem outro fluxo de upload/download na Vercel.

## Watcher opcional

```powershell
$env:WATCH_ENABLED = "true"
python -m app.watcher
```

Ele usa as pastas `storage/incoming`, `storage/processing`, `storage/completed` e `storage/error`, sempre chamando o mesmo serviço usado pela web.

## Limitação atual

O parser é deliberadamente restritivo: espera PDF TOTVS textual com as colunas do layout especificado. Para layout diferente, ele falha com `UNSUPPORTED_TOTVS_PDF_LAYOUT` em vez de produzir planilha possivelmente incorreta. OCR não faz parte deste MVP.

## Vercel

O deploy da primeira versão na Vercel usa o runtime **Python** com a instância
FastAPI existente em `app/main.py`. Não use Edge Runtime: o conversor depende de
arquivos temporários e de bibliotecas Python. `pyproject.toml` declara o
entrypoint explicitamente e `vercel.json` reserva até 60 segundos para a
Function; os diretórios de templates, estáticos e recursos continuam montados
pela aplicação.

Os arquivos de trabalho são criados por requisição e descartados ao final. Em
produção, o sistema de arquivos é somente leitura, exceto por `/tmp`; não use
esse diretório como armazenamento persistente. Mantenha `MAX_UPLOAD_MB` e
`MAX_OUTPUT_MB` em até 4 MB: a Vercel limita tanto o corpo da requisição quanto
a resposta da Function a 4,5 MB.

Valide localmente antes de criar um preview:

```powershell
python -B -m pytest -q -p no:cacheprovider
vercel dev
```

Depois, valide no preview as rotas `/`, `/health`, `/static`, `/assets` e a
conversão do PDF de referência antes de qualquer promoção para produção. Não
remova os arquivos de Cloudflare enquanto a alternativa não estiver aceita.

## Cloudflare Workers

O projeto usa Python Workers com o adaptador ASGI do FastAPI. O processamento do PDF e a criação do Excel ocorrem em armazenamento temporário; a planilha é devolvida na resposta do `POST /process`, pois o sistema de arquivos do Worker não é persistente.

No painel **Workers & Pages > Builds**, configure o repositório com:

```text
Build command: uv sync
Deploy command: uv run pywrangler deploy
```

Não use `npx wrangler deploy`: ele não prepara as dependências Python. A configuração do Worker está em `wrangler.jsonc`; `pyproject.toml`, `uv.lock` e `pylock.toml` fixam as dependências compatíveis com Pyodide.

Para validar ou publicar pelo terminal, com `uv` instalado:

```powershell
uv run pywrangler deploy --dry-run
uv run pywrangler deploy
```
