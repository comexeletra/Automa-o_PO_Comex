# Conversor de Pedido TOTVS para Excel

O arquivo enviado (PDF ou exportação Excel do TOTVS) fornece os dados do pedido. Nos fluxos de Excel e PDF, o Supply code é relacionado ao Product Code do database para preencher a NCM e substituir a descrição reduzida pela descrição completa. Códigos, quantidades e valores do pedido são preservados. Eletra code, finalidade, centro de custos, LI e projeto continuam vazios. Os quadros históricos de substituição de códigos e MOQ são removidos da planilha gerada.

## Excel TOTVS → PO DRAFT

No botão **Processar PO**, envie a exportação `.xlsx` no formato **Listagem do Browse**, como `templates/new_po_model/PO 033193.xlsx`. O sistema identifica o cabeçalho pelos nomes das colunas e gera o layout de 14 colunas do `PO DRAFT.xlsx`. O fluxo de PDF continua usando o template anterior; **Processar PO menor** aceita apenas PDF.

O template do novo fluxo está em `app/resources/po_draft_template.xlsx`, copiado do DRAFT fornecido. Para selecionar outro arquivo compatível, configure `DRAFT_TEMPLATE_PATH`. Os dois arquivos originais em `templates/new_po_model` são preservados.

| Origem TOTVS | Destino no DRAFT |
| --- | --- |
| Numero | PO # e observações |
| Item | # |
| Produto | Supply code (texto, inclusive códigos alfanuméricos) |
| Descricao | Product description, substituída pela descrição completa do database quando encontrada |
| Quantidade | Total QTY |
| Prc Unitario | Unit price |
| Vlr.Total | Total (CNY) |
| Desc Fornec | Seller |
| Filial | Buyer, caso o DRAFT não contenha a razão social |
| Numero da SC | Observações, sem duplicação |

Os totais são valores numéricos somados da origem. O bloco de itens cresce ou diminui preservando aprovação, assinaturas, logo, observações e área de impressão. A data no cabeçalho é a data da conversão em São Paulo, como no fluxo PDF. Os dados cadastrais fixos do DRAFT são preservados: razão social do comprador, CNPJ, endereços e contatos. A NCM vem do database. Modal, Incoterm, pagamento, solicitante e mês de produção ficam vazios neste fluxo, pois esses dados não estão preenchidos na exportação de referência. As abas auxiliares ocultas do DRAFT são preservadas, sem consulta às suas tabelas ou fórmulas.

O conversor rejeita planilhas sem o cabeçalho esperado, com mais de uma PO/filial/fornecedor, itens duplicados ou números inválidos. Os testes cobrem redução e expansão do layout, valores decimais brasileiros, preservação dos desenhos e download pela interface.

## Database de produtos

A base usada pela plataforma está em `app/resources/product_database.xlsx`, copiada de `templates/new_po_model/DATABASE - Renew.xlsx`. A primeira aba é `DATABASE - SANEAMENTO`, com as colunas `Product Code`, `Product Description` e `NCM DATABASE`. Configure `PRODUCT_DATABASE_PATH` para selecionar outro arquivo; para atualizar a base padrão, substitua a cópia em `app/resources`.

A correspondência é exata, preservando zeros à esquerda e códigos alfanuméricos. A NCM é gravada como texto com oito dígitos. Itens ausentes ou com cadastros duplicados divergentes mantêm a descrição original, sem NCM, e recebem aviso para revisão. Campos incompletos também geram aviso; base ausente ou sem as colunas obrigatórias interrompe a conversão. As descrições completas usam quebra de texto e altura de linha ajustada.

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

Os testes de regressão cobrem os layouts extensos com coordenadas normais ou texto escalado e o layout compacto com código numérico. Testes sintéticos cobrem código alfanumérico, preço adjacente ao IPI e pagamento AVISTA. O fluxo principal `POST /process` reconhece todas essas variações; `POST /process-small` permanece disponível para o layout compacto. PDFs internos usados em validações locais não precisam ser publicados no repositório.

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

O parser é deliberadamente restritivo: espera PDF TOTVS textual com um dos layouts já suportados. Para layout diferente, ele falha com erro de layout não suportado em vez de produzir planilha possivelmente incorreta. OCR não faz parte deste MVP.

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
