# CODEX IMPLEMENTATION SPEC
## TOTVS PDF -> Excel estruturado (MVP determinístico)

> **Use este arquivo como instrução principal no Codex dentro do VS Code.**
>
> O objetivo é **implementar**, não apenas planejar. Leia este documento inteiro antes de alterar o repositório.

---

# 1. Objetivo do projeto

Construir uma aplicação Python que receba um relatório PDF de Pedido de Compras gerado pela TOTVS e gere um arquivo Excel com a mesma estrutura/layout do workbook de referência fornecido.

A aplicação deve funcionar como um **conversor determinístico**:

**PDF TOTVS -> extração -> mapeamento -> template Excel -> XLSX final**

A regra principal do sistema é:

> **O PDF é a única fonte de verdade para os dados do pedido.**
>
> Se um valor existe no PDF e possui um campo claramente correspondente na planilha, copiar o valor do PDF.
>
> Se uma coluna/campo existe na planilha, mas não existe diretamente no PDF ou o mapeamento é ambíguo, manter a estrutura e deixar o valor vazio.

A aplicação **NÃO é um motor de regras comerciais**.

---

# 2. Mudança obrigatória em relação ao MVP anterior

Se o repositório já possuir código/configuração para:

- substituição automática de códigos;
- MOQ;
- arredondamento de quantidade;
- cadastro mestre;
- NCM por lookup;
- Eletra Code por lookup;
- preço por lookup externo;
- regras históricas;
- comparação com alterações manuais da planilha de referência;

essas funcionalidades devem ser **removidas do fluxo principal do MVP**.

Arquivos antigos como estes, caso existam, não devem participar do processamento:

```text
app/services/rules.py
app/services/master_data.py
config/rules.sample.yml
config/master_data.sample.csv
```

Eles podem ser excluídos ou movidos para uma área claramente marcada como `future/`, mas **não podem alterar o resultado**.

Também devem ser removidos/reescritos testes que esperem substituições ou MOQ.

A planilha de referência é usada para **layout e estrutura**, não como fonte dos valores finais.

---

# 3. Regras de negócio do MVP

## 3.1 Regras obrigatórias

1. O código do produto no Excel deve ser exatamente o código existente no PDF.
2. A quantidade no Excel deve ser exatamente a quantidade existente no PDF.
3. O preço unitário deve ser exatamente o preço existente no PDF.
4. O valor total do item deve ser exatamente o valor total existente no PDF.
5. A descrição deve ser formada exclusivamente pelo texto de descrição visível no PDF.
6. Nenhum código pode ser substituído.
7. Nenhuma quantidade pode ser ajustada por MOQ.
8. Nenhum dado pode ser consultado em outra planilha/banco/API.
9. Nenhuma informação ausente pode ser inferida.
10. Nenhum valor da planilha manual de referência pode sobrescrever o PDF.
11. Campos extras da planilha ficam vazios.
12. O processamento deve manter a ordem dos itens do PDF.
13. O sistema deve preservar zeros à esquerda quando o campo for identificador textual.
14. Valores financeiros devem ser processados com `Decimal`.
15. O PDF fornecido possui camada de texto; OCR não faz parte do caminho principal do MVP.

## 3.2 Exemplos do comportamento correto

Se o PDF possuir:

```text
2010101620
Quantidade: 2.520.000,000
```

o Excel deve conter:

```text
Supply code = 2010101620
Total QTY   = 2520000
```

Mesmo que a planilha manual de referência tenha trocado esse código por outro.

Se o PDF possuir:

```text
2020302563
Quantidade: 263.800,000
```

o Excel deve conter:

```text
Total QTY = 263800
```

e **não** `264000`.

Se o PDF possuir:

```text
2054300520
Quantidade: 95.497,000
```

o Excel deve conter `95497`, e **não** `96000`.

---

# 4. Arquivos de referência

O workspace deve conter, preferencialmente:

```text
samples/
    sample_po_027956.pdf

templates/
    reference_template.xlsx
```

Se os nomes forem diferentes, localizar os arquivos fornecidos pelo usuário e ajustar os testes/documentação.

Arquivos de referência originais:

```text
PO 027956 - SEA OCT.2026 - MRP HEXING HANGZHOU - SC 019347.pdf
PO 027956 - SEA OCT.2026 - MRP HEXING HANGZHOU - SC 019347 - REV 03 - 29.06.2026.xlsx
```

Nunca modificar os arquivos de referência originais.

---

# 5. Baseline confirmado do PDF de referência

O PDF de teste possui:

```text
Pedido:                  027956
Fornecedor:              HEXING ELECTRICAL CO.,LTD HANGZHOU
Código fornecedor:       EX000135
Empresa:                  ELETRA ENERGY SOLUTIONS
CNPJ:                     12.115.480/0001-15
Data de emissão:          16/06/2026
Condição de pagamento:    240 DIAS
Quantidade de páginas:    8
Quantidade de itens:      193
Total das mercadorias:    9.361.869,50
Soma das quantidades:     53.389.684
SCs encontrados:          019347 e 019458
```

Primeiro item:

```text
Supply code:  2010100040
Quantidade:   75.000,000
Unit price:   0,0104160
Valor total:  781,20
```

Último item do PDF:

```text
Supply code:  2150184130
Quantidade:   652,000
Unit price:   8,5302630
Valor total:  5.561,73
```

Estes valores devem existir em testes de regressão.

---

# 6. Casos obrigatórios para provar que não existem regras manuais

Criar testes que confirmem explicitamente os seguintes valores do PDF:

| Supply code | Quantidade correta |
|---|---:|
| 2020302563 | 263800 |
| 2020303211 | 77000 |
| 2020302542 | 25766 |
| 2052700170 | 232000 |
| 2054300520 | 95497 |
| 2061500150 | 55000 |
| 2150160560 | 519 |
| 3030502830 | 74350 |
| 3030503490 | 134800 |
| 3030702050 | 150 |
| 3030702090 | 50 |
| 3030900500 | 105600 |

Também confirmar:

```text
2010101620 permanece 2010101620
2010101960 permanece 2010101960
2020101980 permanece 2020101980
2020700071 permanece 2020700071
2060500112 permanece 2060500112
```

O output não deve introduzir os códigos substitutos presentes na versão manual da planilha.

---

# 7. Estrutura confirmada do workbook de referência

O workbook contém estas abas:

```text
Material for production
Revision status 11.10.22
Old descriptions
Checklist (for Eletra's use)
Plan1
```

A aba principal é:

```text
Material for production
```

Ela utiliza colunas `A:S`.

O cabeçalho da tabela de materiais está na linha `18`.

Os headers são:

```text
A  PO #
B  #
C  Supply code
D  NCM
E  Eletra code
F  Product description
G:L área visual associada à descrição / sem novos campos de negócio
M  Total QTY
N  Unit price
O  Total (CNY)
P  Finalidade
Q  Centro de custos
R  Requer LI?
S  Projeto
```

No arquivo de referência, `F18:L18` está mesclado para o título `Product description`.

No exemplo de 193 itens:

```text
header:      linha 18
itens:       linhas 19 até 211
total:       linha 212
approval:    inicia na linha 214
remarks:     aparece mais abaixo
```

**Não assumir que futuros PDFs sempre terão 193 itens.**

O bloco de itens deve ser redimensionado dinamicamente.

---

# 8. Mapeamento PDF -> Excel

## 8.1 Tabela de materiais

Este mapeamento é autoritativo para o MVP:

| Excel | Valor |
|---|---|
| `PO #` | número do pedido do PDF |
| `#` | sequência estrutural `1..N` |
| `Supply code` | código do produto do PDF |
| `NCM` | vazio |
| `Eletra code` | vazio |
| `Product description` | descrição do PDF |
| `Total QTY` | quantidade do PDF |
| `Unit price` | valor unitário do PDF |
| `Total (CNY)` | valor total do item do PDF |
| `Finalidade` | vazio |
| `Centro de custos` | vazio |
| `Requer LI?` | vazio |
| `Projeto` | vazio |

Não usar o campo `CC` do PDF para preencher `Centro de custos`, pois não foi confirmado que representam o mesmo conceito.

## 8.2 Cabeçalho do workbook

Preencher somente correspondências claras.

Sugestão conservadora:

```text
Seller name        <- Razão Social do fornecedor
Seller address     <- Endereço do fornecedor disponível no PDF
Buyer name         <- Empresa do PDF
Buyer CNPJ         <- CNPJ/CPF da empresa
Buyer address      <- endereço da empresa
Date               <- Data de Emissão
Payment terms      <- condição de pagamento
```

Manter vazios:

```text
Modal
Incoterm
Requested by
Month of production
Seller phone
Seller e-mail
```

Não inferir `Month of production` a partir da data de entrega.

Não inferir `Requested by` a partir de `Comprador Responsável`.

Se algum campo do cabeçalho se mostrar semanticamente ambíguo durante a implementação, deixá-lo vazio e documentar.

---

# 9. Modelo de dados

Criar modelos independentes de PDF e Excel.

Exemplo:

```python
from dataclasses import dataclass
from datetime import date
from decimal import Decimal


@dataclass(frozen=True)
class PurchaseOrderItem:
    source_item_group: str
    sequence: int
    product_code: str
    description: str
    unit: str | None
    quantity: Decimal
    unit_price: Decimal
    total_value: Decimal
    delivery_date: date | None
    pdf_cost_center: str | None
    sc_number: str | None


@dataclass(frozen=True)
class PurchaseOrder:
    po_number: str
    company_name: str | None
    company_cnpj: str | None
    company_address: str | None
    supplier_name: str | None
    supplier_code: str | None
    supplier_address: str | None
    issue_date: date | None
    payment_terms: str | None
    merchandise_total: Decimal | None
    items: tuple[PurchaseOrderItem, ...]
```

Não adicionar ao modelo campos como:

```text
substitute_code
moq
adjusted_quantity
ncm_lookup
master_price
```

no MVP.

---

# 10. Extração do PDF

## 10.1 Biblioteca principal

Usar:

```text
PyMuPDF (fitz)
```

O PDF é digital/textual.

Não usar OCR como solução principal.

Caso não exista camada de texto suficiente, retornar erro:

```text
PDF_TEXT_LAYER_NOT_FOUND
```

e documentar que OCR é uma evolução futura.

## 10.2 Estratégia

Não depender apenas de `page.get_text()` com split por espaços, pois algumas colunas visuais do TOTVS podem aparecer concatenadas.

Exemplo real:

```text
0,01041601,30
```

visualiza no PDF:

```text
Valor Unitário = 0,0104160
% IP           = 1,30
```

Portanto, o parser deve preferir **posição X/Y dos caracteres/palavras**.

Implementar extração por layout:

1. localizar o header da tabela:
   - `Ite`
   - `Produto`
   - `Descricao`
   - `Quantidade`
   - `Valor Unitario`
   - `% IP`
   - `Valor Total`
   - `Dt. Entreg`
   - `CC`
   - `Nro.S`

2. identificar linhas de início de item com padrão:
   ```regex
   ^\d{3}\s+\d{10}
   ```

3. usar posição horizontal para separar:
   - item;
   - produto;
   - descrição;
   - unidade;
   - quantidade;
   - preço unitário;
   - IPI;
   - valor total;
   - data;
   - CC;
   - SC.

4. concatenar as linhas subsequentes localizadas na coluna de descrição até o próximo início de item.

5. ignorar:
   - cabeçalhos repetidos;
   - `Continua na Proxima Pagina`;
   - `- continuacao`;
   - rodapé;
   - bloco final de totais/aprovação.

6. preservar a ordem física dos itens no documento.

## 10.3 Colunas por coordenada

No PDF de referência, aproximadamente:

```text
Item             x ~ 19
Product code     x ~ 34
Description      x ~ 95 até ~191
Unit             x ~ 193
Quantity         x ~ 210 até ~253
Second qty       x ~ 309 até ~325
Unit price       x ~ 351 até limite antes de % IP
% IP             inicia ~382
Total value      x ~ 415 até ~445
Delivery date    x ~ 446 até ~479
CC               x ~ 493
SC               x ~ 530
```

Não tratar esses números como verdade universal sem validação.

Criar uma classe/configuração, por exemplo:

```python
@dataclass(frozen=True)
class TotvsPdfLayout:
    product_x0: float
    description_x0: float
    description_x1: float
    unit_x0: float
    quantity_x0: float
    quantity_x1: float
    unit_price_x0: float
    ip_x0: float
    total_x0: float
    delivery_x0: float
    cc_x0: float
    sc_x0: float
```

O parser deve validar que o header está em posições compatíveis.

Se o layout estiver muito diferente, retornar:

```text
UNSUPPORTED_TOTVS_PDF_LAYOUT
```

em vez de gerar um Excel potencialmente incorreto.

---

# 11. Descrição do produto

A descrição deve ser a descrição **visível no PDF**.

Exemplo do primeiro item:

```text
RESISTOR 390K 1% 100PPMM 1/4W
1206 (R06027) SMT (2010100040)
```

O Excel pode receber:

```text
RESISTOR 390K 1% 100PPMM 1/4W 1206 (R06027) SMT (2010100040)
```

Normalização permitida:

- remover quebras de linha internas;
- reduzir múltiplos espaços a um;
- trim nas extremidades.

Não permitido:

- consultar descrição de outro arquivo;
- corrigir ortografia;
- traduzir;
- substituir caracteres por interpretação comercial;
- completar trechos que não estão visíveis no PDF.

---

# 12. Parsing numérico

Criar funções isoladas e testadas.

## 12.1 Decimal brasileiro

Exemplos:

```text
75.000,000    -> Decimal("75000.000")
1.155.000,000 -> Decimal("1155000.000")
781,20        -> Decimal("781.20")
12.127,50     -> Decimal("12127.50")
0,0104160     -> Decimal("0.0104160")
39,6300000    -> Decimal("39.6300000")
```

Função sugerida:

```python
def parse_ptbr_decimal(value: str) -> Decimal:
    ...
```

Não usar `float` na camada de parsing.

## 12.2 Validação financeira

Para cada item:

```python
expected = (quantity * unit_price).quantize(
    Decimal("0.01"),
    rounding=ROUND_HALF_UP,
)
```

Comparar com `total_value`.

Permitir diferença máxima de `0.01` devido a arredondamento/exibição.

Se a diferença for maior, não modificar valores.

Registrar warning:

```text
ITEM_TOTAL_MISMATCH
```

O PDF continua sendo a fonte de verdade.

---

# 13. SC / Nro.S

No relatório TOTVS, o número de SC pode quebrar em duas linhas.

Exemplo visual/textual:

```text
01934
7
```

deve ser reconstruído como:

```text
019347
```

Outro exemplo:

```text
01945
8
```

deve ser:

```text
019458
```

Preservar como `str`, nunca número inteiro.

O SC não precisa preencher nenhuma coluna da tabela principal neste MVP, mas deve ser extraído para:

- metadados;
- logs;
- tela de resultado;
- testes.

---

# 14. Excel: princípio de implementação

Usar:

```text
openpyxl
```

A saída deve ser criada a partir de uma cópia do workbook de referência/template.

Objetivos:

- preservar visual;
- preservar larguras;
- preservar alturas;
- preservar fontes;
- preservar bordas;
- preservar fills;
- preservar merges estruturais;
- preservar configuração de impressão;
- preservar nomes das abas;
- preservar abas ocultas;
- preservar o checklist e demais partes estruturais.

O arquivo de referência contém dados manuais e fórmulas externas.

**Esses valores não podem vazar para novos pedidos.**

---

# 15. Criar um template limpo

Implementar um utilitário:

```text
scripts/build_clean_template.py
```

Uso:

```powershell
python scripts/build_clean_template.py `
  --reference templates/reference_template.xlsx `
  --output templates/po_template.xlsx
```

O script deve:

1. copiar a estrutura/layout do workbook;
2. remover fórmulas externas da tabela de materiais;
3. remover valores específicos da PO 027956 no bloco variável;
4. deixar as colunas extras vazias;
5. manter labels/cabeçalhos/formatação;
6. remover links externos do workbook sempre que possível;
7. salvar `templates/po_template.xlsx`.

A aplicação em produção usa:

```text
templates/po_template.xlsx
```

e **não** o arquivo manual de referência.

O template limpo deve ser versionado no projeto depois de gerado.

---

# 16. External links e VLOOKUP

A planilha de referência possui fórmulas semelhantes a:

```excel
=VLOOKUP(...)
```

com referências externas.

No arquivo gerado:

- `D` deve estar vazio;
- `E` deve estar vazio;
- `F` deve conter valor da descrição do PDF;
- `N` deve conter valor do preço do PDF;
- `O` deve conter valor total do PDF;
- `P:S` devem estar vazios.

Não deixar `VLOOKUP` nas linhas dos materiais.

Não deixar `#NAME?`, `#REF!` ou fórmulas externas na tabela de materiais.

Preferir abrir o template com:

```python
load_workbook(..., keep_links=False)
```

e verificar o arquivo final.

Criar um teste que abra o `.xlsx` como ZIP e confirme, se tecnicamente viável, que não existem partes `xl/externalLinks/` no arquivo final.

Se `openpyxl` não remover todas as external references automaticamente, implementar uma limpeza OOXML isolada e testada.

---

# 17. Redimensionamento dinâmico do bloco de itens

Não hardcode `19:211`.

Localizar:

```text
header_row = linha que contém:
PO # | # | Supply code | ... | Total (CNY)
```

Localizar também a seção:

```text
APPROVAL
```

e o total existente.

Fluxo sugerido:

1. localizar `header_row`;
2. `item_start = header_row + 1`;
3. localizar `approval_row`;
4. identificar o total/spacer entre itens e approval;
5. descobrir a capacidade original;
6. redimensionar para exatamente `N` itens;
7. copiar o estilo de uma linha modelo para novas linhas quando necessário;
8. remover linhas excedentes para PDFs menores;
9. nunca deixar produtos antigos do template.

No exemplo:

```text
N = 193
item_start = 19
item_end = 211
total_row = 212
```

Para `N` genérico:

```text
item_end = item_start + N - 1
total_row = item_end + 1
```

A seção `APPROVAL` e tudo abaixo deve ser deslocado corretamente.

---

# 18. Merges dinâmicos

O workbook de referência possui, entre outros:

```text
F18:L18
A19:A211
A212:L212
N212:O212
```

Ao redimensionar o bloco:

- preservar `F18:L18`;
- recriar a merge da PO:
  ```text
  A{item_start}:A{item_end}
  ```
- colocar o PO no top-left:
  ```text
  A{item_start}
  ```
- recriar merges do total conforme o template.

Antes de inserir/deletar linhas, desfazer merges que cruzem a região dinâmica se necessário.

Depois, recriar os merges com os endereços corretos.

---

# 19. Escrita das linhas

Para cada item `i`:

```text
row = item_start + i
```

Preencher:

```python
ws.cell(row, 1).value  # somente top-left do merge da PO
ws.cell(row, 2).value = item.sequence
ws.cell(row, 3).value = item.product_code
ws.cell(row, 4).value = None
ws.cell(row, 5).value = None
ws.cell(row, 6).value = item.description

# G:L permanecem vazias

ws.cell(row, 13).value = item.quantity
ws.cell(row, 14).value = item.unit_price
ws.cell(row, 15).value = item.total_value

ws.cell(row, 16).value = None
ws.cell(row, 17).value = None
ws.cell(row, 18).value = None
ws.cell(row, 19).value = None
```

Não criar fórmulas em `D`, `F`, `N` ou `O`.

---

# 20. Formatação numérica

Preservar/aplicar:

```text
Supply code      -> texto ou número sem notação científica
Total QTY        -> número, sem separador decimal quando .000
Unit price       -> até 7 casas decimais no mínimo
Total (CNY)      -> 2 casas decimais
```

Sugestão:

```python
quantity_format = '#,##0.###'
unit_price_format = '0.0000000'
money_format = '#,##0.00'
```

Não transformar `product_code` em notação científica.

Para segurança, pode ser armazenado como `str`.

---

# 21. Linha de total

Após o último item, manter a linha de total do template.

É permitido gerar valores determinísticos derivados dos itens.

Preencher:

```text
Total QTY = soma das quantidades extraídas do PDF
Total CNY = Total das Mercadorias do PDF
```

No PDF de referência:

```text
Total QTY = 53.389.684
Total CNY = 9.361.869,50
```

Se `Total das Mercadorias` não puder ser extraído, usar a soma dos `total_value` dos itens e registrar warning.

O total da mercadoria do PDF deve ser preferido quando disponível.

---

# 22. Conteúdo manual da planilha de referência

A planilha manual possui conteúdo que **não deve ser tratado como fonte**.

Exemplos:

```text
MATERIA-PRIMA
COMPRAS
NÃO
N/A
substituições
MOQ
quantidades ajustadas
observações de revisão
```

Esses valores específicos não devem ser copiados para as novas linhas de materiais.

No bloco `P:S`, todas as linhas de item devem ficar vazias.

As abas auxiliares podem ser preservadas estruturalmente, mas **não podem ser consultadas para enriquecer a saída**.

---

# 23. Serviço central

Criar uma única API interna:

```python
def process_purchase_order(
    pdf_path: Path,
    output_path: Path,
    template_path: Path,
) -> ProcessingResult:
    ...
```

Fluxo:

```text
parse PDF
    ↓
validate parsed structure
    ↓
copy/load clean template
    ↓
resize item block
    ↓
write mapped fields
    ↓
validate workbook
    ↓
save XLSX
    ↓
return ProcessingResult
```

A interface web e o watcher devem usar exatamente esse serviço.

---

# 24. Estrutura recomendada do repositório

```text
totvs_po_automation/
│
├── app/
│   ├── __init__.py
│   ├── main.py
│   ├── config.py
│   ├── models.py
│   │
│   ├── services/
│   │   ├── __init__.py
│   │   ├── pdf_parser.py
│   │   ├── excel_writer.py
│   │   ├── validator.py
│   │   └── processor.py
│   │
│   ├── web/
│   │   ├── templates/
│   │   │   ├── index.html
│   │   │   ├── result.html
│   │   │   └── error.html
│   │   └── static/
│   │       └── style.css
│   │
│   └── watcher.py
│
├── scripts/
│   └── build_clean_template.py
│
├── templates/
│   ├── reference_template.xlsx
│   └── po_template.xlsx
│
├── samples/
│   └── sample_po_027956.pdf
│
├── storage/
│   ├── incoming/
│   ├── processing/
│   ├── completed/
│   └── error/
│
├── tests/
│   ├── conftest.py
│   ├── test_numbers.py
│   ├── test_pdf_parser.py
│   ├── test_excel_writer.py
│   ├── test_end_to_end.py
│   └── test_web.py
│
├── .gitignore
├── requirements.txt
├── README.md
├── run_dev.ps1
└── run_tests.ps1
```

---

# 25. Dependências

Manter o MVP simples.

`requirements.txt` sugerido:

```text
fastapi
uvicorn[standard]
python-multipart
jinja2
pymupdf
openpyxl
pytest
httpx
```

Para watcher de pasta, se realmente necessário:

```text
watchdog
```

Não adicionar neste momento:

```text
Redis
Celery
RabbitMQ
React
Node
Docker Swarm
Kubernetes
OCR
LLM/API de IA
banco de dados
```

---

# 26. Interface web obrigatória

Criar FastAPI.

Rotas mínimas:

```text
GET  /
POST /process
GET  /download/{job_id}
GET  /health
```

## Tela inicial

Mostrar:

```text
Conversor Pedido TOTVS

[ Selecionar PDF ]

[ Processar ]
```

Aceitar somente `.pdf`.

Validar content type e assinatura do arquivo.

Tamanho máximo configurável, por exemplo:

```text
MAX_UPLOAD_MB=20
```

## Resultado

Mostrar:

```text
PO
Fornecedor
Quantidade de itens
Total das mercadorias
SCs encontrados
warnings
```

E:

```text
[ Baixar Excel ]
```

Não mostrar paths internos do servidor.

Usar UUID para jobs.

Exemplo:

```text
storage/jobs/<uuid>/
    input.pdf
    output.xlsx
```

Opcionalmente limpar jobs antigos por tempo.

---

# 27. Tratamento de erros

Criar exceções tipadas, por exemplo:

```python
class PurchaseOrderError(Exception): ...
class InvalidPdfError(PurchaseOrderError): ...
class UnsupportedLayoutError(PurchaseOrderError): ...
class PdfTextLayerNotFoundError(PurchaseOrderError): ...
class ExcelTemplateError(PurchaseOrderError): ...
class ValidationError(PurchaseOrderError): ...
```

Mensagens de UI em português.

Exemplos:

```text
O arquivo enviado não é um PDF válido.

Não foi possível identificar o layout esperado do relatório TOTVS.

O PDF não possui camada de texto suficiente para processamento.

O arquivo foi lido, mas nenhum item de pedido foi identificado.
```

Nunca retornar traceback na interface.

Logar traceback no servidor.

---

# 28. Validações antes de gerar o Excel

Bloquear se:

```text
po_number ausente
items vazio
product_code ausente
quantity inválida
unit_price inválido
total_value inválido
```

Warnings, sem alterar dados, para:

```text
ITEM_TOTAL_MISMATCH
SC_INCOMPLETE
DESCRIPTION_EMPTY
MERCHANDISE_TOTAL_MISMATCH
```

Validação do total:

```python
sum_items = sum(item.total_value for item in items)
```

No caso de referência:

```text
sum_items == Decimal("9361869.50")
merchandise_total == Decimal("9361869.50")
```

---

# 29. Validação do XLSX gerado

Após salvar, reabrir o output e verificar:

1. aba `Material for production` existe;
2. headers continuam presentes;
3. quantidade de linhas = quantidade de itens do PDF;
4. B contém sequência `1..N`;
5. C contém códigos do PDF;
6. D vazio;
7. E vazio;
8. F contém descrição do PDF;
9. M contém quantidade do PDF;
10. N contém preço do PDF;
11. O contém total do PDF;
12. P:S vazios;
13. nenhuma linha antiga do template permanece;
14. nenhuma fórmula externa existe no bloco dos itens;
15. total final confere;
16. arquivo abre normalmente com `openpyxl`.

---

# 30. Testes obrigatórios do parser

Criar pelo menos:

```python
def test_reference_pdf_header(): ...
def test_reference_pdf_has_193_items(): ...
def test_first_item(): ...
def test_last_item(): ...
def test_original_codes_are_preserved(): ...
def test_original_quantities_are_preserved(): ...
def test_merchandise_total(): ...
def test_sc_numbers_are_reconstructed(): ...
def test_ptbr_decimal_parser(): ...
```

Baseline:

```python
assert po.po_number == "027956"
assert po.supplier_name == "HEXING ELECTRICAL CO.,LTD HANGZHOU"
assert len(po.items) == 193
assert sum(i.quantity for i in po.items) == Decimal("53389684.000")
assert sum(i.total_value for i in po.items) == Decimal("9361869.50")
assert po.merchandise_total == Decimal("9361869.50")
```

Primeiro item:

```python
item = po.items[0]

assert item.product_code == "2010100040"
assert item.quantity == Decimal("75000.000")
assert item.unit_price == Decimal("0.0104160")
assert item.total_value == Decimal("781.20")
```

Último:

```python
item = po.items[-1]

assert item.product_code == "2150184130"
assert item.quantity == Decimal("652.000")
assert item.unit_price == Decimal("8.5302630")
assert item.total_value == Decimal("5561.73")
```

---

# 31. Testes obrigatórios do Excel

Gerar um arquivo temporário e verificar.

Exemplo conceitual:

```python
wb = load_workbook(output_path, data_only=False)
ws = wb["Material for production"]
```

Confirmar:

```text
A19 = 027956
B19 = 1
C19 = 2010100040
D19 = vazio
E19 = vazio
F19 começa com RESISTOR 390K
M19 = 75000
N19 = 0.0104160
O19 = 781.20
P19:S19 = vazio
```

Confirmar que a linha correspondente a `2010101620` contém esse código original.

Confirmar que a linha correspondente a `2020302563` contém `263800`.

Confirmar que a linha correspondente a `2054300520` contém `95497`.

Confirmar que nenhum item possui fórmula do tipo:

```text
VLOOKUP
XLOOKUP
INDEX
MATCH
```

no bloco de dados.

---

# 32. End-to-end acceptance test

Criar:

```python
def test_reference_pdf_to_excel_end_to_end():
    ...
```

Etapas:

```text
sample_po_027956.pdf
    ↓
process_purchase_order()
    ↓
arquivo temporário .xlsx
    ↓
reabrir
    ↓
comparar somente dados mapeados com o PDF
```

**Não comparar contra valores modificados manualmente no workbook de referência.**

O reference workbook é layout, não ground truth.

---

# 33. Web test

Usar `TestClient`.

Testar:

```text
GET / -> 200
GET /health -> 200
POST /process com TXT -> rejeitado
POST /process com PDF de referência -> sucesso
download gerado -> XLSX válido
```

---

# 34. Watcher de pasta - Fase 2

Depois que:

```text
parser
excel writer
end-to-end
web
```

estiverem estáveis, implementar a entrada por pasta de rede.

Estrutura:

```text
\\servidor\PO_AUTOMACAO\
    01_ENTRADA\
    02_PROCESSANDO\
    03_CONCLUIDOS\
    04_ERROS\
    05_ARQUIVADOS\
```

Ou equivalente configurável via ambiente.

Fluxo:

```text
PDF chega em 01_ENTRADA
       ↓
aguardar arquivo estabilizar
       ↓
mover para 02_PROCESSANDO
       ↓
process_purchase_order()
       ↓
03_CONCLUIDOS/PO_XXXXXX.xlsx
       ↓
original -> 05_ARQUIVADOS
```

Erro:

```text
PDF -> 04_ERROS
+ arquivo .txt ou .json com motivo
```

O watcher deve usar exatamente o mesmo `process_purchase_order()` da web.

Não duplicar parser/escritor.

---

# 35. Proteção contra arquivo ainda sendo copiado

No watcher, não processar imediatamente.

Verificar tamanho/modification time em pelo menos dois ciclos.

Exemplo conceitual:

```python
size1 = path.stat().st_size
await sleep(stability_seconds)
size2 = path.stat().st_size

if size1 != size2:
    continuar aguardando
```

Também tentar abrir o arquivo antes de mover.

---

# 36. Configuração

Usar variáveis de ambiente com defaults seguros:

```text
APP_HOST=127.0.0.1
APP_PORT=8000
MAX_UPLOAD_MB=20
TEMPLATE_PATH=templates/po_template.xlsx

WATCH_ENABLED=false
WATCH_INTERVAL_SECONDS=10
WATCH_STABILITY_SECONDS=5
WATCH_INCOMING=storage/incoming
WATCH_PROCESSING=storage/processing
WATCH_COMPLETED=storage/completed
WATCH_ERROR=storage/error
```

Não hardcode path de rede específico no código.

---

# 37. Logging

Usar `logging`.

Cada processamento deve registrar:

```text
job_id
nome do arquivo
PO
fornecedor
quantidade de páginas
quantidade de itens
total
tempo de processamento
warnings
status final
```

Nunca logar conteúdo binário.

Evitar dados desnecessários.

---

# 38. Segurança mínima

Para upload web:

- aceitar somente PDF;
- sanitizar filename;
- não usar filename do usuário como path direto;
- usar UUID;
- impedir `../`;
- limitar tamanho;
- não executar conteúdo do PDF;
- gerar output em diretório isolado;
- não expor paths internos.

Para download:

- validar `job_id`;
- garantir que arquivo pertence à pasta de jobs;
- retornar `404` se inexistente.

---

# 39. README obrigatório

Atualizar README com:

## Instalação Windows/VS Code

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## Criar template limpo

```powershell
python scripts/build_clean_template.py `
    --reference templates/reference_template.xlsx `
    --output templates/po_template.xlsx
```

## Testes

```powershell
pytest -q
```

## Servidor

```powershell
uvicorn app.main:app --reload
```

Abrir:

```text
http://127.0.0.1:8000
```

## Watcher

```powershell
python -m app.watcher
```

Somente quando `WATCH_ENABLED=true`.

---

# 40. Scripts PowerShell

`run_dev.ps1`:

```powershell
$ErrorActionPreference = "Stop"

if (-not (Test-Path ".venv")) {
    python -m venv .venv
}

& .\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
uvicorn app.main:app --reload
```

`run_tests.ps1`:

```powershell
$ErrorActionPreference = "Stop"
& .\.venv\Scripts\Activate.ps1
pytest -q
```

Ajustar se necessário.

---

# 41. Critérios de aceite do MVP

O MVP só está concluído quando TODOS os itens abaixo forem verdadeiros.

## Parser

- [ ] PDF 027956 identificado.
- [ ] 193 itens extraídos.
- [ ] ordem dos itens preservada.
- [ ] primeiro item correto.
- [ ] último item correto.
- [ ] preços corretos.
- [ ] quantidades corretas.
- [ ] totais corretos.
- [ ] SCs reconstruídos.
- [ ] nenhum código substituído.
- [ ] nenhuma quantidade ajustada.

## Excel

- [ ] estrutura visual baseada no workbook de referência.
- [ ] `Material for production` preservada.
- [ ] headers preservados.
- [ ] número de linhas dinâmico.
- [ ] dados antigos removidos.
- [ ] D/E/P/Q/R/S vazios.
- [ ] F preenchido pelo PDF.
- [ ] M preenchido pelo PDF.
- [ ] N preenchido pelo PDF.
- [ ] O preenchido pelo PDF.
- [ ] sem VLOOKUP no bloco de itens.
- [ ] sem `#NAME?`/`#REF!` no bloco de itens.
- [ ] total das mercadorias = 9.361.869,50 no exemplo.
- [ ] output abre normalmente.

## Web

- [ ] upload de PDF.
- [ ] processamento.
- [ ] página de resultado.
- [ ] download do XLSX.
- [ ] mensagens de erro compreensíveis.
- [ ] `/health`.

## Engenharia

- [ ] `pytest -q` verde.
- [ ] código tipado.
- [ ] `Decimal` para números.
- [ ] parser desacoplado de Excel.
- [ ] Excel desacoplado de PDF.
- [ ] serviço central reutilizável.
- [ ] README atualizado.
- [ ] nenhum dado específico da PO hardcoded na lógica de produção.

---

# 42. O que NÃO fazer

É proibido no MVP:

```text
inventar NCM
inventar Eletra Code
usar VLOOKUP externo
usar cadastro mestre
usar histórico de PO
substituir código
ajustar MOQ
arredondar quantidade por regra comercial
usar IA para adivinhar descrição
usar OCR sem necessidade
alterar preço do PDF
recalcular e substituir total do PDF
preencher Finalidade
preencher Centro de custos
preencher Requer LI?
preencher Projeto
```

Testes podem possuir os valores do PDF de referência como fixtures/asserts.

A lógica de produção não pode conter hacks como:

```python
if product_code == "2010101620":
    ...
```

ou:

```python
if po_number == "027956":
    ...
```

---

# 43. Ordem de implementação para o Codex

Execute nesta ordem.

## Etapa 1 - limpar escopo antigo

1. ler repositório;
2. rodar `pytest -q`;
3. identificar código de substituição/MOQ/master data;
4. remover do pipeline;
5. atualizar modelos.

## Etapa 2 - parser

1. implementar `parse_ptbr_decimal`;
2. implementar header parser;
3. implementar item detection;
4. implementar extração por coordenadas;
5. implementar descrição multilinha;
6. implementar SC quebrado;
7. implementar totais;
8. testes de regressão.

Não avance até os 193 itens estarem corretos.

## Etapa 3 - template limpo

1. criar `build_clean_template.py`;
2. remover dados do exemplo;
3. remover external formulas/links;
4. manter estilo/layout;
5. validar workbook.

## Etapa 4 - Excel writer

1. localizar table header;
2. redimensionar bloco;
3. recriar merges;
4. escrever valores;
5. escrever totais;
6. reabrir e validar.

## Etapa 5 - end-to-end

Executar:

```text
PDF -> XLSX -> reabrir -> validar
```

Não usar a planilha manual como truth de conteúdo.

## Etapa 6 - web

Criar FastAPI e interface simples.

## Etapa 7 - watcher

Somente depois do MVP web estar estável.

---

# 44. Comandos que o Codex deve executar

No início:

```powershell
pytest -q
```

Depois de cada etapa relevante:

```powershell
pytest -q
```

Ao final:

```powershell
pytest -q
python -m compileall app
```

Se houver linter configurado, executá-lo também.

Testar manualmente um processamento do PDF de referência e gerar:

```text
storage/completed/PO_027956.xlsx
```

ou arquivo temporário equivalente.

---

# 45. Relatório final que o Codex deve entregar

Ao concluir, responder com:

```text
1. Resumo da implementação
2. Arquivos criados
3. Arquivos removidos/alterados
4. Como executar
5. Resultado dos testes
6. Resultado do PDF de referência
7. Quantidade de itens extraídos
8. Total extraído
9. Validações do Excel
10. Limitações restantes
```

Informar explicitamente:

```text
PDF de referência: 193 itens
Total: 9.361.869,50
Substituições aplicadas: 0
Ajustes MOQ aplicados: 0
Campos extras preenchidos por inferência: 0
```

---

# 46. Definition of Done

Considerar o trabalho concluído somente quando este fluxo funcionar:

```text
Usuário acessa a página
        ↓
seleciona o PDF TOTVS
        ↓
clica em Processar
        ↓
parser identifica 193 itens no exemplo
        ↓
todos os dados vêm do PDF
        ↓
template é preenchido
        ↓
colunas sem origem ficam vazias
        ↓
XLSX é validado
        ↓
usuário baixa o arquivo
```

Resultado esperado:

```text
PDF = fonte da verdade
Excel = representação estruturada do PDF
Sem enriquecimento
Sem inferência
Sem substituição
Sem MOQ
Sem lookup externo
```

---

# 47. Instrução final ao Codex

**Implemente o sistema descrito neste arquivo. Não responda apenas com um plano.**

Antes de alterar:

1. leia todo o repositório relevante;
2. rode os testes;
3. preserve os arquivos de referência;
4. elimine as premissas antigas de substituição/MOQ;
5. implemente por etapas;
6. rode os testes continuamente.

Quando houver dúvida entre:

```text
usar um dado da planilha manual
```

e:

```text
deixar vazio
```

**deixe vazio**, a menos que exista uma correspondência direta e inequívoca no PDF.

Quando houver dúvida entre:

```text
corrigir um dado do PDF
```

e:

```text
preservar o PDF
```

**preserve o PDF**.

O objetivo deste MVP é fidelidade de transformação, não inteligência comercial.

---

# 48. Tratamento obrigatório de todo o bloco inferior da aba `Material for production`

> **Esta seção é autoritativa e complementa/sobrescreve qualquer instrução anterior que possa sugerir tratar somente a tabela de materiais.**
>
> A aplicação deve preservar e reposicionar **todo o conteúdo estrutural existente abaixo da tabela de itens**, e não apenas gerar os produtos e a linha de total.

No workbook de referência, a aba `Material for production` continua muito além da tabela de produtos.

No exemplo de referência, a estrutura observada é aproximadamente:

```text
linha 18       cabeçalho da tabela de materiais
linhas 19-211 itens do pedido (193 itens)
linha 212      total
linha 213      espaçamento / separação
linha 214+     bloco APPROVAL
linha 229+     bloco REMARKS
linhas 230+    instruções / informações operacionais
linhas 238+    tabela histórica de substituições
linhas 246+    tabela histórica de ajustes de quantidade / MOQ
linha 260      observação histórica de inclusão de item
```

Essas posições são válidas para o arquivo de referência e **não devem ser hardcoded como posições fixas para futuros pedidos**.

## 48.1 Regra geral

A estrutura final deve permanecer conceitualmente assim:

```text
CABEÇALHO DA PLANILHA
        ↓
TABELA DE MATERIAIS
        ↓
TOTAL
        ↓
APPROVAL
        ↓
REMARKS
        ↓
DEMAIS BLOCOS INFERIORES DO TEMPLATE
```

Quando o número de itens mudar, **todo o bloco abaixo dos itens deve ser deslocado junto**.

Exemplo:

### Arquivo de referência

```text
18    Cabeçalho
19    Item 1
...
211   Item 193
212   Total
213   Espaço
214   APPROVAL
...
229   REMARKS
...
260   Último bloco inferior
```

### Novo PDF com 100 itens

O resultado esperado deve ser aproximadamente:

```text
18    Cabeçalho
19    Item 1
...
118   Item 100
119   Total
120   Espaço
121   APPROVAL
...
136   REMARKS
...
```

Não deixar 93 linhas vazias apenas para manter `APPROVAL` na linha 214.

Também não sobrescrever nem perder os blocos inferiores.

---

# 49. Conteúdo inferior: estrutura versus dados específicos do exemplo

A planilha de referência contém dois tipos de conteúdo abaixo da tabela:

1. **estrutura/template que deve ser preservada**;
2. **dados específicos da PO manual usada como referência, que devem ser limpos**.

A aplicação deve distinguir os dois.

## 49.1 Blocos que devem permanecer estruturalmente

Preservar:

```text
APPROVAL
REMARKS
títulos
bordas
fills
fontes
merges
alinhamentos
larguras
alturas
áreas de texto
labels fixos
configurações de impressão
```

Esses elementos fazem parte do template.

## 49.2 Dados variáveis que não devem vazar do arquivo de referência

Remover/limpar valores específicos do exemplo, incluindo, quando existentes:

```text
PO027956
SC - 019347 - 019458
OCT/2026
29/06/2026
Code substitution and quantities update
registros de código antigo -> código novo
registros de Material Code / New Code
quantidades ajustadas
MOQ
observação "Included the item ..."
qualquer comentário específico da revisão manual
```

O princípio é:

> **Preservar o recipiente visual; limpar o conteúdo específico do pedido de referência.**

---

# 50. Preenchimento permitido nos blocos inferiores

Somente preencher automaticamente valores cuja origem seja **direta e inequívoca no PDF**.

Tabela autoritativa para o MVP:

| Informação no bloco inferior | Comportamento |
|---|---|
| `PO` | preencher com o número do pedido do PDF |
| `SC` | preencher com os SCs extraídos do PDF |
| `Payment Terms` | preencher com a condição de pagamento do PDF |
| mês de produção (`OCT/2026`, etc.) | deixar vazio |
| local de loading | deixar vazio, salvo se existir correspondência inequívoca no PDF |
| local de delivery operacional/importação | deixar vazio, salvo correspondência inequívoca |
| data de revisão manual | deixar vazio |
| texto de atualização de código/quantidade | deixar vazio |
| substituições de código | deixar linhas de dados vazias |
| tabela MOQ / Adjust Qty | deixar linhas de dados vazias |
| observações históricas de inclusão/exclusão | deixar vazio |

### Importante

O PDF de referência possui `Local de Entrega`, porém isso **não autoriza automaticamente** preencher um campo operacional de importação chamado `Place of Delivery`, `Port of Delivery` ou equivalente sem validar que o significado no template é exatamente o mesmo.

Na dúvida, deixar vazio.

---

# 51. PO e SC nos blocos inferiores

Se o template possuir uma linha destinada ao número da PO, gerar o valor a partir do PDF.

Exemplo:

```text
PO027956
```

pode ser gerado a partir de:

```text
po_number = "027956"
```

O formato deve respeitar o padrão visual do template.

Para SCs:

```text
019347
019458
```

o template pode receber, quando houver campo apropriado:

```text
SC - 019347 - 019458
```

Regras:

1. remover duplicados;
2. preservar zeros à esquerda;
3. preservar ordem de primeira ocorrência no PDF;
4. não ordenar numericamente se isso alterar a ordem do documento;
5. armazenar SC internamente como `str`.

Função sugerida:

```python
def unique_preserving_order(values: list[str]) -> list[str]:
    ...
```

---

# 52. Payment Terms no bloco inferior

O PDF de referência contém:

```text
Condicao de Pagto 170
240 DIAS
```

O valor textual de condição de pagamento disponível para apresentação é:

```text
240 DIAS
```

Se o template possuir algo como:

```text
Payment Terms:
```

é permitido preencher:

```text
Payment Terms: 240 DIAS
```

Não transformar em:

```text
Net 240
240 days after Acceptance
30 months
```

a menos que esses textos estejam literalmente presentes no PDF.

Texto manual presente no workbook, por exemplo:

```text
Net 240 days after Acceptance, but no later than...
```

deve ser removido se não vier do PDF.

---

# 53. Mês de produção

Não inferir:

```text
OCT/2026
```

a partir de:

```text
data de entrega
nome do arquivo
data do pedido
SC
observações
```

Mesmo que o nome original do PDF contenha algo como:

```text
SEA OCT.2026
```

o MVP definido neste documento trata **o conteúdo do PDF** como fonte de dados do pedido, não o filename como fonte de regra de negócio.

Portanto:

```text
Month of production = vazio
```

salvo futura decisão explícita de escopo.

---

# 54. Tabela de substituições no final da planilha

O workbook de referência possui uma tabela histórica semelhante a:

```text
PO
Old code
New code
substitute code
request Qty
change Qty
```

ou equivalente visual.

Essa tabela deve ser preservada como **estrutura visual**, se fizer parte do template, porém as linhas de dados devem ficar vazias no arquivo recém-gerado.

Exemplo:

```text
┌──────────┬────────────┬────────────┬───────────────┬─────────────┐
│ PO       │ Old code   │ New code   │ Request Qty   │ Change Qty  │
├──────────┼────────────┼────────────┼───────────────┼─────────────┤
│          │            │            │               │             │
│          │            │            │               │             │
│          │            │            │               │             │
└──────────┴────────────┴────────────┴───────────────┴─────────────┘
```

**Nunca preencher essa tabela a partir da diferença entre o PDF e a planilha manual de referência.**

O sistema não conhece substituições neste MVP.

---

# 55. Tabela MOQ / Adjust Qty no final da planilha

O workbook possui um bloco histórico semelhante a:

```text
PO
Material Code
New Code
Request Qty
Adjust Qty
MOQ
```

A estrutura pode permanecer para que a planilha mantenha o mesmo formato de envio.

Porém:

```text
Material Code = vazio
New Code      = vazio
Request Qty   = vazio
Adjust Qty    = vazio
MOQ           = vazio
```

nas linhas históricas do template.

Não usar os próprios itens do pedido para popular essa tabela.

Não calcular MOQ.

Não ajustar quantidade.

---

# 56. Observações históricas no final da planilha

Textos específicos como:

```text
29/06/2026 - Code substitution and quantities update.
```

ou:

```text
29/06/2026 - Included the item HXA...
```

devem ser removidos do template limpo.

Preservar apenas a área visual destinada a observações.

Se existir um título fixo:

```text
REMARKS
```

preservá-lo.

Se não houver observação direta e inequívoca a copiar do PDF:

```text
conteúdo de remarks = vazio
```

Não copiar automaticamente as `Observacoes` genéricas do PDF para o bloco `REMARKS` sem confirmar equivalência semântica.

---

# 57. `APPROVAL`

O bloco `APPROVAL` deve permanecer estruturalmente intacto.

Preservar:

```text
título
bordas
campos
merges
estilo
altura das linhas
layout
```

Valores preenchidos manualmente no arquivo de referência devem ser removidos se forem específicos daquele pedido.

O MVP não deve:

```text
aprovar automaticamente
preencher nome de aprovador
copiar aprovadores do PDF
copiar status BLQ/OK/REJ
```

sem mapeamento explícito definido neste documento.

Portanto, salvo valores realmente fixos do template:

```text
campos de aprovação = vazios
```

---

# 58. Algoritmo recomendado para deslocar o bloco inferior

Não tentar reconstruir manualmente todas as linhas inferiores.

Preferir preservar o bloco existente do template e deslocá-lo como unidade lógica.

Estratégia recomendada:

```text
1. identificar header_row
2. identificar item_start
3. identificar total_row_template
4. identificar approval_row_template
5. definir lower_block_start = primeira linha após a área variável de materiais
6. calcular item_count
7. calcular delta_rows em relação à capacidade do template
8. desfazer merges que cruzam a área dinâmica
9. inserir/deletar linhas na região imediatamente antes do bloco inferior
10. recriar merges dinâmicos da tabela de materiais
11. preencher os itens
12. limpar dados variáveis do bloco inferior
13. preencher somente PO/SC/Payment Terms quando houver campo confirmado
14. validar estilos/merges após o deslocamento
```

Exemplo de cálculo:

```python
reference_capacity = template_item_end - item_start + 1
delta = len(po.items) - reference_capacity

if delta > 0:
    ws.insert_rows(insert_at, amount=delta)
elif delta < 0:
    ws.delete_rows(delete_at, amount=abs(delta))
```

A implementação real deve considerar merges e fórmulas/formatos do workbook.

---

# 59. Identificação robusta dos blocos inferiores

Não depender somente do número da linha.

Localizar por labels.

Criar helpers, por exemplo:

```python
def find_row_by_value(ws, value: str) -> int | None:
    ...

def find_row_containing(ws, text: str) -> int | None:
    ...
```

Labels esperados:

```text
APPROVAL
REMARKS
```

Para os blocos de substituição/MOQ, identificar os headers visuais existentes no template.

Se o template esperado não contiver `APPROVAL` ou `REMARKS`, registrar warning/erro de template em vez de assumir linhas.

---

# 60. Construção do template limpo deve incluir o bloco inferior

Atualizar obrigatoriamente:

```text
scripts/build_clean_template.py
```

para limpar também as células abaixo da tabela de materiais.

O script deve produzir um template onde:

### Permanece

```text
APPROVAL
REMARKS
headers das tabelas inferiores
bordas
cores
merges
layout
labels institucionais realmente fixos
```

### É apagado

```text
PO específica
SC específica
mês específico
datas de revisão
substituições históricas
quantidades ajustadas
MOQ histórico
observações históricas
nomes/decisões específicas do pedido
valores originados por VLOOKUP
```

O teste do template deve confirmar que `027956` não permanece em áreas variáveis do workbook limpo.

Também deve confirmar que códigos históricos de substituição conhecidos do arquivo manual não permanecem no template limpo.

---

# 61. Testes obrigatórios do bloco inferior

Adicionar testes específicos.

Exemplos:

```python
def test_lower_sections_are_preserved(): ...
def test_approval_section_moves_with_item_count(): ...
def test_remarks_section_moves_with_item_count(): ...
def test_manual_substitution_rows_are_cleared(): ...
def test_manual_moq_rows_are_cleared(): ...
def test_manual_revision_notes_are_cleared(): ...
def test_po_is_written_to_confirmed_lower_field(): ...
def test_sc_is_written_to_confirmed_lower_field(): ...
def test_payment_terms_are_written_to_confirmed_lower_field(): ...
def test_production_month_is_blank(): ...
```

Criar pelo menos dois testes sintéticos de quantidade:

```text
100 itens
250 itens
```

e verificar que:

```text
APPROVAL continua depois do total
REMARKS continua depois de APPROVAL
nenhum bloco é sobrescrito
nenhuma linha histórica do exemplo reaparece
```

---

# 62. Critério adicional de aceite visual/estrutural

Além dos critérios anteriores, o MVP só estará concluído se:

- [ ] todo o bloco inferior da aba principal for preservado;
- [ ] `APPROVAL` permanecer depois do total;
- [ ] `REMARKS` permanecer depois de `APPROVAL`;
- [ ] o bloco inferior se mover conforme a quantidade de itens;
- [ ] nenhuma substituição histórica permanecer preenchida;
- [ ] nenhuma tabela MOQ permanecer com dados do exemplo;
- [ ] nenhuma observação de revisão específica permanecer;
- [ ] PO inferior, quando mapeada, venha do PDF;
- [ ] SC inferior, quando mapeado, venha do PDF;
- [ ] Payment Terms, quando mapeado, venha do PDF;
- [ ] mês de produção permaneça vazio;
- [ ] nenhuma informação operacional ambígua seja inferida;
- [ ] formatação/merges/bordas do bloco inferior sejam preservados.

---

# 63. Regra definitiva para o workbook inteiro

A lógica não deve ser:

```text
"copiar somente a tabela de itens"
```

A lógica correta é:

```text
Workbook de referência
        ↓
remover conteúdo variável/manual
        ↓
preservar estrutura visual COMPLETA
        ↓
redimensionar tabela de itens
        ↓
deslocar todo o bloco inferior
        ↓
preencher dados provenientes diretamente do PDF
        ↓
deixar todo o restante vazio
```

Em outras palavras:

> **A aplicação deve reproduzir o formato completo da planilha enviada, incluindo os blocos abaixo da tabela de materiais, mas nunca transportar para novos pedidos as decisões manuais, substituições, MOQ, comentários ou outros dados específicos do exemplo.**

