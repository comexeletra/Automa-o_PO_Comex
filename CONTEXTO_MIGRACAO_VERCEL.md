# Contexto para migração do conversor TOTVS para Vercel

Atualizado em 2026-09-14. Este arquivo é uma passagem de contexto para uma nova janela de trabalho. **A migração ainda não foi implementada nem publicada.** O usuário já autorizou reduzir os tamanhos aceitos para a Vercel, e o código local foi ajustado para isso. Reconfirme limites, preços e instruções de deploy nas fontes oficiais antes de tomar decisões.

## Objetivo e decisão tomada sobre tamanho

Migrar o aplicativo Python/FastAPI que converte PDFs de pedidos TOTVS em planilhas Excel de Cloudflare Workers para **Vercel Functions com runtime Python**, preservando as rotas, a interface e o resultado do Excel. A Vercel é tecnicamente plausível para eliminar o limite de CPU já identificado no Worker gratuito, mas **não há garantia de funcionamento completo sem um teste de ponta a ponta**.

**Decisão do usuário:** PDFs e planilhas grandes são improváveis; reduzir os dois limites para **4 MB** e manter o fluxo de envio/processamento/download direto, sem armazenamento de objetos nesta fase. O código local agora usa `MAX_UPLOAD_MB=4` e `MAX_OUTPUT_MB=4` por padrão, informa o limite na interface, rejeita PDFs maiores antes do envio no navegador e no servidor, e rejeita planilhas geradas maiores com uma página de erro. O limite oficial da Vercel continua sendo **4,5 MB para o corpo da requisição e para o da resposta**, inclusive no Pro; os 4 MB deixam margem para o multipart. O PDF de referência tem **93.221 bytes**. Se surgir necessidade real de arquivos maiores, reabrir a decisão arquitetural.

Se for uso empresarial/comercial, não assumir que o plano gratuito Hobby é permitido; a Vercel o reserva para uso pessoal não comercial. O Pro estava anunciado a partir de US$ 20/mês na data acima, sujeito a uso adicional. **Não contratar plano nem publicar nada sem autorização.**

## Diagnóstico confirmado no Cloudflare

- URL de produção observada: `convesorcomexpo.renatojanericomexeletraenergy.workers.dev`. O navegador exibia **Cloudflare Error 1105 — Temporarily unavailable** e o painel de Observability não mostrava eventos ao abrir a página.
- Foi criado um Worker temporário de diagnóstico, `convesorcomexpo-diagnostico`, com uma cópia do bundle de produção. A importação de `main.py` falhou com `introspection.CpuLimitExceeded: Python Worker exceeded CPU time limit`, chegando à importação de `Jinja2Templates`/Jinja2. Importações isoladas de dependências passaram. Isso comprova o bloqueio de CPU na inicialização; não comprova que a conversão completa funcionaria após remover esse bloqueio.
- A tentativa de configurar `limits.cpu_ms = 30000` no diagnóstico foi recusada pela API da Cloudflare com erro **100328**: limites de CPU personalizados não são suportados no plano Workers Free.
- O Worker de diagnóstico foi removido e a existência apenas do Worker de produção foi verificada. O Worker de produção não foi alterado nesse diagnóstico. Nenhuma conversão PDF→Excel foi validada remotamente depois disso.
- Limites documentados do Workers Free: **10 ms de CPU por requisição** e **128 MB de memória**. Workers Paid: **30 s de CPU por padrão**, configurável até **5 min**, mas mantém **128 MB de memória**. Portanto, Workers Paid é uma alternativa de menor mudança; a migração para Vercel ainda é uma escolha a avaliar, não uma exigência técnica já demonstrada.

## Estado do repositório

Raiz: `C:\05_automação_formularios\01_piloto`.

- `app/main.py`: instância `app = FastAPI(...)`; rotas `GET /`, `GET /health`, `GET /result` e `POST /process`. Usa Jinja2, `StaticFiles`, `UploadFile`, `TemporaryDirectory` e devolve o `.xlsx` diretamente no corpo da resposta. No final registra `Default = asgi.entrypoint(app)` somente se o pacote `workers` estiver disponível.
- `app/config.py`: `MAX_UPLOAD_MB` e `MAX_OUTPUT_MB` padrão **4** cada; template em `app/resources/po_template.xlsx`. Os arquivos de entrada/saída de uma conversão são temporários. Não configurar esses limites acima de 4 MB em uma Vercel Function com o fluxo HTTP atual.
- `app/services/pdf_parser.py`, `validator.py`, `excel_writer.py`, `processor.py`: lógica de negócio a preservar. Bibliotecas de runtime no `pyproject.toml`: FastAPI, Jinja2, openpyxl, pypdf e python-multipart.
- `app/web/templates/` e `app/web/static/`: páginas HTML e CSS. `app/resources/` contém o template `.xlsx` e a imagem usada no Excel/interface; garantir que entrem no pacote de deploy.
- `wrangler.jsonc`, `app/worker.py`, `pylock.toml` e configuração `workers` são específicos do Cloudflare. Não apagar antes de existir uma alternativa testada; manter possibilidade de retorno.
- `pyproject.toml` declara Python `>=3.12` e grupo de desenvolvimento com `workers-py`/`workers-runtime-sdk`. A Vercel aceita `pyproject.toml` com ou sem `uv.lock`; verificar o pacote de produção sem dependências desnecessárias.
- `samples/sample_po_027956.pdf` tem **93.221 bytes**. `app/resources/po_template.xlsx` tem **901.626 bytes**.
- Testes existentes: `tests/test_web.py` cobre home, health, rejeição de não PDF, os novos limites de entrada/saída de 4 MB e o download real do PDF de referência abaixo do limite; `tests/test_end_to_end.py` confere a conversão do pedido 027956 e campos/estrutura do Excel. Há ainda testes de parser, números e writer. `python -B -m pytest -q -p no:cacheprovider` passou localmente em 2026-09-14 (**13 testes**, 2 avisos de depreciação em dependências). Isso não substitui o teste na Vercel.
- `mcp_cloudflare.md` já era um arquivo não rastreado antes deste trabalho; preservá-lo. Verificar o `git status` atualizado na próxima janela antes de qualquer edição.
- `AGENTS.md` determina usar as skills e integração Cloudflare para qualquer tarefa de Cloudflare e validar a configuração/pedir confirmação antes de mudança remota relevante ou irreversível.

## O que muda na Vercel

1. **Runtime correto:** usar Vercel Functions no **runtime Python**, não o Edge Runtime, já que o aplicativo usa FastAPI e arquivos temporários. A documentação atual da Vercel reconhece `app/main.py` com variável `app` como entrypoint. Ainda assim, validar a detecção no build; se necessário, declarar `[tool.vercel] entrypoint = "app.main:app"` no `pyproject.toml`.
2. **Adaptador Cloudflare:** isolar/remover do caminho Vercel a importação `workers.asgi` e o registro `Default`, preservando o entrypoint Cloudflare até a decisão de desligá-lo. A lógica de serviço deve permanecer compartilhada.
3. **Arquivos e templates:** confirmar que `app/web/templates`, CSS, logo e `app/resources/po_template.xlsx` estão no bundle. A documentação Vercel informa suporte a `StaticFiles`/`app.mount()` para assets; testar as rotas `/static` e `/assets` em preview. Se necessário, usar configuração de inclusão de arquivos; não mover o template Excel para uma pasta pública por engano.
4. **Temporários:** o filesystem da função é somente leitura, exceto `/tmp` com até **500 MB** de espaço temporário. A conversão já usa `TemporaryDirectory`; confirmar o diretório efetivo, o tamanho de pico e limpeza. Não tratar `/tmp` como armazenamento persistente.
5. **Limites:** com Fluid Compute, a documentação consultada informava até **300 s de duração** e **2 GB de memória** no Hobby; Pro tem limite maior configurável. Isso é duração total, não a mesma métrica de CPU ativa usada pela Cloudflare. Bundle Python padrão: até **500 MB descompactado**. Verificar os números novamente na implementação.
6. **Payload de 4,5 MB:** o usuário aprovou o teto operacional de **4 MB por PDF e por Excel**; manter esses limites ao migrar. Medir PDF de entrada **mais multipart** e Excel de saída no preview. Para voltar a suportar arquivos maiores, redesenhar upload/download com armazenamento de objetos (por exemplo Vercel Blob ou outro serviço), mas isso não faz parte da migração inicial.
7. **Segurança/privacidade:** PDFs de pedidos e planilhas podem conter dados comerciais. Se usar armazenamento externo, definir acesso privado, expiração, retenção e exclusão. Não registrar conteúdo dos arquivos nos logs.

## Plano de verificação para a próxima janela

1. Ler `AGENTS.md`, este documento, `README.md`, `app/main.py`, `app/config.py`, `pyproject.toml` e os testes. Verificar o estado do git e preservar alterações do usuário.
2. Usar a decisão já tomada de **4 MB por PDF/Excel** para a primeira versão. Confirmar se o uso é comercial antes de escolher plano; **não publicar nem contratar plano por suposição**.
3. Implementar a adaptação **local** da Vercel com alterações mínimas na lógica de conversão, mantendo o deploy Cloudflare reversível.
4. Rodar `pytest -q` e testar localmente `GET /`, `GET /health`, `/static`, `/assets`, envio de PDF inválido e conversão real do PDF de referência. Conferir que a planilha abre e preserva os campos esperados.
5. Validar build/preview da Vercel e fazer o mesmo teste de ponta a ponta no ambiente de preview, observando logs, tempo, memória, tamanho dos corpos e bundle. Deploy remoto e cobrança exigem autorização apropriada.
6. Só considerar troca da URL/produção depois da validação do preview; manter rollback para o Worker atual até o aceite do usuário.

## Fontes oficiais consultadas em 2026-09-14

- Cloudflare Workers, limites: https://developers.cloudflare.com/workers/platform/limits/
- Vercel, FastAPI e entrypoints: https://vercel.com/docs/frameworks/backend/fastapi
- Vercel, runtime Python e dependências: https://vercel.com/docs/functions/runtimes/python
- Vercel Functions, duração, memória, bundle e payload: https://vercel.com/docs/functions/limitations
- Vercel Functions, `/tmp`: https://vercel.com/docs/functions/runtimes
- Vercel, planos e restrição Hobby: https://vercel.com/pricing
