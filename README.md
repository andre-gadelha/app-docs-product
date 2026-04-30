# App Docs Product

Aplicacao Flask para gerar documentos `.docx` a partir de templates de proposta de OS e relatorio de entrega.

## Tecnologias

- Python 3.11+
- Flask
- python-docx
- PyMuPDF
- python-dotenv
- Flask-SQLAlchemy
- Flask-Migrate
- Marshmallow

## Estrutura principal

```text
app_docs_product/
  templates_docx/
    proposta_os/
    relatorio_entrega/
  app/
    routes/
    services/
    templates/
    __init__.py
  uploads/
  config.py
  run.py
  requirements.txt
  .env
```

## Requisitos

- Python instalado (recomendado 3.11 ou superior)
- `pip`
- Template DOCX no caminho configurado em `config.py`

## Instalacao

1. Criar e ativar ambiente virtual:

```bash
python -m venv .venv
```

No Windows (PowerShell):

```bash
.venv\Scripts\Activate.ps1
```

2. Instalar dependencias:

```bash
pip install -r requirements.txt
```

## Configuracao de ambiente

O projeto carrega variaveis a partir do arquivo `.env`.
Se o arquivo nao existir, crie manualmente um `.env` na raiz do projeto antes de executar a aplicacao.

Exemplo minimo:

```env
FLASK_APP=run.py
FLASK_DEBUG=1
SECRET_KEY=sua-chave-secreta
# Opcional:
# DATABASE_URL=sqlite:///app.db
```

Observacoes importantes:

- Se `DATABASE_URL` nao for definida, o projeto usa SQLite local em `app.db`.
- O caminho do template DOCX e definido por `TEMPLATE_DOCX` em `config.py`.
- O template de relatorio de entrega e definido por `TEMPLATE_RELATORIO_ENTREGA_DOCX`.
- Os arquivos gerados sao salvos em `UPLOAD_FOLDER` (pasta `uploads/`) com separacao por `propostas/`, `hus/` e `relatorios/`.

## Executando o projeto

Opcao 1 (direto com Python):

```bash
python run.py
```

Opcao 2 (com Flask CLI):

```bash
flask run
```

A aplicacao sobe, por padrao, em:

- [http://127.0.0.1:5000](http://127.0.0.1:5000)

## Docker

Build da imagem:

```bash
docker build -t app-docs-product:latest .
```

Build com versionamento por commit curto (recomendado):

- O commit curto e os 7 primeiros caracteres do commit Git atual (ex.: `590f7e6`).
- Isso ajuda a rastrear exatamente qual codigo gerou a imagem.

No Windows (PowerShell):

```powershell
$sha = git rev-parse --short HEAD
docker build -t app-docs-product:1.0.0-$sha -t app-docs-product:latest .
```

Executar container:

```bash
docker run --rm -p 5000:5000 --env-file .env app-docs-product:latest
```

Com volume para persistir os arquivos gerados em `uploads/`:

```bash
docker run --rm -p 5000:5000 --env-file .env -v ${PWD}/uploads:/app/uploads app-docs-product:latest
```

## Funcionalidades atuais

- Pagina inicial em `/`
- Geracao de Proposta de OS em `/documentos/proposta_os`
- Geracao de Relatorio de Entrega em `/documentos/relatorio_entrega`
- Download de arquivo gerado em `/documentos/download/<filename>`

## Fluxo de geracao de documento

1. Usuario preenche um formulario de Proposta de OS ou Relatorio de Entrega.
2. Frontend envia dados para os endpoints:
   - `POST /documentos/proposta_os` (JSON)
   - `POST /documentos/relatorio_entrega` (multipart com arquivos)
3. O servico `DocxService`:
   - carrega o template DOCX configurado para cada funcionalidade;
   - substitui placeholders por dados do formulario;
   - para relatorio de entrega, replica a introducao da proposta e inclui HUs (lista e conteudo detalhado no item 9);
   - salva o arquivo final em `uploads/`.
4. A API retorna o nome do arquivo para download.

## Observacoes de desenvolvimento

- A estrutura ja usa o padrao Application Factory (`create_app`).
- Dependencias de banco/migracao estao instaladas, mas ainda nao ha pasta de migracoes nem modelos persistentes implementados no estado atual.
