# App Docs Product

Aplicacao Flask para gerar documentos `.docx` a partir de um template de proposta de OS.

## Tecnologias

- Python 3.11+
- Flask
- python-docx
- python-dotenv
- Flask-SQLAlchemy
- Flask-Migrate
- Marshmallow

## Estrutura principal

```text
app_docs_product/
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
- Os arquivos gerados sao salvos em `UPLOAD_FOLDER` (pasta `uploads/`).

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

## Funcionalidades atuais

- Pagina inicial em `/`
- Geracao de Proposta de OS em `/documentos/proposta_os`
- Download de arquivo gerado em `/documentos/download/<filename>`
- Pagina de relatorio de entrega em `/documentos/relatorio_entrega`

## Fluxo de geracao de documento

1. Usuario preenche o formulario de Proposta de OS.
2. Frontend envia os dados em JSON para `POST /documentos/proposta_os`.
3. O servico `DocxService`:
   - carrega o template DOCX configurado;
   - substitui placeholders por dados do formulario;
   - calcula valor com base em `qtd_hst * 200`;
   - salva o arquivo final em `uploads/`.
4. A API retorna o nome do arquivo para download.

## Observacoes de desenvolvimento

- A estrutura ja usa o padrao Application Factory (`create_app`).
- Dependencias de banco/migracao estao instaladas, mas ainda nao ha pasta de migracoes nem modelos persistentes implementados no estado atual.
