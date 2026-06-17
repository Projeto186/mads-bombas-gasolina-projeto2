# Projeto Flask - Bombas de Gasolina

README_SHAREPOINT_ONLY_20260617

Aplicação Flask para visualização e análise de dados de bombas de gasolina, preparada para deploy no Render.

A aplicação usa **apenas um ficheiro Excel alojado no SharePoint** como fonte de dados.

Não usa ficheiros locais como:

- `Base_Dados_Projeto2.xlsx`
- `Base_Dados_Projeto2_Melhorada_Localizacoes.xlsx`
- `compras.csv`
- `localizacoes.csv`

## 1) Fonte de dados

A fonte de dados é definida através da variável de ambiente:

```text
SHAREPOINT_EXCEL_URL
```

Esta variável deve conter o link de partilha do ficheiro Excel no SharePoint.

Se o ficheiro estiver protegido e não puder ser lido diretamente pelo link público, também devem ser configuradas estas variáveis no Render:

```text
MS_TENANT_ID
MS_CLIENT_ID
MS_CLIENT_SECRET
```

## 2) Variáveis de ambiente no Render

No Render, abrir o serviço da aplicação e configurar em **Environment**:

```text
SHAREPOINT_EXCEL_URL=<link_do_excel_no_sharepoint>
FLASK_SECRET_KEY=<uma_chave_segura>
```

Opcionalmente, para ficheiros privados no SharePoint:

```text
MS_TENANT_ID=<tenant_id>
MS_CLIENT_ID=<client_id>
MS_CLIENT_SECRET=<client_secret>
```

Também é possível definir o tempo de cache do Excel:

```text
SHAREPOINT_CACHE_SECONDS=60
```

## 3) Como executar localmente

Instalar as dependências:

```bash
pip install -r requirements.txt
```

Definir a variável de ambiente com o link do Excel.

No PowerShell:

```powershell
$env:SHAREPOINT_EXCEL_URL="https://link-do-ficheiro-sharepoint"
python app.py
```

Depois abrir no browser:

```text
http://127.0.0.1:5000
```

## 4) Deploy no Render

Build Command:

```bash
pip install -r requirements.txt
```

Start Command:

```bash
gunicorn app:app
```

O ficheiro `Procfile` também inclui este arranque:

```text
web: gunicorn app:app
```

Depois de fazer push para o GitHub, no Render fazer:

```text
Manual Deploy > Deploy latest commit
```

## 5) Páginas da aplicação

- `/` — página pública com mapa, gráfico e compras sem NIF.
- `/acesso` — página para inserir chave privada.
- `/privado/compras` — compras com NIF.
- `/privado/avancado` — gráficos e tabelas avançadas.
- `/privado/localizacoes` — lista completa de localizações e mapa com legenda.
- `/privado/integridade` — verificação da integridade dos dados.

## 6) Chaves privadas

- `compras123` — permite entrar em `/privado/compras`.
- `avancado123` — permite entrar em `/privado/avancado`.
- `localizacoes123` — permite entrar em `/privado/localizacoes`.
- `integridade123` — permite entrar em `/privado/integridade`.

Cada chave só dá acesso à sua própria página privada.

## 7) Notas importantes

Este projeto está configurado para funcionar com dados vindos do SharePoint.

Se `SHAREPOINT_EXCEL_URL` não estiver configurada corretamente, a aplicação pode falhar no arranque ou ao carregar os dados.

Os ficheiros CSV locais e ficheiros Excel locais não fazem parte do fluxo atual da aplicação e não devem ser usados como fallback.
