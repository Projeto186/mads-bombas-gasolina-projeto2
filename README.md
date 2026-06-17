# Projeto Flask - Bombas de Gasolina

Versão simples para Visual Studio Code.

## Como executar

1. Abrir esta pasta no Visual Studio Code.
2. Abrir o terminal.
3. Executar:

```bash
python app.py
```

4. Abrir no browser:

```text
http://127.0.0.1:5000
```

## Páginas

- `/` - página pública com mapa, gráfico e 300 compras sem NIF.
- `/acesso` - página para inserir chave privada.
- `/privado/compras` - compras com NIF.
- `/privado/avancado` - gráficos e tabelas avançadas.
- `/privado/localizacoes` - lista completa de localizações e mapa com legenda.
- `/privado/integridade` - verificação da integridade dos dados.

## Chaves privadas

- `compras123` - permite entrar em `/privado/compras`.
- `avancado123` - permite entrar em `/privado/avancado`.
- `localizacoes123` - permite entrar em `/privado/localizacoes`.
- `integridade123` - permite entrar em `/privado/integridade`.

Cada chave só dá acesso à sua própria página privada.

## Base de dados

O ficheiro Excel completo enviado foi incluído na pasta como:

`Base_Dados_Projeto2_Melhorada_Localizacoes.xlsx`

Para facilitar a execução sem instalar bibliotecas extra, as folhas do Excel foram convertidas para:

- `compras.csv`
- `localizacoes.csv`

A aplicação Flask lê estes dois ficheiros CSV.
"# mads-bombas-gasolina-projeto2"

## Deploy no Render

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

A aplicação tenta ler primeiro `Base_Dados_Projeto2.xlsx`. Se esse ficheiro não estiver disponível no deploy, usa automaticamente os ficheiros `compras.csv` e `localizacoes.csv`.

