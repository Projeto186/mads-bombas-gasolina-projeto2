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

## Chaves privadas

- `compras123` - permite entrar em `/privado/compras`.
- `avancado123` - permite entrar em `/privado/avancado`.
- `localizacoes123` - permite entrar em `/privado/localizacoes`.

Cada chave só dá acesso à sua própria página privada.

## Base de dados

O ficheiro Excel completo enviado foi incluído na pasta como:

`Base_Dados_Projeto2_Melhorada_Localizacoes.xlsx`

Para facilitar a execução sem instalar bibliotecas extra, as folhas do Excel foram convertidas para:

- `compras.csv`
- `localizacoes.csv`

A aplicação Flask lê estes dois ficheiros CSV.
