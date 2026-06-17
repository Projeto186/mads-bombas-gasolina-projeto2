# MADS - Bombas de Gasolina

Aplicação Flask para visualizar e analisar dados de compras de combustível e localizações de postos, usando Google Sheets como fonte de dados.

## Fonte de dados

A aplicação lê os dados diretamente a partir de um Google Sheets.

Google Sheets configurado:

```txt
https://docs.google.com/spreadsheets/d/189SiAMZfhSN-VXUazREAoStx-sEMSdVBso3s3wWccj0/edit?usp=sharing
```

O ficheiro deve ter, pelo menos, estas duas folhas:

```txt
Compras
Localizações
```

## Permissões do Google Sheets

No Google Sheets, confirma que o ficheiro está partilhado como:

```txt
Anyone with the link can view
```

Ou em português:

```txt
Qualquer pessoa com o link pode ver
```

Sem esta permissão, o Render não consegue descarregar os dados.

## Funcionamento da atualização

Quando alteras dados no Google Sheets:

```txt
Alterar Google Sheets
↓
Guardar/sincronizar automaticamente
↓
Fazer refresh no site
↓
A página mostra os dados novos
```

Não é necessário fazer commit para atualizar os dados.

O commit só é necessário quando alteras código.

## Variáveis de ambiente no Render

O `app.py` já tem o link do Google Sheets configurado por defeito.

Mesmo assim, podes configurar no Render se quiseres controlar tudo por variáveis:

```txt
GOOGLE_SHEET_ID=https://docs.google.com/spreadsheets/d/189SiAMZfhSN-VXUazREAoStx-sEMSdVBso3s3wWccj0/edit?usp=sharing
GOOGLE_COMPRAS_SHEET=Compras
GOOGLE_LOCALIZACOES_SHEET=Localizações
FLASK_SECRET_KEY=uma_chave_segura_qualquer
```

Se não colocares `GOOGLE_SHEET_ID`, a aplicação usa o link que já está escrito no código.

## Teste de ligação ao Google Sheets

Depois do deploy, abre:

```txt
/debug-google-sheets
```

Se estiver tudo correto, deve aparecer uma mensagem a indicar que o Google Sheets foi lido pelo servidor.

Se der erro, verifica:

1. se o link do Google Sheets está correto;
2. se o ficheiro está público para leitura;
3. se as folhas se chamam exatamente `Compras` e `Localizações`;
4. se as colunas mantêm os nomes esperados.

## Rotas principais

```txt
/                       Página pública
/acesso                 Página para inserir chave privada
/privado/compras        Compras privadas
/privado/avancado       Estatísticas avançadas
/privado/localizacoes   Localizações privadas
/privado/integridade    Verificação de integridade dos dados
/debug-google-sheets    Teste técnico da ligação ao Google Sheets
/logout                 Terminar sessão
```

## Chaves privadas

As páginas privadas usam estas chaves:

```txt
compras123       → compras
avancado123      → avançado
localizacoes123  → localizações
integridade123   → integridade
```

## Deploy no Render

Depois de substituir o `app.py` e o `README.md`, faz:

```bash
git add app.py README.md
git commit -m "Atualizar projeto para Google Sheets"
git push origin main
```

O Render deve fazer novo deploy automaticamente.

## Dependências

O projeto precisa destas dependências principais:

```txt
Flask
pandas
requests
gunicorn
```

Exemplo de `requirements.txt`:

```txt
Flask
pandas
requests
gunicorn
```

## Comando de arranque no Render

No Render, o comando de start deve ser:

```bash
gunicorn app:app
```

## Notas importantes

- A página pública limita a tabela a 300 compras para não ficar pesada.
- Os cálculos do mapa e dos gráficos devem usar todas as compras.
- A página privada de compras deve mostrar todas as compras.
- O browser recebe cabeçalhos para evitar cache de HTML antigo.
- Sempre que fizeres refresh, o servidor tenta ler novamente os dados do Google Sheets.
