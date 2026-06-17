from flask import Flask, render_template, request, redirect, url_for, session, g, has_request_context
import html
from io import BytesIO
import json
import os
import re
import time
import hashlib
from collections import defaultdict
from datetime import datetime
from urllib.parse import urlparse, quote

import pandas as pd
import requests

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "mads_bombas_gasolina_2026_chave_segura")

# Cada página privada tem a sua própria chave
CHAVES_PRIVADAS = {
    "compras123": "compras",
    "avancado123": "avancado",
    "localizacoes123": "localizacoes",
    "integridade123": "integridade",
}

# Fonte de dados: Google Sheets
# Podes substituir por variável de ambiente no Render, mas este link já fica definido por defeito.
GOOGLE_SHEET_URL = os.environ.get(
    "GOOGLE_SHEET_URL",
    os.environ.get(
        "GOOGLE_SHEET_ID",
        "https://docs.google.com/spreadsheets/d/189SiAMZfhSN-VXUazREAoStx-sEMSdVBso3s3wWccj0/edit?usp=sharing",
    ),
)

GOOGLE_COMPRAS_SHEET = os.environ.get("GOOGLE_COMPRAS_SHEET", "Compras")
GOOGLE_LOCALIZACOES_SHEET = os.environ.get("GOOGLE_LOCALIZACOES_SHEET", "Localizações")

HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0 Safari/537.36"
    ),
    "Accept": "text/csv,text/plain,*/*",
    "Cache-Control": "no-cache, no-store, must-revalidate",
    "Pragma": "no-cache",
}


@app.after_request
def impedir_cache_browser(response):
    # Impede o browser/CDN de mostrar HTML antigo quando fazes refresh.
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


@app.errorhandler(Exception)
def mostrar_erro_visivel(erro):
    mensagem = str(erro)
    return f"""
    <!doctype html>
    <html lang="pt">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Erro ao carregar dados</title>
        <style>
            body {{ font-family: Arial, sans-serif; background:#f4f6f8; color:#111827; padding:40px; }}
            .box {{ max-width:900px; margin:auto; background:white; border-radius:12px; padding:24px; box-shadow:0 2px 12px rgba(0,0,0,.12); }}
            pre {{ white-space:pre-wrap; background:#f3f4f6; padding:16px; border-radius:8px; overflow:auto; }}
            a {{ color:#2563eb; }}
        </style>
    </head>
    <body>
        <div class="box">
            <h1>Erro ao carregar dados do Google Sheets</h1>
            <p>A aplicação está online, mas não conseguiu obter ou ler os dados do Google Sheets.</p>
            <pre>{html.escape(mensagem)}</pre>
            <p>Confirma que o Google Sheets está partilhado como "Anyone with the link can view" e que as abas se chamam Compras e Localizações.</p>
            <p><a href="/">Tentar novamente</a></p>
        </div>
    </body>
    </html>
    """, 500


def extrair_google_sheet_id(valor):
    """Aceita link completo ou apenas o ID do Google Sheets."""
    valor = str(valor or "").strip()
    if not valor:
        raise RuntimeError("GOOGLE_SHEET_URL/GOOGLE_SHEET_ID está vazio.")

    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", valor)
    if match:
        return match.group(1)

    # Se não for link, assume que já é o ID.
    if re.fullmatch(r"[a-zA-Z0-9-_]+", valor):
        return valor

    raise RuntimeError("Não foi possível extrair o ID do Google Sheets do valor configurado.")


def construir_url_csv(nome_aba):
    sheet_id = extrair_google_sheet_id(GOOGLE_SHEET_URL)
    cache_buster = str(int(time.time() * 1000))
    return (
        f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq"
        f"?tqx=out:csv&sheet={quote(nome_aba)}&_cb={cache_buster}"
    )


def pedir_url(url):
    headers = dict(HTTP_HEADERS)
    agora = str(int(time.time() * 1000))
    headers["X-Cache-Buster"] = agora
    headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    headers["Pragma"] = "no-cache"
    headers["Expires"] = "0"
    headers["If-Modified-Since"] = "Thu, 01 Jan 1970 00:00:00 GMT"
    return requests.get(url, headers=headers, allow_redirects=True, timeout=15)


def descarregar_csv_google_sheets(nome_aba, com_info=False):
    url = construir_url_csv(nome_aba)
    resposta = pedir_url(url)
    content_type = resposta.headers.get("Content-Type", "")

    if not resposta.ok:
        inicio = resposta.text[:300] if resposta.text else ""
        raise RuntimeError(
            f"Falha ao descarregar a aba {nome_aba}: status {resposta.status_code}, "
            f"content-type {content_type}, início da resposta: {inicio}"
        )

    conteudo = resposta.content or b""
    texto_inicio = conteudo[:300].decode("utf-8", errors="ignore").lower()

    # Quando a folha não está pública, o Google costuma devolver HTML em vez de CSV.
    if "<html" in texto_inicio or "<!doctype" in texto_inicio:
        raise RuntimeError(
            f"A aba {nome_aba} devolveu HTML em vez de CSV. "
            "Provavelmente o Google Sheets não está partilhado publicamente com permissão de visualização."
        )

    info = {
        "url": url,
        "content_type": content_type,
        "tamanho": str(len(conteudo)),
        "sha256": hashlib.sha256(conteudo).hexdigest(),
        "hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    return (conteudo, info) if com_info else conteudo


def limpar_valor(valor):
    """Converte valores vindos do Google Sheets para formatos simples usados pelos templates."""
    if pd.isna(valor):
        return ""
    if isinstance(valor, pd.Timestamp):
        return valor.strftime("%Y-%m-%d")
    return valor


def ler_google_sheet(nome_aba):
    # Dentro do mesmo request guardamos em g só para não descarregar a mesma aba duas vezes.
    # Em cada refresh/pedido novo, volta a pedir os dados ao Google Sheets.
    if has_request_context():
        if not hasattr(g, "google_sheet_tables_request"):
            g.google_sheet_tables_request = {}
        if nome_aba not in g.google_sheet_tables_request:
            conteudo = descarregar_csv_google_sheets(nome_aba)
            tabela = pd.read_csv(BytesIO(conteudo))
            try:
                tabela = tabela.map(limpar_valor)
            except AttributeError:
                tabela = tabela.applymap(limpar_valor)
            g.google_sheet_tables_request[nome_aba] = tabela.to_dict(orient="records")
        return g.google_sheet_tables_request[nome_aba]

    conteudo = descarregar_csv_google_sheets(nome_aba)
    tabela = pd.read_csv(BytesIO(conteudo))
    try:
        tabela = tabela.map(limpar_valor)
    except AttributeError:
        tabela = tabela.applymap(limpar_valor)
    return tabela.to_dict(orient="records")


def ler_compras():
    return ler_google_sheet(GOOGLE_COMPRAS_SHEET)


def ler_localizacoes():
    return ler_google_sheet(GOOGLE_LOCALIZACOES_SHEET)


def numero(valor, defeito=0):
    try:
        return float(str(valor).replace(",", "."))
    except (ValueError, TypeError):
        return defeito


def data_para_ordenar(valor):
    valor = str(valor or "").strip()
    formatos = ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"]
    for formato in formatos:
        try:
            return datetime.strptime(valor[:10], formato)
        except ValueError:
            continue
    return datetime.min


def compras_sem_nif(compras):
    resultado = []
    for compra in compras:
        nova = dict(compra)
        nova.pop("NIF", None)
        nova.pop("Metodo_Pagamento", None)
        resultado.append(nova)
    return resultado


def ultimos_precos_por_marca(compras):
    ultimos = {}
    compras_ordenadas = sorted(compras, key=lambda c: data_para_ordenar(c.get("Data")))
    for compra in compras_ordenadas:
        marca = compra.get("Marca", "Sem marca")
        combustivel = compra.get("Tipo_Combustivel", "combustível")
        preco = compra.get("Preço_Litro", "")
        data = compra.get("Data", "")
        ultimos[marca] = {
            "combustivel": combustivel,
            "preco": preco,
            "data": data,
        }
    return ultimos


def dados_mapa(localizacoes, compras):
    ultimos = ultimos_precos_por_marca(compras)
    pontos = []
    for loc in localizacoes:
        marca = loc.get("Marca", "Sem marca")
        info_preco = ultimos.get(marca, {})
        pontos.append({
            "id": loc.get("LocalizacaoID", ""),
            "marca": marca,
            "nome": loc.get("Nome_Posto", ""),
            "lat": numero(loc.get("Latitude")),
            "lng": numero(loc.get("Longitude")),
            "cor": loc.get("Cor_Mapa", "#3388ff"),
            "concelho": loc.get("Concelho", ""),
            "preco": info_preco.get("preco", "Sem registo"),
            "combustivel": info_preco.get("combustivel", ""),
            "data_preco": info_preco.get("data", ""),
        })
    return pontos


def dados_grafico_precos(compras):
    agrupado = defaultdict(list)
    for compra in compras:
        data = compra.get("Data", "")
        combustivel = compra.get("Tipo_Combustivel", "")
        preco = numero(compra.get("Preço_Litro"))
        if data and combustivel and preco:
            agrupado[(data, combustivel)].append(preco)

    datas = sorted({data for data, _ in agrupado.keys()}, key=data_para_ordenar)
    combustiveis = sorted({comb for _, comb in agrupado.keys()})

    datasets = []
    for combustivel in combustiveis:
        valores = []
        for data in datas:
            lista = agrupado.get((data, combustivel), [])
            valores.append(round(sum(lista) / len(lista), 3) if lista else None)
        datasets.append({"label": combustivel, "data": valores})

    return {"labels": datas, "datasets": datasets}


def estatisticas_avancadas(compras):
    por_marca = defaultdict(lambda: {"compras": 0, "litros": 0, "total": 0})
    por_combustivel = defaultdict(lambda: {"compras": 0, "litros": 0, "total": 0})

    for compra in compras:
        marca = compra.get("Marca", "Sem marca")
        combustivel = compra.get("Tipo_Combustivel", "Sem combustível")
        litros = numero(compra.get("Quantidade_Litros"))
        total = numero(compra.get("Valor_Total"))

        por_marca[marca]["compras"] += 1
        por_marca[marca]["litros"] += litros
        por_marca[marca]["total"] += total

        por_combustivel[combustivel]["compras"] += 1
        por_combustivel[combustivel]["litros"] += litros
        por_combustivel[combustivel]["total"] += total

    tabela_marca = []
    for marca, valores in por_marca.items():
        tabela_marca.append({
            "marca": marca,
            "compras": valores["compras"],
            "litros": round(valores["litros"], 2),
            "total": round(valores["total"], 2),
            "media": round(valores["total"] / valores["litros"], 3) if valores["litros"] else 0,
        })

    tabela_combustivel = []
    for combustivel, valores in por_combustivel.items():
        tabela_combustivel.append({
            "combustivel": combustivel,
            "compras": valores["compras"],
            "litros": round(valores["litros"], 2),
            "total": round(valores["total"], 2),
            "media": round(valores["total"] / valores["litros"], 3) if valores["litros"] else 0,
        })

    tabela_marca = sorted(tabela_marca, key=lambda x: x["total"], reverse=True)
    tabela_combustivel = sorted(tabela_combustivel, key=lambda x: x["total"], reverse=True)

    grafico_marcas = {
        "labels": [linha["marca"] for linha in tabela_marca],
        "data": [linha["total"] for linha in tabela_marca],
    }

    grafico_combustiveis = {
        "labels": [linha["combustivel"] for linha in tabela_combustivel],
        "data": [linha["litros"] for linha in tabela_combustivel],
    }

    return tabela_marca, tabela_combustivel, grafico_marcas, grafico_combustiveis


def verificar_acesso(tipo_necessario):
    return session.get("tipo_acesso") == tipo_necessario


def texto(valor):
    return str(valor or "").strip()


def obter_coluna(linha, nomes):
    for nome in nomes:
        if nome in linha:
            return texto(linha.get(nome))
    return ""


def data_valida_e_nao_futura(valor):
    valor = texto(valor)
    if not valor:
        return False, "A data está vazia."

    formatos = ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"]
    for formato in formatos:
        try:
            data = datetime.strptime(valor[:10], formato).date()
            if data > datetime.now().date():
                return False, "A data está no futuro."
            return True, ""
        except ValueError:
            continue

    return False, "A data não está num formato reconhecido."


def nif_bem_estruturado(nif):
    return bool(re.fullmatch(r"\d{9}", texto(nif)))


def adicionar_erro(erros, tipo, linha, problema, motivo):
    erros.append({
        "tipo": tipo,
        "linha": linha,
        "problema": problema,
        "motivo": motivo,
    })


def verificar_integridade_dados():
    erros = []
    compras = ler_compras()
    localizacoes = ler_localizacoes()

    marcas_validas = {texto(loc.get("Marca")).lower() for loc in localizacoes if texto(loc.get("Marca"))}
    nomes_postos_validos = {texto(loc.get("Nome_Posto")).lower() for loc in localizacoes if texto(loc.get("Nome_Posto"))}
    ids_compras = set()
    ids_localizacoes = set()

    for indice, compra in enumerate(compras, start=2):
        id_compra = obter_coluna(compra, ["CompraID", "ID_Compra", "ID", "Id"])
        marca = obter_coluna(compra, ["Marca"])
        nome_posto = obter_coluna(compra, ["Nome_Posto", "Nome_Bomba", "Bomba", "Posto"])
        nif = obter_coluna(compra, ["NIF", "NIF_Cliente", "Nif"])
        data = obter_coluna(compra, ["Data", "Data_Compra", "Data Compra"])
        preco = obter_coluna(compra, ["Preço_Litro", "Preco_Litro", "Preco", "Preço"])
        litros = obter_coluna(compra, ["Quantidade_Litros", "Litros", "Quantidade"])
        total = obter_coluna(compra, ["Valor_Total", "Total"])

        if not id_compra:
            adicionar_erro(erros, "Compra", indice, "Compra sem ID", "Cada compra deve ter um identificador.")
        elif id_compra in ids_compras:
            adicionar_erro(erros, "Compra", indice, f"ID de compra duplicado: {id_compra}", "O mesmo ID não deve aparecer em mais do que uma compra.")
        else:
            ids_compras.add(id_compra)

        if not marca:
            adicionar_erro(erros, "Compra", indice, "Marca vazia", "A compra tem de indicar a marca da bomba.")
        elif marcas_validas and marca.lower() not in marcas_validas:
            adicionar_erro(erros, "Compra", indice, f"Marca/bomba não registada: {marca}", "A marca da compra não existe no ficheiro de localizações, por isso pode ter sido inventada ou escrita de forma diferente.")

        if nome_posto and nomes_postos_validos and nome_posto.lower() not in nomes_postos_validos:
            adicionar_erro(erros, "Compra", indice, f"Nome de bomba não registado: {nome_posto}", "O nome da bomba não existe na lista de localizações registadas.")

        if not nif:
            adicionar_erro(erros, "Compra", indice, "NIF vazio", "O NIF deve existir nas compras privadas.")
        elif not nif_bem_estruturado(nif):
            adicionar_erro(erros, "Compra", indice, f"NIF mal estruturado: {nif}", "O NIF deve ter exatamente 9 dígitos numéricos.")

        data_ok, motivo_data = data_valida_e_nao_futura(data)
        if not data_ok:
            adicionar_erro(erros, "Compra", indice, f"Data inválida: {data}", motivo_data)

        preco_num = numero(preco, None)
        litros_num = numero(litros, None)
        total_num = numero(total, None)

        if preco_num is None or preco_num <= 0:
            adicionar_erro(erros, "Compra", indice, f"Preço por litro inválido: {preco}", "O preço por litro deve ser um número maior do que zero.")

        if litros_num is None or litros_num <= 0:
            adicionar_erro(erros, "Compra", indice, f"Quantidade de litros inválida: {litros}", "A quantidade de litros deve ser um número maior do que zero.")

        if total_num is None or total_num <= 0:
            adicionar_erro(erros, "Compra", indice, f"Valor total inválido: {total}", "O valor total deve ser um número maior do que zero.")

        if preco_num and litros_num and total_num:
            total_esperado = round(preco_num * litros_num, 2)
            if abs(total_esperado - total_num) > 0.05:
                adicionar_erro(erros, "Compra", indice, f"Valor total incoerente: {total}", f"O total deveria ser aproximadamente {total_esperado}, porque preço x litros = total.")

    for indice, loc in enumerate(localizacoes, start=2):
        id_loc = obter_coluna(loc, ["LocalizacaoID", "ID", "Id"])
        marca = obter_coluna(loc, ["Marca"])
        nome = obter_coluna(loc, ["Nome_Posto", "Nome", "Posto"])
        lat = obter_coluna(loc, ["Latitude", "Lat"])
        lng = obter_coluna(loc, ["Longitude", "Lng", "Lon"])
        cor = obter_coluna(loc, ["Cor_Mapa", "Cor"])

        if not id_loc:
            adicionar_erro(erros, "Localização", indice, "Localização sem ID", "Cada localização deve ter um identificador.")
        elif id_loc in ids_localizacoes:
            adicionar_erro(erros, "Localização", indice, f"ID de localização duplicado: {id_loc}", "O mesmo ID não deve aparecer em mais do que uma localização.")
        else:
            ids_localizacoes.add(id_loc)

        if not marca:
            adicionar_erro(erros, "Localização", indice, "Marca vazia", "Cada localização deve indicar a marca da bomba.")
        if not nome:
            adicionar_erro(erros, "Localização", indice, "Nome do posto vazio", "Cada localização deve indicar o nome da bomba/posto.")

        lat_num = numero(lat, None)
        lng_num = numero(lng, None)
        if lat_num is None or not (36 <= lat_num <= 43):
            adicionar_erro(erros, "Localização", indice, f"Latitude inválida: {lat}", "A latitude deve ser numérica e estar dentro de uma zona plausível para Portugal continental.")
        if lng_num is None or not (-10 <= lng_num <= -6):
            adicionar_erro(erros, "Localização", indice, f"Longitude inválida: {lng}", "A longitude deve ser numérica e estar dentro de uma zona plausível para Portugal continental.")

        if cor and not re.fullmatch(r"#[0-9A-Fa-f]{6}", cor):
            adicionar_erro(erros, "Localização", indice, f"Cor inválida: {cor}", "A cor do mapa deve estar no formato hexadecimal, por exemplo #FF0000.")

    return erros, len(compras), len(localizacoes)


@app.route("/debug-google-sheets")
def debug_google_sheets():
    try:
        compras_csv, compras_info = descarregar_csv_google_sheets(GOOGLE_COMPRAS_SHEET, com_info=True)
        localizacoes_csv, localizacoes_info = descarregar_csv_google_sheets(GOOGLE_LOCALIZACOES_SHEET, com_info=True)

        compras_preview = pd.read_csv(BytesIO(compras_csv))
        localizacoes_preview = pd.read_csv(BytesIO(localizacoes_csv))

        return (
            "OK - Google Sheets descarregado AGORA pelo servidor.<br>"
            f"Hora do servidor: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}<br>"
            f"Aba compras: {html.escape(GOOGLE_COMPRAS_SHEET)}<br>"
            f"Linhas compras: {len(compras_preview)}<br>"
            f"SHA256 compras: {compras_info['sha256']}<br>"
            f"Aba localizações: {html.escape(GOOGLE_LOCALIZACOES_SHEET)}<br>"
            f"Linhas localizações: {len(localizacoes_preview)}<br>"
            f"SHA256 localizações: {localizacoes_info['sha256']}<br>"
        )
    except Exception as erro:
        return f"ERRO GOOGLE SHEETS: {html.escape(str(erro))}", 500


@app.route("/")
def home():
    compras_todas = ler_compras()
    localizacoes = ler_localizacoes()

    mapa = dados_mapa(localizacoes, compras_todas)
    grafico_precos = dados_grafico_precos(compras_todas)

    compras_visiveis = compras_todas[:300]

    return render_template(
        "home.html",
        compras=compras_sem_nif(compras_visiveis),
        mapa_json=json.dumps(mapa, ensure_ascii=False),
        grafico_precos_json=json.dumps(grafico_precos, ensure_ascii=False),
    )


@app.route("/acesso", methods=["GET", "POST"])
def acesso():
    erro = None

    if request.method == "POST":
        chave = request.form.get("chave", "").strip()

        if chave in CHAVES_PRIVADAS:
            tipo_acesso = CHAVES_PRIVADAS[chave]
            session["tipo_acesso"] = tipo_acesso

            if tipo_acesso == "compras":
                return redirect(url_for("compras_privadas"))
            if tipo_acesso == "avancado":
                return redirect(url_for("avancado"))
            if tipo_acesso == "localizacoes":
                return redirect(url_for("localizacoes_privadas"))
            if tipo_acesso == "integridade":
                return redirect(url_for("integridade"))

        erro = "Chave inválida."

    return render_template("acesso.html", erro=erro)


@app.route("/privado/compras")
def compras_privadas():
    if not verificar_acesso("compras"):
        return redirect(url_for("acesso"))

    compras = ler_compras()
    return render_template("compras_privadas.html", compras=compras)


@app.route("/privado/avancado")
def avancado():
    if not verificar_acesso("avancado"):
        return redirect(url_for("acesso"))

    compras = ler_compras()
    tabela_marca, tabela_combustivel, grafico_marcas, grafico_combustiveis = estatisticas_avancadas(compras)
    grafico_precos = dados_grafico_precos(compras)

    return render_template(
        "avancado.html",
        tabela_marca=tabela_marca,
        tabela_combustivel=tabela_combustivel,
        grafico_marcas_json=json.dumps(grafico_marcas, ensure_ascii=False),
        grafico_combustiveis_json=json.dumps(grafico_combustiveis, ensure_ascii=False),
        grafico_precos_json=json.dumps(grafico_precos, ensure_ascii=False),
    )


@app.route("/privado/localizacoes")
def localizacoes_privadas():
    if not verificar_acesso("localizacoes"):
        return redirect(url_for("acesso"))

    compras = ler_compras()
    localizacoes = ler_localizacoes()
    mapa = dados_mapa(localizacoes, compras)

    marcas = {}
    for ponto in mapa:
        marcas[ponto["marca"]] = ponto["cor"]

    return render_template(
        "localizacoes.html",
        localizacoes=mapa,
        mapa_json=json.dumps(mapa, ensure_ascii=False),
        marcas=marcas,
    )


@app.route("/privado/integridade")
def integridade():
    if not verificar_acesso("integridade"):
        return redirect(url_for("acesso"))

    erros, total_compras, total_localizacoes = verificar_integridade_dados()

    return render_template(
        "integridade.html",
        erros=erros,
        total_erros=len(erros),
        total_compras=total_compras,
        total_localizacoes=total_localizacoes,
    )


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("home"))


if __name__ == "__main__":
    app.run(debug=True)
