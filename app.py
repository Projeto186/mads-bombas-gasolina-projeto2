from flask import Flask, render_template, request, redirect, url_for, session
import base64
from io import BytesIO
import json
import os
import re
import time
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "mads_bombas_gasolina_2026_chave_segura")

# Cada página privada tem a sua própria chave
CHAVES_PRIVADAS = {
    "compras123": "compras",
    "avancado123": "avancado",
    "localizacoes123": "localizacoes",
    "integridade123": "integridade"
}

BASE_DIR = Path(__file__).resolve().parent
EXCEL_DB = BASE_DIR / "Base_Dados_Projeto2.xlsx"

# Para o Render ir buscar os dados ao SharePoint, define estas variáveis no Render:
# SHAREPOINT_EXCEL_URL = link de partilha do Excel no SharePoint
# MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET = credenciais Microsoft Graph
#
# Se o link do SharePoint for público/anónimo, podes usar só SHAREPOINT_EXCEL_URL.
# Se o ficheiro exigir login Microsoft, tens obrigatoriamente de usar Microsoft Graph.
SHAREPOINT_EXCEL_URL = os.environ.get("SHAREPOINT_EXCEL_URL", "https://ismaipt-my.sharepoint.com/:x:/g/personal/a044946_ipmaia_pt/IQDT1uQzM02ZQ41TmO4wCEVGATNiizUfEMOJ6bBHQGIOEAw?e=EHqrQw").strip()
MS_TENANT_ID = os.environ.get("MS_TENANT_ID", "").strip()
MS_CLIENT_ID = os.environ.get("MS_CLIENT_ID", "").strip()
MS_CLIENT_SECRET = os.environ.get("MS_CLIENT_SECRET", "").strip()
SHAREPOINT_CACHE_SECONDS = int(os.environ.get("SHAREPOINT_CACHE_SECONDS", "60"))

_excel_cache = {
    "created_at": 0,
    "content": None
}


def limpar_valor_excel(valor):
    """Converte valores vindos do Excel para formatos simples usados pelos templates."""
    if pd.isna(valor):
        return ""
    if isinstance(valor, pd.Timestamp):
        return valor.strftime("%Y-%m-%d")
    return valor


def validar_conteudo_excel(conteudo):
    # Ficheiros .xlsx são ficheiros ZIP e começam normalmente por PK.
    # Se o SharePoint devolver uma página de login HTML, isto evita um erro confuso no pandas.
    if not conteudo or not conteudo[:2] == b"PK":
        inicio = conteudo[:120].decode("utf-8", errors="ignore") if conteudo else ""
        raise RuntimeError(
            "O SharePoint não devolveu um ficheiro Excel válido. "
            "Provavelmente devolveu uma página de login/HTML. "
            "Se o ficheiro não for público, tens de configurar Microsoft Graph no Render. "
            f"Início da resposta recebida: {inicio}"
        )


def criar_share_id(share_url):
    encoded = base64.urlsafe_b64encode(share_url.encode("utf-8")).decode("utf-8")
    encoded = encoded.rstrip("=")
    return "u!" + encoded


def obter_token_graph():
    if not (MS_TENANT_ID and MS_CLIENT_ID and MS_CLIENT_SECRET):
        raise RuntimeError(
            "Faltam credenciais Microsoft Graph. Define MS_TENANT_ID, "
            "MS_CLIENT_ID e MS_CLIENT_SECRET no Render."
        )

    token_url = f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token"
    resposta = requests.post(
        token_url,
        data={
            "client_id": MS_CLIENT_ID,
            "client_secret": MS_CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials"
        },
        timeout=30
    )

    if not resposta.ok:
        raise RuntimeError(
            "Não foi possível obter token do Microsoft Graph. "
            f"Status {resposta.status_code}: {resposta.text[:500]}"
        )

    return resposta.json()["access_token"]


def descarregar_excel_por_graph():
    token = obter_token_graph()
    share_id = criar_share_id(SHAREPOINT_EXCEL_URL)
    url = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem/content"

    resposta = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        allow_redirects=True,
        timeout=60
    )

    if not resposta.ok:
        raise RuntimeError(
            "Não foi possível descarregar o Excel pelo Microsoft Graph. "
            f"Status {resposta.status_code}: {resposta.text[:500]}"
        )

    validar_conteudo_excel(resposta.content)
    return resposta.content


def descarregar_excel_por_link_publico():
    url = SHAREPOINT_EXCEL_URL
    if "download=1" not in url:
        separador = "&" if "?" in url else "?"
        url = f"{url}{separador}download=1"

    resposta = requests.get(url, allow_redirects=True, timeout=60)

    if not resposta.ok:
        raise RuntimeError(
            "Não foi possível descarregar o Excel pelo link público do SharePoint. "
            f"Status {resposta.status_code}: {resposta.text[:500]}"
        )

    validar_conteudo_excel(resposta.content)
    return resposta.content


def obter_excel_bytes():
    agora = time.time()

    if _excel_cache["content"] is not None and (agora - _excel_cache["created_at"] < SHAREPOINT_CACHE_SECONDS):
        return _excel_cache["content"]

    if SHAREPOINT_EXCEL_URL:
        if MS_TENANT_ID and MS_CLIENT_ID and MS_CLIENT_SECRET:
            conteudo = descarregar_excel_por_graph()
        else:
            conteudo = descarregar_excel_por_link_publico()

        _excel_cache["content"] = conteudo
        _excel_cache["created_at"] = agora
        return conteudo

    if not EXCEL_DB.exists():
        raise FileNotFoundError(
            f"Ficheiro Excel local não encontrado: {EXCEL_DB}. "
            "Ou adiciona o ficheiro localmente, ou define SHAREPOINT_EXCEL_URL no Render."
        )

    return EXCEL_DB.read_bytes()


def ler_excel(nome_folha):
    conteudo_excel = obter_excel_bytes()
    tabela = pd.read_excel(BytesIO(conteudo_excel), sheet_name=nome_folha, engine="openpyxl")
    tabela = tabela.applymap(limpar_valor_excel)
    return tabela.to_dict(orient="records")


def ler_compras():
    return ler_excel("Compras")


def ler_localizacoes():
    return ler_excel("Localizações")


def numero(valor, defeito=0):
    try:
        return float(str(valor).replace(",", "."))
    except (ValueError, TypeError):
        return defeito


def data_para_ordenar(valor):
    try:
        return datetime.strptime(str(valor)[:10], "%Y-%m-%d")
    except ValueError:
        return datetime.min


def compras_sem_nif(compras):
    resultado = []
    for compra in compras:
        nova = dict(compra)
        nova.pop("NIF", None)
        # Removido da página pública, como tinhas pedido antes
        nova.pop("Metodo_Pagamento", None)
        return_nome = nova
        resultado.append(return_nome)
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
            "data": data
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
            "data_preco": info_preco.get("data", "")
        })
    return pontos


def dados_grafico_precos(compras):
    # Média diária por tipo de combustível
    agrupado = defaultdict(list)
    for compra in compras:
        data = compra.get("Data", "")
        combustivel = compra.get("Tipo_Combustivel", "")
        preco = numero(compra.get("Preço_Litro"))
        if data and combustivel and preco:
            agrupado[(data, combustivel)].append(preco)

    datas = sorted({data for data, _ in agrupado.keys()})
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
            "media": round(valores["total"] / valores["litros"], 3) if valores["litros"] else 0
        })

    tabela_combustivel = []
    for combustivel, valores in por_combustivel.items():
        tabela_combustivel.append({
            "combustivel": combustivel,
            "compras": valores["compras"],
            "litros": round(valores["litros"], 2),
            "total": round(valores["total"], 2),
            "media": round(valores["total"] / valores["litros"], 3) if valores["litros"] else 0
        })

    tabela_marca = sorted(tabela_marca, key=lambda x: x["total"], reverse=True)
    tabela_combustivel = sorted(tabela_combustivel, key=lambda x: x["total"], reverse=True)

    grafico_marcas = {
        "labels": [linha["marca"] for linha in tabela_marca],
        "data": [linha["total"] for linha in tabela_marca]
    }

    grafico_combustiveis = {
        "labels": [linha["combustivel"] for linha in tabela_combustivel],
        "data": [linha["litros"] for linha in tabela_combustivel]
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
    # Para este projeto valida-se a estrutura: exatamente 9 dígitos.
    return bool(re.fullmatch(r"\d{9}", texto(nif)))


def adicionar_erro(erros, tipo, linha, problema, motivo):
    erros.append({
        "tipo": tipo,
        "linha": linha,
        "problema": problema,
        "motivo": motivo
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
                adicionar_erro(
                    erros,
                    "Compra",
                    indice,
                    f"Valor total incoerente: {total}",
                    f"O total deveria ser aproximadamente {total_esperado}, porque preço x litros = total."
                )

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


@app.route("/")
def home():
    compras = ler_compras()[:300]
    localizacoes = ler_localizacoes()
    mapa = dados_mapa(localizacoes, compras)
    grafico_precos = dados_grafico_precos(compras)

    return render_template(
        "home.html",
        compras=compras_sem_nif(compras),
        mapa_json=json.dumps(mapa, ensure_ascii=False),
        grafico_precos_json=json.dumps(grafico_precos, ensure_ascii=False)
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

    compras = ler_compras()[:300]
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
        grafico_precos_json=json.dumps(grafico_precos, ensure_ascii=False)
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
        marcas=marcas
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
        total_localizacoes=total_localizacoes
    )


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("home"))


if __name__ == "__main__":
    app.run(debug=True)
