from flask import Flask, render_template, request, redirect, url_for, session
import csv
import json
from collections import defaultdict
from datetime import datetime

app = Flask(__name__)
app.secret_key = "trocar_esta_chave_secreta"

# Cada página privada tem a sua própria chave
CHAVES_PRIVADAS = {
    "compras123": "compras",
    "avancado123": "avancado",
    "localizacoes123": "localizacoes"
}

COMPRAS_CSV = "compras.csv"
LOCALIZACOES_CSV = "localizacoes.csv"


def ler_csv(caminho):
    with open(caminho, encoding="utf-8-sig", newline="") as ficheiro:
        return list(csv.DictReader(ficheiro))


def ler_compras():
    return ler_csv(COMPRAS_CSV)


def ler_localizacoes():
    return ler_csv(LOCALIZACOES_CSV)


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


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("home"))


if __name__ == "__main__":
    app.run(debug=True)
