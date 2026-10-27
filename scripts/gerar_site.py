#!/usr/bin/env python3
"""Gera o site de preço do m² por bairro a partir das planilhas de ITBI.

Uso:
  python scripts/gerar_site.py                           baixa as planilhas dos Releases
  python scripts/gerar_site.py --arquivos a.xlsx b.xlsx  usa arquivos locais
"""
import argparse
import json
import re
import shutil
import unicodedata
import urllib.request
from html import escape
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
MODELO = RAIZ / "scripts" / "modelo_bairro.html"
DADOS = RAIZ / "dados"
SITE = RAIZ / "site"

RELEASES = {
    "2025": "https://github.com/jlgrego/itbi-sp/releases/download/dados-2025/itbi_2025.xlsx",
    "2026": "https://github.com/jlgrego/itbi-sp/releases/download/dados-2026/itbi_2026.xlsx",
}

MIN_TRANSACOES = 30  # bairros com menos registros não ganham página

# A planilha traz os nomes sem acento. Acrescente aqui os bairros que quiser com acento.
NOMES = {"TATUAPE": "Tatuapé", "SAUDE": "Saúde"}

ABREV = {
    "R": "Rua", "AV": "Avenida", "PC": "Praça", "ES": "Estrada", "AL": "Alameda",
    "TV": "Travessa", "LG": "Largo", "STA": "Santa", "STO": "Santo", "PROF": "Professor",
    "DR": "Doutor", "CAP": "Capitão", "GEN": "General", "CEL": "Coronel", "EN": "Estrada",
    "PDE": "Padre",
}

# Como reconhecer cada coluna pelo nome do cabeçalho (sem acento, minúsculas, sem símbolos)
BUSCAS = {
    "cadastro": lambda c: "cadastro" in c and "sql" in c and "situacao" not in c,
    "logradouro": lambda c: "logradouro" in c,
    "numero": lambda c: c == "numero",
    "bairro": lambda c: c == "bairro",
    "natureza": lambda c: c.startswith("natureza"),
    "valor": lambda c: "valordetransacao" in c,
    "data": lambda c: "datadetransacao" in c,
    "situacao": lambda c: c.startswith("situacao"),
    "area": lambda c: c.startswith("areaconstruida"),
    "referencia": lambda c: c == "referencia",
}
OPCIONAIS = {"referencia"}


def chave(texto_):
    s = unicodedata.normalize("NFD", str(texto_)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]", "", s)


def mapear(colunas):
    mapa = {}
    for col in colunas:
        c = chave(col)
        for campo, regra in BUSCAS.items():
            if campo not in mapa and regra(c):
                mapa[campo] = col
    return mapa


def numero(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return float("nan")
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if "," in s:  # formato brasileiro: 428.832,00
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return float("nan")


def texto(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def expande(logradouro):
    return " ".join(ABREV.get(p, p).title() if p in ABREV else p.title() for p in logradouro.split())


def baixar(ano, url):
    DADOS.mkdir(exist_ok=True)
    destino = DADOS / f"itbi_{ano}.xlsx"
    print(f"Baixando {ano} do Release...")
    req = urllib.request.Request(url, headers={"User-Agent": "itbi-sp-gerador"})
    with urllib.request.urlopen(req, timeout=600) as resp, open(destino, "wb") as f:
        shutil.copyfileobj(resp, f)
    tamanho = destino.stat().st_size / 1_000_000
    if tamanho < 1:
        raise SystemExit(f"Arquivo de {ano} muito pequeno ({tamanho:.2f} MB). Verifique o Release.")
    return destino


def ler(caminho):
    print(f"Lendo {caminho.name}")
    partes = []
    xls = pd.ExcelFile(caminho)
    for aba in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=aba, dtype=object)
        mapa = mapear(df.columns)
        faltando = [c for c in BUSCAS if c not in mapa and c not in OPCIONAIS]
        if faltando:
            print(f"  aba {aba}: ignorada, colunas não encontradas: {faltando}")
            continue
        col = lambda campo: df[mapa[campo]]  # noqa: E731
        out = pd.DataFrame({
            "cadastro": col("cadastro").map(texto),
            "logradouro": col("logradouro").map(texto),
            "numero": col("numero").map(texto),
            "bairro": col("bairro").map(texto).str.upper(),
            "natureza": col("natureza").map(texto),
            "situacao": col("situacao").map(texto),
            "valor": col("valor").map(numero),
            "data": pd.to_datetime(col("data"), errors="coerce", dayfirst=True, format="mixed"),
            "area": col("area").map(numero),
        })
        print(f"  aba {aba}: {len(out)} linhas")
        partes.append(out)
    return pd.concat(partes, ignore_index=True)


def aplicar_regras(df):
    print(f"\nTotal de linhas lidas: {len(df)}")
    df = df[df.natureza.str.startswith("1.")]
    print(f"Compra e venda: {len(df)}")
    df = df[df.situacao.str.startswith("Ativo")]
    print(f"Imóveis ativos: {len(df)}")
    df = df[df.valor.between(1, 1e9) & df.data.notna()]
    print(f"Com valor e data válidos: {len(df)}")
    df = df[df.cadastro != ""].sort_values("data").drop_duplicates("cadastro", keep="last")
    print(f"Um registro por cadastro (o mais recente): {len(df)}")
    df = df[df.situacao.str.startswith("Ativo Predial")]
    print(f"Prediais (apartamentos e casas): {len(df)}")
    df = df[df.area.between(20, 1000)]
    print(f"Área entre 20 e 1.000 m²: {len(df)}")
    df = df.assign(preco_m2=df.valor / df.area)
    df = df[df.preco_m2.between(500, 100000)]
    print(f"R$/m² entre 500 e 100 mil: {len(df)}")
    return df


def serie_de(df):
    out = []
    for (ano, mes), g in df.groupby([df.data.dt.year, df.data.dt.month]):
        g = g.sort_values("data", ascending=False)
        out.append({
            "mes": f"{mes:02d}/{str(ano)[2:]}",
            "n": int(len(g)),
            "valor": round(float(g.valor.median())),
            "m2": round(float(g.preco_m2.median())),
            "lista": [{
                "rua": expande(r.logradouro), "numero": r.numero,
                "data": r.data.strftime("%d/%m/%Y"), "valor": round(r.valor),
                "m2": round(r.preco_m2), "cadastro": r.cadastro,
            } for r in g.itertuples()],
        })
    return out


def dados_do_bairro(df, nome):
    bairro = {
        "nome": nome,
        "n": int(len(df)),
        "mediana": round(float(df.preco_m2.median())),
        "p25": round(float(df.preco_m2.quantile(0.25))),
        "p75": round(float(df.preco_m2.quantile(0.75))),
        "valor_mediano": round(float(df.valor.median())),
        "inicio": df.data.min().strftime("%d/%m/%Y"),
        "fim": df.data.max().strftime("%d/%m/%Y"),
        "serie": serie_de(df),
    }
    ruas = []
    for logr, g in df.groupby("logradouro"):
        ruas.append({
            "nome": expande(logr), "mediana": round(float(g.preco_m2.median())),
            "n": int(len(g)), "valor": round(float(g.valor.median())),
            "ultima": g.data.max().strftime("%d/%m/%Y"), "serie": serie_de(g),
        })
    ruas.sort(key=lambda r: -r["n"])

    predios = []
    for (logr, num), g in df.groupby(["logradouro", "numero"]):
        g = g.sort_values("data", ascending=False)
        refs = [x for x in g.referencia if x] if "referencia" in g else []
        predios.append({
            "rua": expande(logr), "numero": num,
            "ref": refs[0] if refs else "",
            "mediana": round(float(g.preco_m2.median())), "n": int(len(g)),
            "area": round(float(g.area.median())), "valor": round(float(g.valor.median())),
            "ultima": g.data.max().strftime("%d/%m/%Y"),
            "transacoes": [{
                "cadastro": r.cadastro, "data": r.data.strftime("%d/%m/%Y"),
                "valor": round(r.valor), "preco_m2": round(r.preco_m2),
            } for r in g.itertuples()],
        })
    predios.sort(key=lambda p: -p["n"])
    return {"bairro": bairro, "ruas": ruas, "predios": predios}


# Nomes que aparecem na coluna de bairro mas não são bairros (condomínios, lixo de digitação)
EXCLUIR = {"THE COLLECTION"}


def bairro_valido(nome):
    if not nome or len(nome) < 4 or nome in EXCLUIR or re.search(r"\d", nome):
        return False
    return not re.match(r"^(TORRE|BLOCO|BL|AP|APTO)\b", nome)


def slug(nome):
    s = unicodedata.normalize("NFD", nome).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "-", s).strip("-")


ABREV_BAIRRO = {"VL": "VILA", "JD": "JARDIM", "STA": "SANTA", "STO": "SANTO",
                "PQ": "PARQUE", "CJ": "CONJUNTO", "SAO": "SÃO", "S": "SÃO"}


def normalizar_bairro(nome):
    return " ".join(ABREV_BAIRRO.get(t, t) for t in nome.split())


def nome_exibido(nome_original):
    return NOMES.get(nome_original, nome_original.title())


def reais(v):
    return f"R$ {v:,.0f}".replace(",", ".")


INDICE = """<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Preço do m² por bairro · São Paulo</title>
<style>
:root { --bg:#f5f3ee; --fg:#1c1e21; --muted:#6a6e75; --line:#dad6cc; --card:#fbfaf7; --accent:#0d5b58; }
@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) { --bg:#15171a; --fg:#ecebe6; --muted:#9a9ea5; --line:#2e3238; --card:#1c1f23; --accent:#6fc3b8; color-scheme: dark; }
}
body { margin:0; background:var(--bg); color:var(--fg); font:15px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif;
  padding-left:max(16px, env(safe-area-inset-left, 0px)); padding-right:max(16px, env(safe-area-inset-right, 0px)); padding-block:24px; }
main { max-width:900px; margin:0 auto; }
h1 { font:600 30px/1.2 Georgia,"Times New Roman",serif; margin:0 0 8px; }
p { color:var(--muted); max-width:70ch; }
input { width:100%; padding:12px 14px; font:16px system-ui,sans-serif; border:1px solid var(--line); border-radius:8px; background:var(--card); color:var(--fg); box-sizing:border-box; margin:16px 0; }
table { width:100%; border-collapse:collapse; background:var(--card); border:1px solid var(--line); border-radius:10px; overflow:hidden; }
th, td { padding:10px 12px; text-align:left; border-bottom:1px solid var(--line); }
th { font-size:12px; text-transform:uppercase; letter-spacing:.06em; color:var(--muted); }
.n { text-align:right; font-variant-numeric:tabular-nums; }
a { color:var(--accent); }
tr[hidden] { display:none; }
</style>
</head>
<body>
<main>
<h1>Preço do m² por bairro</h1>
<p>Valor declarado por m² de área construída, a partir das transações de ITBI publicadas pela Prefeitura de São Paulo. Escolha um bairro para ver ruas, prédios e a evolução mensal.</p>
<input id="busca" type="search" placeholder="Buscar bairro" aria-label="Buscar bairro" autocomplete="off">
<table>
<thead><tr><th>Bairro</th><th class="n">R$/m² mediano</th><th class="n">Transações</th></tr></thead>
<tbody id="lista">__LINHAS__</tbody>
</table>
<p style="font-size:12px; margin-top:20px">Fonte: Secretaria Municipal da Fazenda, guias de ITBI pagas (__PERIODO_GERAL__). Valor declarado não é avaliação oficial.</p>
</main>
<script>
const busca = document.getElementById("busca");
busca.addEventListener("input", () => {
  const q = busca.value.trim().toLowerCase();
  document.querySelectorAll("#lista tr").forEach(r => { r.hidden = !!q && !r.dataset.nome.includes(q); });
});
</script>
</body>
</html>
"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--arquivos", nargs="*", help="planilhas locais; se omitido, baixa dos Releases")
    args = ap.parse_args()

    if args.arquivos:
        caminhos = [Path(a) for a in args.arquivos]
    else:
        caminhos = [baixar(ano, url) for ano, url in RELEASES.items()]

    df = pd.concat([ler(c) for c in caminhos], ignore_index=True)
    df = aplicar_regras(df)
    df = df.assign(bairro=df.bairro.map(normalizar_bairro))
    if df.empty:
        raise SystemExit("Nenhuma transação restou depois das regras. Confira as planilhas.")

    periodo_geral = f"{df.data.min():%d/%m/%Y} a {df.data.max():%d/%m/%Y}"

    if SITE.exists():
        shutil.rmtree(SITE)
    (SITE / "bairros").mkdir(parents=True)
    modelo = MODELO.read_text(encoding="utf-8")

    resumo = []
    print("\nGerando páginas de bairro:")
    for nome, g in df.groupby("bairro"):
        if not bairro_valido(nome) or len(g) < MIN_TRANSACOES:
            continue
        dados = dados_do_bairro(g, nome)
        periodo = f"{dados['bairro']['inicio']} a {dados['bairro']['fim']}"
        exibido = nome_exibido(nome)
        html = (modelo
                .replace("/*__DATA__*/", json.dumps(dados, ensure_ascii=False))
                .replace("__BAIRRO__", escape(exibido))
                .replace("__PERIODO_CURTO__", periodo)
                .replace("__PERIODO__", periodo))
        arq = slug(nome)
        (SITE / "bairros" / f"{arq}.html").write_text(html, encoding="utf-8")
        resumo.append((exibido, arq, dados["bairro"]["mediana"], dados["bairro"]["n"]))
        print(f"  {exibido}: {dados['bairro']['n']} transações")

    resumo.sort(key=lambda r: -r[3])
    linhas = "\n".join(
        f'<tr data-nome="{escape(slug(n).replace("-", " "))} {escape(n.lower())}">'
        f'<td><a href="bairros/{a}.html">{escape(n)}</a></td>'
        f'<td class="n">{reais(m)}</td><td class="n">{t}</td></tr>'
        for n, a, m, t in resumo
    )
    indice = INDICE.replace("__LINHAS__", linhas).replace("__PERIODO_GERAL__", periodo_geral)
    (SITE / "index.html").write_text(indice, encoding="utf-8")
    print(f"\nPronto: {len(resumo)} bairros em {SITE}")


if __name__ == "__main__":
    main()
