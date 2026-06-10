"""
═══════════════════════════════════════════════════════════════════════
 TIM Consulta de Preços — Processador da Planilha Oficial (Revendas)
═══════════════════════════════════════════════════════════════════════

 Lê a planilha `tabela_precos.xlsx`, extrai aparelhos e acessórios de
 todas as colunas de plano e atualiza automaticamente:

   • script.js   → array DATA com todos os produtos
   • index.html  → data de vigência e contagem de itens
   • README.md   → selo de data e contagem de itens

 IMPORTANTE: as colunas de plano são DETECTADAS automaticamente entre as
 colunas-âncora "PÓS PAGO NÃO FIDELIZADO" e "DATA FIM DA OFERTA". Assim,
 se a TIM adicionar ou remover um plano, o script se adapta sozinho — sem
 necessidade de editar nenhuma lista fixa.

 Uso:  python processar_planilha.py
═══════════════════════════════════════════════════════════════════════
"""

import json
import re
import urllib.parse
from datetime import datetime
from collections import defaultdict

import pandas as pd


# ──────────────────────────────────────────────────────────────────────
# CONFIGURAÇÃO
# ──────────────────────────────────────────────────────────────────────

ARQUIVO_XLSX = "tabela_precos.xlsx"

# Cabeçalhos (linha onde começam os títulos das colunas) de cada aba
HEADER_REV = 11   # aba "REV" (aparelhos)
HEADER_ACC = 6    # aba "ACESSÓRIOS REV" (acessórios)

# Colunas-âncora que delimitam o intervalo de planos.
# Tudo que estiver ENTRE elas é tratado como uma coluna de plano.
ANCORA_INICIO = "PÓS PAGO NÃO FIDELIZADO - VIDE ABA DE REGRAS"
ANCORA_FIM = "DATA FIM DA OFERTA"

# Nomes (normalizados) das colunas de preço sem fidelização
COL_CODIGO      = "CÓDIGO"
COL_DESCRICAO   = "DESCRIÇÃO COMERCIAL"
COL_DESC_SAP    = "DESCRIÇÃO SAP"
COL_CATEGORIA   = "CATEGORIA\nMKT"
COL_PRE         = "PRÉ"
COL_BASE        = "PREÇO BASE FATURAMENTO RETAIL"
COL_CTRL_NF     = "CONTROLE NÃO FIDELIZADO - TODOS\n(Fatura / Express)"
COL_POS_NF      = "PÓS PAGO NÃO FIDELIZADO - VIDE ABA DE REGRAS"
COL_DATA_FIM    = "DATA FIM DA OFERTA"
COL_TECNOLOGIA  = "TECNOLOGIA"

# Acessórios fixos da loja (não constam na planilha oficial da TIM)
LOJA_ITEMS = [
    ("CAPA TRANSPARENTE",            "ACESSÓRIOS LOJA", 39.99),
    ("CARREGADOR APPLE USB C 20W",   "ACESSÓRIOS LOJA", 219.00),
    ("CARREGADOR FAST CHARGING 25W", "ACESSÓRIOS LOJA", 199.00),
    ("CABO C 2M IMENSO",             "ACESSÓRIOS LOJA", 49.00),
    ("POWERBANK 10000MAH",           "ACESSÓRIOS LOJA", 99.00),
    ("CAIXINHA MINI LEHMOX",         "ACESSÓRIOS LOJA", 59.00),
]


# ──────────────────────────────────────────────────────────────────────
# FUNÇÕES AUXILIARES DE LIMPEZA
# ──────────────────────────────────────────────────────────────────────

def clean_val(val):
    """Converte um valor de célula em float positivo, ou None se inválido."""
    if pd.isna(val):
        return None
    try:
        numero = float(str(val).replace(",", "."))
        return numero if numero > 0 else None
    except (ValueError, TypeError):
        return None


def clean_str(val):
    """Converte um valor de célula em string limpa (sem espaços nas pontas)."""
    return str(val).strip() if pd.notna(val) else ""


def detectar_colunas_de_plano(df):
    """
    Detecta dinamicamente as colunas de plano de um DataFrame.

    Retorna a lista de nomes de coluna situados ENTRE a âncora de início
    e a âncora de fim. Dessa forma, novos planos adicionados pela TIM são
    capturados automaticamente.
    """
    colunas = list(df.columns)
    nomes_norm = [str(c).strip() for c in colunas]

    try:
        idx_inicio = nomes_norm.index(ANCORA_INICIO) + 1
    except ValueError:
        print(f"⚠️  Âncora de início não encontrada: {ANCORA_INICIO!r}")
        return []

    # A âncora de fim pode não existir em todas as abas; nesse caso vai até o fim
    try:
        idx_fim = nomes_norm.index(ANCORA_FIM)
    except ValueError:
        idx_fim = len(colunas)

    plano_cols = colunas[idx_inicio:idx_fim]
    # Remove colunas vazias ou "Unnamed" geradas pelo pandas
    return [c for c in plano_cols if not str(c).startswith("Unnamed") and str(c).strip()]


# ──────────────────────────────────────────────────────────────────────
# EXTRAÇÃO DOS PRODUTOS
# ──────────────────────────────────────────────────────────────────────

def process_sheet(df, tipo, tem_tecnologia):
    """
    Extrai todos os produtos de uma aba.

    Args:
        df: DataFrame da aba.
        tipo: 'aparelho' ou 'acessorio'.
        tem_tecnologia: se a aba possui coluna TECNOLOGIA.

    Returns:
        Lista de dicionários de produto.
    """
    plano_cols = detectar_colunas_de_plano(df)
    print(f"   • {len(plano_cols)} colunas de plano detectadas em '{tipo}'")

    produtos = []
    for _, row in df.iterrows():
        codigo = clean_str(row.get(COL_CODIGO, ""))
        descricao = clean_str(row.get(COL_DESCRICAO, ""))

        # Ignora linhas sem código ou descrição válidos
        if not codigo or codigo in ("nan", "0") or not descricao or descricao == "nan":
            continue

        # Monta o dicionário de planos (apenas com preço válido)
        planos = {}
        for col in plano_cols:
            preco = clean_val(row.get(col))
            if preco is not None:
                planos[str(col).strip()] = preco

        produtos.append({
            "codigo":           codigo,
            "descricao":        descricao,
            "descricao_sap":    clean_str(row.get(COL_DESC_SAP, "")),
            "categoria":        clean_str(row.get(COL_CATEGORIA, "")).replace("\n", " "),
            "pre":              clean_val(row.get(COL_PRE)),
            "base_retail":      clean_val(row.get(COL_BASE)),
            "controle_nao_fid": clean_val(row.get(COL_CTRL_NF)),
            "pos_nao_fid":      clean_val(row.get(COL_POS_NF)),
            "data_fim":         clean_str(row.get(COL_DATA_FIM, "")),
            "tecnologia":       clean_str(row.get(COL_TECNOLOGIA, "")) if tem_tecnologia else "",
            "tipo":             tipo,
            "planos":           planos,
            "variants":         [clean_str(row.get(COL_DESC_SAP, ""))],
        })

    return produtos


def dedup(produtos):
    """
    Agrupa produtos com a mesma descrição (variantes de cor) em um só item,
    mantendo o menor preço de cada plano e de cada campo sem fidelização.
    """
    grupos = defaultdict(list)
    for p in produtos:
        grupos[p["descricao"].upper().strip()].append(p)

    resultado = []
    for itens in grupos.values():
        base = itens[0].copy()

        # Menor preço de cada plano entre as variantes
        planos_merged = {}
        for item in itens:
            for plano, preco in item["planos"].items():
                if plano not in planos_merged or preco < planos_merged[plano]:
                    planos_merged[plano] = preco

        # Menor preço dos campos sem fidelização
        for campo in ["base_retail", "controle_nao_fid", "pos_nao_fid", "pre"]:
            valores = [i[campo] for i in itens if i.get(campo) is not None]
            base[campo] = min(valores) if valores else None

        base["planos"] = planos_merged
        base["variants"] = list({i["descricao_sap"] for i in itens if i.get("descricao_sap")})
        resultado.append(base)

    return resultado


def adicionar_itens_loja(acessorios):
    """Acrescenta os acessórios fixos da loja que não constam na planilha."""
    existentes = {a["descricao"].upper() for a in acessorios}
    for desc, cat, preco in LOJA_ITEMS:
        if desc.upper() not in existentes:
            acessorios.append({
                "codigo": "LOJA", "descricao": desc, "descricao_sap": desc,
                "categoria": cat, "pre": preco, "base_retail": None,
                "controle_nao_fid": None, "pos_nao_fid": None,
                "data_fim": "", "tecnologia": "", "tipo": "acessorio",
                "planos": {}, "variants": [],
            })
    return acessorios


# ──────────────────────────────────────────────────────────────────────
# LEITURA DA DATA DA TABELA
# ──────────────────────────────────────────────────────────────────────

def ler_data_tabela():
    """Lê a data de vigência do campo 'DATA INÍCIO:' nas primeiras linhas da aba REV."""
    df_raw = pd.read_excel(ARQUIVO_XLSX, sheet_name="REV", header=None, nrows=13)
    for i in range(len(df_raw)):
        row = df_raw.iloc[i]
        for j, val in enumerate(row):
            if str(val).strip() == "DATA INÍCIO:" and j + 1 < len(row):
                bruto = row.iloc[j + 1]
                if isinstance(bruto, datetime):
                    return bruto.strftime("%d/%m/%Y")
                partes = str(bruto).strip().split(" ")[0].split("-")
                if len(partes) == 3:
                    return f"{partes[2]}/{partes[1]}/{partes[0]}"
    return datetime.now().strftime("%d/%m/%Y")


# ──────────────────────────────────────────────────────────────────────
# ATUALIZAÇÃO DOS ARQUIVOS DO SITE
# ──────────────────────────────────────────────────────────────────────

def atualizar_script_js(data_json):
    """Substitui o array DATA dentro de script.js."""
    with open("script.js", "r", encoding="utf-8") as f:
        conteudo = f.read()
    conteudo = re.sub(
        r"const DATA = \[.*?\];",
        f"const DATA = {data_json};",
        conteudo,
        flags=re.DOTALL,
        count=1,
    )
    with open("script.js", "w", encoding="utf-8") as f:
        f.write(conteudo)
    print("✅ script.js atualizado")


def atualizar_index_html(data_br, total):
    """Atualiza data de vigência e contagem de itens em index.html."""
    with open("index.html", "r", encoding="utf-8") as f:
        html = f.read()

    # Data em formato ISO para o atributo datetime das tags <time>
    data_iso = f"{data_br[6:]}-{data_br[3:5]}-{data_br[:2]}"
    novo_time = f'<time datetime="{data_iso}">{data_br}</time>'

    # Todas as tags <time> (header e footer)
    html = re.sub(r'<time datetime="[\d-]+">[^<]+</time>', novo_time, html)

    # Banner de atualização
    html = re.sub(
        r"Tabela atualizada em <strong>[\d/]+</strong>",
        f"Tabela atualizada em <strong>{data_br}</strong>",
        html,
    )
    html = re.sub(r"Vigência: [\d/]+", f"Vigência: {data_br}", html)

    # Badge de contagem no header
    html = re.sub(
        r'<p class="header-badge"[^>]*>\d+ itens</p>',
        f'<p class="header-badge" aria-label="Total de itens">{total} itens</p>',
        html,
    )

    # Contagem no rodapé
    html = re.sub(
        r"\d+ itens &middot; Canal Revendas &middot;",
        f"{total} itens &middot; Canal Revendas &middot;",
        html,
    )

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("✅ index.html atualizado")


def atualizar_readme(data_br, total):
    """Atualiza o selo de data e a contagem de itens em README.md."""
    with open("README.md", "r", encoding="utf-8") as f:
        readme = f.read()

    data_enc = urllib.parse.quote(data_br, safe="")
    readme = re.sub(r"badge/tabela-[^-]+-blue", f"badge/tabela-{data_enc}-blue", readme)
    readme = re.sub(r"\*\*\d+ itens catalogados\*\*", f"**{total} itens catalogados**", readme)
    readme = re.sub(r"Atualizado em [\d/]+", f"Atualizado em {data_br}", readme)

    with open("README.md", "w", encoding="utf-8") as f:
        f.write(readme)
    print("✅ README.md atualizado")


# ──────────────────────────────────────────────────────────────────────
# FLUXO PRINCIPAL
# ──────────────────────────────────────────────────────────────────────

def main():
    print("📋 Lendo planilha...")
    data_br = ler_data_tabela()
    print(f"📅 Data identificada: {data_br}")

    df_rev = pd.read_excel(ARQUIVO_XLSX, sheet_name="REV", header=HEADER_REV)
    df_acc = pd.read_excel(ARQUIVO_XLSX, sheet_name="ACESSÓRIOS REV", header=HEADER_ACC)

    aparelhos  = dedup(process_sheet(df_rev, "aparelho", tem_tecnologia=True))
    acessorios = dedup(process_sheet(df_acc, "acessorio", tem_tecnologia=False))
    acessorios = adicionar_itens_loja(acessorios)

    todos = sorted(aparelhos + acessorios, key=lambda x: x["descricao"])
    total = len(todos)
    data_json = json.dumps(todos, ensure_ascii=False, separators=(",", ":"))

    print(f"✅ {total} itens extraídos ({len(aparelhos)} aparelhos + {len(acessorios)} acessórios)")

    atualizar_script_js(data_json)
    atualizar_index_html(data_br, total)
    atualizar_readme(data_br, total)

    print(f"\n🎉 Concluído! Tabela de {data_br} com {total} itens publicada.")


if __name__ == "__main__":
    main()
