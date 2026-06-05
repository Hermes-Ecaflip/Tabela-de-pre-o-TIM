import pandas as pd
import json
import re
import urllib.parse
from datetime import datetime
from collections import defaultdict

# ─────────────────────────────────────────────
# Nomes exatos das colunas de plano na planilha
# ─────────────────────────────────────────────
PLAN_COLS = [
    'TIM CONTROLE ', 'TIM CONTROLE PLUS', 'TIM CONTROLE PREMIUM',
    'TIM CONTROLE SMART', 'TIM CONTROLE REDES SOCIAIS ',
    'TIM BLACK', 'TIM BLACK PLUS', 'TIM BLACK PREMIUM',
    'TIM BLACK A (15GB)', 'TIM BLACK B (20GB)', 'TIM BLACK C (25GB)',
    'TIM BLACK C ULTRA', 'TIM BLACK FAMÍLIA', 'TIM BLACK FAMÍLIA A ONE',
    'TIM BLACK FAMÍLIA PLUS', 'TIM BLACK FAMÍLIA PREMIUM',
    'TIM BLACK FAMÍLIA C ONE', 'TIM BLACK FAMÍLIA VIP',
    'TIM Fibra 300M 24', 'TIM Fibra 400M 24', 'TIM Fibra 500M 24 - 2',
    'TIM Fibra 600M 24 - 2', 'TIM Fibra 600M P 24 - 2',
    'TIM Fibra 600M M 24 - 2', 'TIM Fibra 600M GP 25',
    'TIM Fibra 1GB 24', 'TIM Fibra 1GB P 24',
    'TIM Fibra 1GB M 24', 'TIM Fibra 2GB 24',
    'TIM FIXO LOCAL TOTAL PLUS (Plano Fidelizado) ',
    'TIM FIXO BRASIL TOTAL PLUS (Plano Fidelizado)',
    'TIM FIXO TOTAL LDI PLUS (Plano Fidelizado)',
    'TIM LIVE INTERNET 30GB (sem fidelização)',
    'TIM LIVE INTERNET 50GB (sem fidelização)',
    'TIM LIVE INTERNET 80GB (sem fidelização)',
    'TIM LIVE INTERNET 30GB PLUS ',
    'TIM LIVE INTERNET 50GB PLUS ',
    'TIM LIVE INTERNET 80GB PLUS',
]

# Acessórios fixos da loja (não estão na planilha TIM)
LOJA_ITEMS = [
    ("CAPA TRANSPARENTE",          "ACESSÓRIOS LOJA", 39.99),
    ("CARREGADOR APPLE USB C 20W", "ACESSÓRIOS LOJA", 219.00),
    ("CARREGADOR FAST CHARGING 25W","ACESSÓRIOS LOJA", 199.00),
    ("CABO C 2M IMENSO",           "ACESSÓRIOS LOJA", 49.00),
    ("POWERBANK 10000MAH",         "ACESSÓRIOS LOJA", 99.00),
    ("CAIXINHA MINI LEHMOX",       "ACESSÓRIOS LOJA", 59.00),
]

def clean_val(val):
    if pd.isna(val): return None
    try:
        v = float(str(val).replace(',', '.'))
        return v if v > 0 else None
    except:
        return None

def clean_str(val):
    return str(val).strip() if pd.notna(val) else ''

def process_sheet(df, tipo, has_tecnologia):
    products = []
    for _, row in df.iterrows():
        codigo   = clean_str(row.get('CÓDIGO', ''))
        descricao = clean_str(row.get('DESCRIÇÃO COMERCIAL', ''))
        if not codigo or codigo in ('nan', '0') or not descricao or descricao == 'nan':
            continue
        plans = {}
        for col in PLAN_COLS:
            if col in df.columns:
                v = clean_val(row.get(col))
                if v is not None:
                    plans[col.strip()] = v  # remove trailing spaces no nome do plano
        products.append({
            'codigo':          codigo,
            'descricao':       descricao,
            'descricao_sap':   clean_str(row.get('DESCRIÇÃO SAP', '')),
            'categoria':       clean_str(row.get('CATEGORIA\nMKT', '')).replace('\n', ' '),
            'pre':             clean_val(row.get('PRÉ')),
            'base_retail':     clean_val(row.get('PREÇO BASE FATURAMENTO RETAIL')),
            'controle_nao_fid':clean_val(row.get('CONTROLE NÃO FIDELIZADO - TODOS\n(Fatura / Express)')),
            'pos_nao_fid':     clean_val(row.get('PÓS PAGO NÃO FIDELIZADO - VIDE ABA DE REGRAS')),
            'data_fim':        clean_str(row.get('DATA FIM DA OFERTA', '')),
            'tecnologia':      clean_str(row.get('TECNOLOGIA', '')) if has_tecnologia else '',
            'tipo':            tipo,
            'planos':          plans,
            'variants':        [clean_str(row.get('DESCRIÇÃO SAP', ''))],
        })
    return products

def dedup(products):
    groups = defaultdict(list)
    for p in products:
        groups[p['descricao'].upper().strip()].append(p)
    merged = []
    for items in groups.values():
        base = items[0].copy()
        all_plans = {}
        for item in items:
            for plan, price in item['planos'].items():
                if plan not in all_plans or price < all_plans[plan]:
                    all_plans[plan] = price
        for field in ['base_retail', 'controle_nao_fid', 'pos_nao_fid', 'pre']:
            vals = [i[field] for i in items if i.get(field) is not None]
            base[field] = min(vals) if vals else None
        base['planos']   = all_plans
        base['variants'] = list(set(i['descricao_sap'] for i in items if i.get('descricao_sap')))
        merged.append(base)
    return merged

def get_date():
    df_raw = pd.read_excel('tabela_precos.xlsx', sheet_name='REV', header=None, nrows=13)
    for i in range(13):
        row = df_raw.iloc[i]
        for j, val in enumerate(row):
            if str(val).strip() == 'DATA INÍCIO:' and j + 1 < len(row):
                raw = row.iloc[j + 1]
                if isinstance(raw, datetime):
                    return raw.strftime("%d/%m/%Y")
                parts = str(raw).strip().split(' ')[0].split('-')
                if len(parts) == 3:
                    return f"{parts[2]}/{parts[1]}/{parts[0]}"
    return "Recente"

def main():
    print("📋 Lendo planilha...")
    data_atualizacao = get_date()
    print(f"📅 Data identificada: {data_atualizacao}")

    df_rev = pd.read_excel('tabela_precos.xlsx', sheet_name='REV',          header=11)
    df_acc = pd.read_excel('tabela_precos.xlsx', sheet_name='ACESSÓRIOS REV', header=6)

    phones = dedup(process_sheet(df_rev, 'aparelho',  has_tecnologia=True))
    acc    = dedup(process_sheet(df_acc, 'acessorio', has_tecnologia=False))

    # Adiciona itens fixos da loja que não estão na planilha TIM
    existing = {a['descricao'].upper() for a in acc}
    for desc, cat, price in LOJA_ITEMS:
        if desc.upper() not in existing:
            acc.append({
                'codigo': 'LOJA', 'descricao': desc, 'descricao_sap': desc,
                'categoria': cat, 'pre': price, 'base_retail': None,
                'controle_nao_fid': None, 'pos_nao_fid': None,
                'data_fim': '', 'tecnologia': '', 'tipo': 'acessorio',
                'planos': {}, 'variants': [],
            })

    all_data   = sorted(phones + acc, key=lambda x: x['descricao'])
    total      = len(all_data)
    data_json  = json.dumps(all_data, ensure_ascii=False, separators=(',', ':'))

    print(f"✅ {total} itens extraídos ({len(phones)} aparelhos + {len(acc)} acessórios)")

    # ── 1. script.js ──────────────────────────────────────────────────────
    with open('script.js', 'r', encoding='utf-8') as f:
        js = f.read()
    js = re.sub(r'const DATA = \[.*?\];', f'const DATA = {data_json};', js, flags=re.DOTALL, count=1)
    with open('script.js', 'w', encoding='utf-8') as f:
        f.write(js)
    print("✅ script.js atualizado")

    # ── 2. index.html ─────────────────────────────────────────────────────
    # Nota: os padrões abaixo foram ajustados para a estrutura HTML semântica
    # que usa <time datetime="..."> e <p class="header-badge"> em vez de <div class="hbadge">
    with open('index.html', 'r', encoding='utf-8') as f:
        html = f.read()

    # Datas dentro de elementos <time> (header e footer)
    data_iso = f"{data_atualizacao[6:]}-{data_atualizacao[3:5]}-{data_atualizacao[:2]}"
    novo_time = f'<time datetime="{data_iso}">{data_atualizacao}</time>'
    html = re.sub(r'<time datetime="[\d-]+">[^<]+</time>', novo_time, html)

    # Banner de atualização
    html = re.sub(r'Tabela atualizada em <strong>[\d/]+</strong>',
                  f'Tabela atualizada em <strong>{data_atualizacao}</strong>', html)
    html = re.sub(r'Vigência: [\d/]+', f'Vigência: {data_atualizacao}', html)

    # Badge de contagem no header
    html = re.sub(r'<p class="header-badge"[^>]*>\d+ itens</p>',
                  f'<p class="header-badge" aria-label="Total de itens">{total} itens</p>', html)

    # Contagem no footer
    html = re.sub(r'\d+ itens &middot; Canal Revendas &middot;',
                  f'{total} itens &middot; Canal Revendas &middot;', html)
    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(html)
    print("✅ index.html atualizado")

    # ── 3. README.md ──────────────────────────────────────────────────────
    with open('README.md', 'r', encoding='utf-8') as f:
        readme = f.read()
    data_enc = urllib.parse.quote(data_atualizacao, safe='')
    readme = re.sub(r'badge/tabela-[^-]+-blue', f'badge/tabela-{data_enc}-blue', readme)
    readme = re.sub(r'\*\*\d+ itens catalogados\*\*',
                    f'**{total} itens catalogados**', readme)
    readme = re.sub(r'Atualizado em [\d/]+', f'Atualizado em {data_atualizacao}', readme)
    with open('README.md', 'w', encoding='utf-8') as f:
        f.write(readme)
    print("✅ README.md atualizado")

    print(f"\n🎉 Concluído! Tabela de {data_atualizacao} com {total} itens publicada.")

if __name__ == "__main__":
    main()
