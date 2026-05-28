import pandas as pd
import json
import re
from datetime import datetime
import urllib.parse

def clean_val(val):
    if pd.isna(val) or val == '' or str(val).strip() == '-' or str(val).strip() == '':
        return 0.0
    try:
        return float(str(val).replace(',', '.'))
    except:
        return 0.0

def clean_str(val):
    if pd.isna(val):
        return ""
    return str(val).strip()

def processar_excel():
    print("A ler a planilha de preços...")
    try:
        # Pega a data dinamicamente da planilha
        df_header = pd.read_excel('tabela_precos.xlsx', sheet_name='REV', header=None, nrows=15)
        data_atualizacao = "Recente"
        
        for i, row in df_header.iterrows():
            for j, val in enumerate(row):
                if str(val).strip() == 'DATA INÍCIO:':
                    raw_date = row[j+1]
                    if isinstance(raw_date, datetime):
                        data_atualizacao = raw_date.strftime("%d/%m/%Y")
                    else:
                        raw_str = str(raw_date).strip().split(' ')[0]
                        if '-' in raw_str:
                            parts = raw_str.split('-')
                            if len(parts) == 3:
                                data_atualizacao = f"{parts[2]}/{parts[1]}/{parts[0]}"
                            else:
                                data_atualizacao = raw_str
                        else:
                            data_atualizacao = raw_str
                    break

        print(f"Data da tabela identificada: {data_atualizacao}")

        df_rev = pd.read_excel('tabela_precos.xlsx', sheet_name='REV', header=11)
        df_acc = pd.read_excel('tabela_precos.xlsx', sheet_name='ACESSÓRIOS REV', header=6)
    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")
        return

    dados_finais = []
    plan_cols = [
        'TIM CONTROLE ', 'TIM CONTROLE PLUS', 'TIM CONTROLE PREMIUM', 'TIM CONTROLE SMART', 'TIM CONTROLE REDES SOCIAIS ',
        'TIM BLACK', 'TIM BLACK PLUS', 'TIM BLACK PREMIUM', 'TIM BLACK A (15GB)', 'TIM BLACK B (20GB)', 'TIM BLACK C (25GB)',
        'TIM BLACK C ULTRA', 'TIM BLACK FAMÍLIA', 'TIM BLACK FAMÍLIA A ONE', 'TIM BLACK FAMÍLIA PLUS', 'TIM BLACK FAMÍLIA PREMIUM',
        'TIM BLACK FAMÍLIA C ONE', 'TIM BLACK FAMÍLIA VIP',
        'TIM Fibra 300M 24', 'TIM Fibra 400M 24', 'TIM Fibra 500M 24 - 2', 'TIM Fibra 600M 24 - 2', 'TIM Fibra 600M P 24 - 2',
        'TIM Fibra 600M M 24 - 2', 'TIM Fibra 600M GP 25', 'TIM Fibra 1GB 24', 'TIM Fibra 1GB P 24', 'TIM Fibra 1GB M 24', 'TIM Fibra 2GB 24',
        'TIM FIXO LOCAL TOTAL PLUS (Plano Fidelizado) ', 'TIM FIXO BRASIL TOTAL PLUS (Plano Fidelizado)', 'TIM FIXO TOTAL LDI PLUS (Plano Fidelizado)',
        'TIM LIVE INTERNET 30GB (sem fidelização)', 'TIM LIVE INTERNET 50GB (sem fidelização)', 'TIM LIVE INTERNET 80GB (sem fidelização)',
        'TIM LIVE INTERNET 30GB PLUS ', 'TIM LIVE INTERNET 50GB PLUS ', 'TIM LIVE INTERNET 80GB PLUS'
    ]

    for is_rev, df in [(True, df_rev), (False, df_acc)]:
        for idx, row in df.iterrows():
            codigo = clean_str(row.get('CÓDIGO'))
            if not codigo or codigo == '0' or codigo == 'nan':
                continue
                
            item = {
                "codigo": codigo,
                "descricao": clean_str(row.get("DESCRIÇÃO COMERCIAL")),
                "descricao_sap": clean_str(row.get("DESCRIÇÃO SAP")),
                "categoria": clean_str(row.get("CATEGORIA\nMKT")).replace('\n', ' '),
                "pre": clean_val(row.get("PRÉ")),
                "base_retail": clean_val(row.get("PREÇO BASE FATURAMENTO RETAIL")),
                "controle_nao_fid": clean_val(row.get("CONTROLE NÃO FIDELIZADO - TODOS\n(Fatura / Express)")),
                "pos_nao_fid": clean_val(row.get("PÓS PAGO NÃO FIDELIZADO - VIDE ABA DE REGRAS")),
                "data_fim": clean_str(row.get("DATA FIM DA OFERTA")),
                "tecnologia": clean_str(row.get("TECNOLOGIA")) if "TECNOLOGIA" in df.columns else "",
                "tipo": "aparelho" if is_rev else "acessorio",
                "planos": {},
                "variants": [clean_str(row.get("DESCRIÇÃO SAP"))]
            }
            
            for p in plan_cols:
                if p in df.columns:
                    val = clean_val(row.get(p))
                    if val > 0:
                        plan_name_clean = p.strip().replace('\n', ' ')
                        item["planos"][plan_name_clean] = val
                        
            dados_finais.append(item)

    print(f"Extraídos {len(dados_finais)} produtos. A atualizar ficheiros...")
    
    # 1. Atualizar script.js
    try:
        with open('script.js', 'r', encoding='utf-8') as f:
            conteudo = f.read()
            
        novo_conteudo = re.sub(
            r'const DATA = \[.*?\];', 
            f'const DATA = {json.dumps(dados_finais, ensure_ascii=False, indent=2)};', 
            conteudo, 
            flags=re.DOTALL,
            count=1
        )
        with open('script.js', 'w', encoding='utf-8') as f:
            f.write(novo_conteudo)
    except Exception as e:
        print(f"Erro no script.js: {e}")

    # 2. Atualizar index.html (Cabeçalho e Contagem de itens)
    try:
        with open('index.html', 'r', encoding='utf-8') as f:
            html = f.read()
            
        html = re.sub(
            r'<p>Tabela Revendas.*?</p>', 
            f'<p>Tabela Revendas · Com Fidelização · 📅 {data_atualizacao}</p>', 
            html
        )
        html = re.sub(
            r'<div class="hbadge">\d+ itens</div>', 
            f'<div class="hbadge">{len(dados_finais)} itens</div>', 
            html
        )
        
        with open('index.html', 'w', encoding='utf-8') as f:
            f.write(html)
    except Exception as e:
        print(f"Erro no index.html: {e}")

    # 3. Atualizar README.md (Selo azul)
    try:
        with open('README.md', 'r', encoding='utf-8') as f:
            readme = f.read()
            
        data_encoded = urllib.parse.quote(data_atualizacao, safe="")
        readme = re.sub(
            r'badge/tabela-.*?-blue', 
            f'badge/tabela-{data_encoded}-blue', 
            readme
        )
        
        with open('README.md', 'w', encoding='utf-8') as f:
            f.write(readme)
    except Exception as e:
        print(f"Erro no README.md: {e}")

if __name__ == "__main__":
    processar_excel()