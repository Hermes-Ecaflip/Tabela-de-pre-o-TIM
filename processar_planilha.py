import pandas as pd
import json
import re

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

    print(f"Extraídos {len(dados_finais)} produtos. A atualizar o script.js...")
    
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
            
        print("Sucesso! script.js atualizado.")
    except Exception as e:
        print(f"Erro ao salvar script.js: {e}")

if __name__ == "__main__":
    processar_excel()