# importar.py — Sincroniza a planilha TOTVS com o Supabase
# Execute este script sempre que colar novos dados na planilha do TOTVS

import sys
import pandas as pd
import requests
import json
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ─── CONFIGURAÇÃO ────────────────────────────────────────────────────────────
EXCEL_PATH  = r"P:\QUALIDADE\USUARIOS\00. Dept. Qualidade\07. Controle de Refugo\2026\Planilha de refugo diário Totvs.xlsx"
SHEET_NAME  = "Listagem do Browse"

SUPABASE_URL = "https://qrauhqqbwnaafljvvjxm.supabase.co"
SUPABASE_KEY = "sb_publishable_Ik7kIP8JSUQi_R4iro7AGA_3qPEZto6"

HEADERS = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json",
    "Prefer": "return=minimal"
}
# ─────────────────────────────────────────────────────────────────────────────

def importar():
    print("=" * 55)
    print("  IMPORTAÇÃO REFUGO — TOTVS → Supabase")
    print(f"  {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print("=" * 55)

    # 1. Lê a planilha Excel
    print(f"\n📂 Lendo: {EXCEL_PATH}")
    try:
        df_raw = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, header=None)
    except FileNotFoundError:
        print(f"\n❌ Arquivo não encontrado: {EXCEL_PATH}")
        return
    except Exception as e:
        print(f"\n❌ Erro ao ler Excel: {e}")
        return

    # 2. Extrai colunas pela posição (estrutura fixa do TOTVS)
    dados = df_raw.iloc[2:].copy()

    df = pd.DataFrame()
    df['ord_producao'] = dados.iloc[:, 0].values   # Coluna A - Ordem de Produção (chave única)
    df['produto']      = dados.iloc[:, 1].values   # Coluna B - Produto
    df['motivo_cod']   = dados.iloc[:, 3].values   # Coluna D - Motivo Perda (código defeito)
    df['qtde']         = dados.iloc[:, 4].values   # Coluna E - Qtd Perda
    df['data']         = dados.iloc[:, 5].values   # Coluna F - Dt. da Perda
    df['tipo']         = dados.iloc[:, 12].values  # Coluna M - Tipo (SCI/SPR)

    # 3. Limpeza
    df = df.dropna(subset=['produto'])
    df = df[df['produto'].astype(str).str.strip() != '']
    df['data'] = pd.to_datetime(df['data'], dayfirst=True, errors='coerce')
    df = df.dropna(subset=['data'])
    df['qtde'] = pd.to_numeric(df['qtde'], errors='coerce').fillna(0)
    df['tipo'] = df['tipo'].astype(str).str.upper().str.strip()

    # 5. Formata para envio — converte tudo para tipos nativos Python
    df['data']       = df['data'].dt.strftime('%Y-%m-%d')
    df['produto']    = df['produto'].astype(str).str.strip()
    df['motivo_cod'] = df['motivo_cod'].astype(str).str.strip()
    df['qtde']       = df['qtde'].astype(float)
    df['tipo']       = df['tipo'].astype(str).str.strip()

    df['ord_producao'] = df['ord_producao'].astype(str).str.strip()

    print(f"📊 Registros lidos:      {len(df)}")

    # A planilha TOTVS pode ter linhas duplicadas pela mesma OP/motivo/data.
    # Somamos a qtde das duplicatas para manter integridade com a constraint única.
    dupl = df.duplicated(['ord_producao', 'motivo_cod', 'data'], keep=False).sum()
    if dupl:
        df = (
            df.groupby(['ord_producao', 'produto', 'motivo_cod', 'data', 'tipo'], as_index=False)
            .agg({'qtde': 'sum'})
        )
        print(f"⚠️  {dupl} linhas duplicadas agrupadas → {len(df)} registros únicos")

    print(f"✅ Registros a gravar:   {len(df)}")

    # Converte para lista de dicts com tipos nativos Python (evita erro de serialização)
    registros = [
        {
            'ord_producao': str(r['ord_producao']),
            'produto':      str(r['produto']),
            'motivo_cod':   str(r['motivo_cod']),
            'qtde':         float(r['qtde']),
            'data':         str(r['data']),
            'tipo':         str(r['tipo'])
        }
        for r in df.to_dict(orient='records')
    ]

    # 6. Apaga todos os registros existentes
    # Usa data=not.is.null para garantir que todos os registros sejam deletados,
    # independente do tipo do campo 'id' (inteiro ou UUID).
    print(f"\n🗄️  Limpando tabela no Supabase...")
    del_resp = requests.delete(
        f"{SUPABASE_URL}/rest/v1/refugo?data=not.is.null",
        headers=HEADERS
    )
    if del_resp.status_code not in (200, 204):
        print(f"❌ Erro ao limpar tabela: {del_resp.status_code} — {del_resp.text}")
        return
    print("✅ Tabela limpa!")

    # 7. Insere em lotes de 500
    # Prefer: resolution=ignore-duplicates → ON CONFLICT DO NOTHING
    # Protege contra duplicatas dentro da própria planilha.
    print(f"\n📤 Enviando dados...")
    lote_size    = 500
    total_lotes  = (len(registros) + lote_size - 1) // lote_size
    headers_post = {**HEADERS, "Prefer": "return=minimal,resolution=ignore-duplicates"}

    for i in range(0, len(registros), lote_size):
        lote     = registros[i:i + lote_size]
        lote_num = (i // lote_size) + 1
        resp = requests.post(
            f"{SUPABASE_URL}/rest/v1/refugo?on_conflict=ord_producao,motivo_cod,data",
            headers=headers_post,
            data=json.dumps(lote)
        )
        if resp.status_code in (200, 201):
            print(f"   Lote {lote_num}/{total_lotes} enviado ({len(lote)} registros) ✅")
        else:
            print(f"   Lote {lote_num}/{total_lotes} — ERRO {resp.status_code}: {resp.text}")
            return

    print(f"\n🎉 Importação concluída! {len(registros)} registros no Supabase.")
    print("\n" + "=" * 55)

if __name__ == "__main__":
    importar()
    input("\nPressione ENTER para fechar...")
