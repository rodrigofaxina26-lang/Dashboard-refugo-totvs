# app.py - VERSÃO CORRIGIDA (MERGE FUNCIONANDO)
from flask import Flask, request, render_template
import pandas as pd
import requests
import json
import os
from datetime import datetime, date, timedelta
import traceback

SUPABASE_URL = "https://qrauhqqbwnaafljvvjxm.supabase.co"
SUPABASE_KEY = "sb_publishable_Ik7kIP8JSUQi_R4iro7AGA_3qPEZto6"

HEADERS = {
    "apikey": SUPABASE_KEY,
    "Authorization": f"Bearer {SUPABASE_KEY}",
    "Content-Type": "application/json"
}

app = Flask(__name__)

class DataSource:
    def __init__(self):
        self.config = {
            'dados_totvs_path': r"P:\QUALIDADE\USUARIOS\00. Dept. Qualidade\07. Controle de Refugo\2026\Planilha de Refugo _ Dados TOTVs.xlsx",
            'dados_sheet': "Dados"
        }
        
        self._cache = {
            "df_sci": pd.DataFrame(),
            "df_spr": pd.DataFrame(),
            "last_modified": 0.0,
            "tabela_custos": pd.DataFrame(),
            "tabela_defeitos": pd.DataFrame()
        }
    
    def carregar_tabelas_auxiliares(self, force_reload=False):
        if not self._cache["tabela_custos"].empty and not force_reload:
            return self._cache["tabela_custos"], self._cache["tabela_defeitos"]
        
        try:
            # Lê a aba "Dados" uma única vez — contém custos e descrição de defeitos
            df_dados = pd.read_excel(
                self.config['dados_totvs_path'],
                sheet_name=self.config['dados_sheet'],
                header=None
            )

            # --- Tabela de custos ---
            # Coluna B (índice 1) = Produto | Coluna R (índice 17) = Valor (custo unitário)
            try:
                df_custos = pd.DataFrame()
                df_custos['PRODUTO'] = df_dados.iloc[1:, 1].values   # Coluna B
                df_custos['VALOR']   = df_dados.iloc[1:, 17].values  # Coluna R
                df_custos['VALOR']   = pd.to_numeric(df_custos['VALOR'], errors='coerce')
                df_custos = df_custos.dropna(subset=['VALOR'])
                df_custos = df_custos[df_custos['PRODUTO'].astype(str).str.strip() != '']
                self.custos_dict = dict(zip(df_custos['PRODUTO'].astype(str).str.strip(), df_custos['VALOR']))
                print(f"✅ Tabela de custos: {len(df_custos)} produtos")
            except Exception as e_custos:
                print(f"⚠️ Erro ao ler tabela de custos: {e_custos}")
                self.custos_dict = {}
                df_custos = pd.DataFrame()

            # --- Tabela de defeitos ---
            # Lê do arquivo local defeitos.csv (mais confiável que a planilha em modo leitura)
            try:
                csv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "defeitos.csv")
                df_defeitos = pd.read_csv(csv_path, dtype=str)
                self.defeitos_dict = {}
                for _, row in df_defeitos.iterrows():
                    codigo    = str(row['codigo']).strip()
                    descricao = str(row['descricao']).strip()
                    self.defeitos_dict[codigo] = f"{codigo} - {descricao}"
                    try:
                        codigo_num = str(int(float(codigo)))
                        self.defeitos_dict[codigo_num] = f"{codigo} - {descricao}"
                    except:
                        pass
                print(f"✅ Tabela de defeitos: {len(df_defeitos)} motivos (defeitos.csv)")
            except Exception as e_def:
                print(f"⚠️ Erro ao ler defeitos.csv: {e_def}")
                self.defeitos_dict = {}
                df_defeitos = pd.DataFrame()
                self.defeitos_dict = {}
                df_defeitos = pd.DataFrame()

            self._cache["tabela_custos"]   = df_custos
            self._cache["tabela_defeitos"] = df_defeitos

            return df_custos, df_defeitos

        except Exception as e:
            print(f"Erro: {e}")
            self.custos_dict = {}
            self.defeitos_dict = {}
            return pd.DataFrame(), pd.DataFrame()
    
    def carregar_dados(self, force_reload=False):
        df_custos, df_defeitos = self.carregar_tabelas_auxiliares(force_reload)

        try:
            print("📂 Carregando dados do Supabase...")

            # Busca todos os registros em páginas de 1000
            todos = []
            offset = 0
            page = 1000
            while True:
                resp = requests.get(
                    f"{SUPABASE_URL}/rest/v1/refugo?select=*",
                    headers={**HEADERS, "Range": f"{offset}-{offset + page - 1}"},
                )
                parte = resp.json()
                if not parte:
                    break
                todos.extend(parte)
                if len(parte) < page:
                    break
                offset += page

            df_refugo = pd.DataFrame(todos)

            if df_refugo.empty:
                print("❌ Nenhum dado encontrado. Execute importar.py primeiro.")
                return pd.DataFrame(), pd.DataFrame()

            # Padroniza nomes de colunas para maiúsculo
            df_refugo.columns = [c.upper() for c in df_refugo.columns]

            print(f"📊 Registros carregados: {len(df_refugo)}")

            # Converte data
            df_refugo['DATA'] = pd.to_datetime(df_refugo['DATA'], errors='coerce')
            df_refugo = df_refugo.dropna(subset=['DATA'])

            # Converte quantidade
            df_refugo['QTDE'] = pd.to_numeric(df_refugo['QTDE'], errors='coerce').fillna(0)
            
            # Adiciona custo unitário usando dicionário
            df_refugo['CUSTO_UNITARIO'] = df_refugo['PRODUTO'].astype(str).map(self.custos_dict).fillna(0)
            
            # Adiciona descrição do defeito usando dicionário
            df_refugo['MOTIVO_STR'] = df_refugo['MOTIVO_COD'].astype(str).str.strip()
            df_refugo['MODO_DE_FALHA'] = df_refugo['MOTIVO_STR'].map(self.defeitos_dict).fillna(df_refugo['MOTIVO_STR'])
            
            # Calcula valor total
            df_refugo['VALOR_CUSTO'] = df_refugo['QTDE'] * df_refugo['CUSTO_UNITARIO']
            
            df_refugo["DATA_STR"] = df_refugo["DATA"].dt.strftime("%Y-%m-%d")
            
            # Filtra ano
            df_refugo = df_refugo[df_refugo["DATA"].dt.year >= 2025]
            
            # Debug: mostrar exemplos
            print(f"\n📋 Exemplos de dados processados:")
            print(df_refugo[['PRODUTO', 'QTDE', 'CUSTO_UNITARIO', 'VALOR_CUSTO', 'MOTIVO_COD', 'MODO_DE_FALHA']].head(10))
            
            # Separa SCI e SPR
            df_sci = df_refugo[df_refugo['TIPO'].astype(str).str.upper() == "SCI"].copy()
            df_spr = df_refugo[df_refugo['TIPO'].astype(str).str.upper() == "SPR"].copy()
            
            df_sci = df_sci.drop(columns=['MOTIVO_COD', 'MOTIVO_STR'], errors='ignore')
            df_spr = df_spr.drop(columns=['MOTIVO_COD', 'MOTIVO_STR'], errors='ignore')
            
            print(f"\n✅ RESULTADO:")
            print(f"   SCI: {len(df_sci)} registros")
            print(f"   SPR: {len(df_spr)} registros")
            print(f"   Total Custo SCI: R$ {df_sci['VALOR_CUSTO'].sum():,.2f}")
            print(f"   Total Custo SPR: R$ {df_spr['VALOR_CUSTO'].sum():,.2f}")
            
            if not df_sci.empty:
                print(f"\n📋 Amostra SCI:")
                print(df_sci[['PRODUTO', 'QTDE', 'CUSTO_UNITARIO', 'VALOR_CUSTO', 'MODO_DE_FALHA']].head(5))
            
            self._cache["df_sci"] = df_sci
            self._cache["df_spr"] = df_spr
            
            return df_sci, df_spr
            
        except Exception as e:
            print(f"💥 Erro: {e}")
            traceback.print_exc()
            return pd.DataFrame(), pd.DataFrame()

data_source = DataSource()

def get_modo_falha(row):
    if "MODO_DE_FALHA" in row and pd.notna(row["MODO_DE_FALHA"]):
        return str(row["MODO_DE_FALHA"]).strip()
    return "NÃO IDENTIFICADO"

def get_top3_diario(df_filtrado):
    if df_filtrado.empty or "QTDE" not in df_filtrado.columns:
        return {"top_qtde": [], "top_custo": []}
    if "PRODUTO" not in df_filtrado.columns:
        df_filtrado = df_filtrado.copy()
        df_filtrado["PRODUTO"] = "SEM PRODUTO"
    cols_base = ["PRODUTO", "QTDE", "VALOR_CUSTO"]
    top_qtde = df_filtrado.nlargest(3, "QTDE")[cols_base].copy()
    if "VALOR_CUSTO" in df_filtrado.columns:
        top_custo = df_filtrado.nlargest(3, "VALOR_CUSTO")[cols_base].copy()
    else:
        top_custo = pd.DataFrame()
    top_qtde_serial = []
    for idx in top_qtde.index:
        row_completa = df_filtrado.loc[idx]
        modo = get_modo_falha(row_completa)
        top_qtde_serial.append({
            "produto": str(row_completa.get("PRODUTO", "SEM PRODUTO")),
            "qtde": float(row_completa.get("QTDE", 0)),
            "custo": float(row_completa.get("VALOR_CUSTO", 0)),
            "modo_falha": modo
        })
    top_custo_serial = []
    for idx in top_custo.index:
        row_completa = df_filtrado.loc[idx]
        modo = get_modo_falha(row_completa)
        top_custo_serial.append({
            "produto": str(row_completa.get("PRODUTO", "SEM PRODUTO")),
            "qtde": float(row_completa.get("QTDE", 0)),
            "custo": float(row_completa.get("VALOR_CUSTO", 0)),
            "modo_falha": modo
        })
    return {"top_qtde": top_qtde_serial, "top_custo": top_custo_serial}

def filtrar_data_diaria(df, data_especifica):
    if df.empty or "DATA" not in df.columns:
        return pd.DataFrame()
    data_target = pd.to_datetime(data_especifica).date()
    return df[df["DATA"].dt.date == data_target].copy()

def criar_pareto_modo_falha(df, tipo="GERAL"):
    if df.empty:
        return {"modo": [], "qtde": [], "cumperc": [], "total": 0}
    df_temp = df.copy()
    df_temp["MODO_FALHA"] = df_temp.apply(get_modo_falha, axis=1)
    pareto = df_temp.groupby("MODO_FALHA")["QTDE"].sum().reset_index()
    pareto = pareto[pareto["QTDE"] > 0].sort_values("QTDE", ascending=False).reset_index(drop=True)
    if len(pareto) > 0 and pareto["QTDE"].sum() > 0:
        pareto["CUMPERC"] = (pareto["QTDE"].cumsum() / pareto["QTDE"].sum()).round(1)
    else:
        pareto["CUMPERC"] = 0
    return {
        "modo": pareto["MODO_FALHA"].tolist()[:10],
        "qtde": pareto["QTDE"].tolist()[:10],
        "cumperc": pareto["CUMPERC"].tolist()[:10],
        "total": int(pareto["QTDE"].sum())
    }

def get_top3_problemas_por_produto(df, produto_nome):
    if df.empty or not produto_nome:
        return []
    termo = str(produto_nome).strip().upper()
    df_temp = df.copy()
    df_temp["PRODUTO_CHECK"] = df_temp["PRODUTO"].astype(str).str.strip().str.upper()
    df_p = df_temp[df_temp["PRODUTO_CHECK"] == termo].copy()
    if df_p.empty:
        df_p = df_temp[df_temp["PRODUTO_CHECK"].str.contains(termo, na=False)].copy()
    if df_p.empty:
        return []
    df_p["QTDE"] = pd.to_numeric(df_p["QTDE"], errors='coerce').fillna(0)
    df_p = df_p[df_p["QTDE"] > 0]
    df_p["MODO_FALHA_FINAL"] = df_p.apply(get_modo_falha, axis=1)
    res = df_p.groupby("MODO_FALHA_FINAL")["QTDE"].sum().nlargest(3).reset_index()
    res.columns = ["MODO_FALHA", "QTDE"]
    return res.to_dict("records")

@app.route("/produto_pareto")
def produto_pareto():
    p = request.args.get("produto", "").strip().upper()
    i, f = request.args.get("inicio"), request.args.get("fim")
    d1, d2 = data_source.carregar_dados()
    def filt(df):
        if df.empty: return df
        t = df.copy()
        if i: t = t[t["DATA"] >= pd.to_datetime(i)]
        if f: t = t[t["DATA"] <= pd.to_datetime(f)]
        return t
    return json.dumps({"sci": get_top3_problemas_por_produto(filt(d1), p), "spr": get_top3_problemas_por_produto(filt(d2), p)})

@app.route("/pareto")
def pareto():
    data_inicio = request.args.get("inicio", (date.today() - timedelta(days=90)).strftime("%Y-%m-%d"))
    data_fim = request.args.get("fim", date.today().strftime("%Y-%m-%d"))
    produto = request.args.get("produto", "").strip().upper()
    force = pd.to_datetime(data_fim).date() >= date.today()
    df_sci_raw, df_spr_raw = data_source.carregar_dados(force_reload=force)
    
    if df_sci_raw.empty and df_spr_raw.empty:
        return json.dumps({"sci": {"modo": [], "qtde": [], "cumperc": [], "total": 0}, "spr": {"modo": [], "qtde": [], "cumperc": [], "total": 0}, "filtros": "SEM DADOS"})
    
    df_sci = df_sci_raw[(df_sci_raw["DATA"] >= pd.to_datetime(data_inicio)) & (df_sci_raw["DATA"] <= pd.to_datetime(data_fim))].copy() if not df_sci_raw.empty else pd.DataFrame()
    df_spr = df_spr_raw[(df_spr_raw["DATA"] >= pd.to_datetime(data_inicio)) & (df_spr_raw["DATA"] <= pd.to_datetime(data_fim))].copy() if not df_spr_raw.empty else pd.DataFrame()
    
    if produto:
        if not df_sci.empty and "PRODUTO" in df_sci.columns:
            df_sci = df_sci[df_sci["PRODUTO"].astype(str).str.upper().str.contains(produto, na=False)]
        if not df_spr.empty and "PRODUTO" in df_spr.columns:
            df_spr = df_spr[df_spr["PRODUTO"].astype(str).str.upper().str.contains(produto, na=False)]
    
    pareto_sci = criar_pareto_modo_falha(df_sci)
    pareto_spr = criar_pareto_modo_falha(df_spr)
    return json.dumps({"sci": pareto_sci, "spr": pareto_spr, "filtros": f"{data_inicio} a {data_fim} | Produto: {produto or 'TODOS'}"})

@app.route("/detalhe")
def detalhe():
    data_especifica = request.args.get("data", date.today().strftime("%Y-%m-%d"))
    force = pd.to_datetime(data_especifica).date() >= date.today()
    df_sci_raw, df_spr_raw = data_source.carregar_dados(force_reload=force)
    df_completo = pd.concat([df_sci_raw, df_spr_raw], ignore_index=True)
    df_filtrado = filtrar_data_diaria(df_completo, data_especifica)
    top3 = get_top3_diario(df_filtrado)
    return json.dumps({"data": data_especifica, "total_registros": len(df_filtrado), "top3": top3})

@app.route("/filtrar")
def filtrar():
    data_inicio = request.args.get("inicio", (date.today() - timedelta(days=90)).strftime("%Y-%m-%d"))
    data_fim = request.args.get("fim", date.today().strftime("%Y-%m-%d"))
    produto = request.args.get("produto", "").strip().upper()
    force = pd.to_datetime(data_fim).date() >= date.today()
    df_sci_raw, df_spr_raw = data_source.carregar_dados(force_reload=force)
    
    if df_sci_raw.empty and df_spr_raw.empty:
        return json.dumps({"run": [], "sci_q": [], "sci_c": [], "spr_q": [], "spr_c": [], "periodo": "SEM DADOS", "total_dias": 0})
    
    df_sci = df_sci_raw[(df_sci_raw["DATA"] >= pd.to_datetime(data_inicio)) & (df_sci_raw["DATA"] <= pd.to_datetime(data_fim))].copy() if not df_sci_raw.empty else pd.DataFrame()
    df_spr = df_spr_raw[(df_spr_raw["DATA"] >= pd.to_datetime(data_inicio)) & (df_spr_raw["DATA"] <= pd.to_datetime(data_fim))].copy() if not df_spr_raw.empty else pd.DataFrame()
    
    if produto:
        if not df_sci.empty and "PRODUTO" in df_sci.columns:
            df_sci = df_sci[df_sci["PRODUTO"].astype(str).str.upper().str.contains(produto, na=False)]
        if not df_spr.empty and "PRODUTO" in df_spr.columns:
            df_spr = df_spr[df_spr["PRODUTO"].astype(str).str.upper().str.contains(produto, na=False)]
    
    if df_sci.empty:
        res_sci = pd.DataFrame(columns=["DATA_STR", "QTDE", "VALOR_CUSTO"])
    else:
        res_sci = df_sci.groupby("DATA_STR")[["QTDE", "VALOR_CUSTO"]].sum().reset_index()
    
    if df_spr.empty:
        res_spr = pd.DataFrame(columns=["DATA_STR", "QTDE", "VALOR_CUSTO"])
    else:
        res_spr = df_spr.groupby("DATA_STR")[["QTDE", "VALOR_CUSTO"]].sum().reset_index()
    
    run_df = pd.merge(res_sci, res_spr, on="DATA_STR", how="outer", suffixes=("_SCI", "_SPR")).fillna(0)
    run_df["TOTAL_Q"] = run_df["QTDE_SCI"] + run_df["QTDE_SPR"]
    run_df["TOTAL_C"] = run_df["VALOR_CUSTO_SCI"] + run_df["VALOR_CUSTO_SPR"]
    run_df = run_df.sort_values("DATA_STR")
    
    def top10(df, col):
        if df.empty or "PRODUTO" not in df.columns:
            return [{"PRODUTO": "SEM DADOS", col: 0}]
        try:
            resultado = df.groupby("PRODUTO")[col].sum().nlargest(10).reset_index()
            return resultado.to_dict("records")
        except:
            return [{"PRODUTO": "ERRO", col: 0}]
    
    dados_json = {
        "run": run_df.to_dict("records"),
        "sci_q": top10(df_sci, "QTDE"),
        "sci_c": top10(df_sci, "VALOR_CUSTO"),
        "spr_q": top10(df_spr, "QTDE"),
        "spr_c": top10(df_spr, "VALOR_CUSTO"),
        "periodo": f"{data_inicio} a {data_fim}",
        "total_dias": len(run_df)
    }
    return json.dumps(dados_json)

@app.route("/")
def index():
    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=False, host="0.0.0.0", port=5003)