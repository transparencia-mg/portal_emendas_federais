from pathlib import Path
import pandas as pd

# ==========================================
# CAMINHOS
# ==========================================

BASE = Path(__file__).resolve().parent.parent

ARQ_DATAS = BASE / "datas.xlsx"
ARQ_EMENDAS = BASE / "upload" / "dados_gerais_emendas.xlsx"

# ==========================================
# COLUNA UTILIZADA PARA O CRUZAMENTO
# ==========================================
COLUNA_CHAVE = "CÓDIGO SIAFI"

print("Lendo arquivos...")

df_datas = pd.read_excel(ARQ_DATAS, dtype=str)
df_emendas = pd.read_excel(ARQ_EMENDAS, dtype=str)

# Remove espaços
df_datas.columns = df_datas.columns.str.strip()
df_emendas.columns = df_emendas.columns.str.strip()

df_datas[COLUNA_CHAVE] = df_datas[COLUNA_CHAVE].astype(str).str.strip()
df_emendas[COLUNA_CHAVE] = df_emendas[COLUNA_CHAVE].astype(str).str.strip()

# Renomeia as colunas de datas
df_datas = df_datas.rename(columns={
    "Data Assinatura ContratoConvênio": "Data_inicio",
    "Data Vencimento ContratoConvênio": "Data_fim"
})

# Mantém apenas as colunas necessárias
df_datas = df_datas[[COLUNA_CHAVE, "Data_inicio", "Data_fim"]]

# Remove duplicados do SIAFI
df_datas = df_datas.drop_duplicates(subset=[COLUNA_CHAVE])

print(f"Registros em datas.xlsx: {len(df_datas):,}")
print(f"Registros em dados_gerais_emendas: {len(df_emendas):,}")

# Remove colunas antigas, caso existam
for coluna in ["Data_inicio", "Data_fim"]:
    if coluna in df_emendas.columns:
        df_emendas = df_emendas.drop(columns=[coluna])

# Junta os dados
df_final = df_emendas.merge(
    df_datas,
    on=COLUNA_CHAVE,
    how="left"
)

# Salva
df_final.to_excel(ARQ_EMENDAS, index=False)

print()
print("Concluído com sucesso!")

encontrados = df_final["Data_inicio"].notna().sum()

print(f"Registros com datas encontradas: {encontrados:,}")
print(f"Registros sem correspondência: {len(df_final)-encontrados:,}")
print(f"Arquivo atualizado: {ARQ_EMENDAS}")