from pathlib import Path
import subprocess
import pandas as pd


# =====================================================
# CONFIGURAÇÃO
# =====================================================

ATUALIZAR_PLANILHA = True
ATUALIZAR_GITHUB = True

# pasta onde está este script
BASE = Path(__file__).resolve().parent

ARQ_LINKS = BASE / "link_convenios.xlsx"
ARQ_DADOS = BASE / "upload" / "dados_gerais_emendas.xlsx"
ARQ_DEPARA = BASE / "de_para_inteiro_teor.xlsx"


# =====================================================
# LEITURA DOS ARQUIVOS
# =====================================================

print("Lendo planilhas...")

dados = pd.read_excel(ARQ_DADOS)
links = pd.read_excel(ARQ_LINKS)

dados.columns = dados.columns.str.strip()
links.columns = links.columns.str.strip()

dados["CÓDIGO SIAFI"] = (
    dados["CÓDIGO SIAFI"]
    .astype(str)
    .str.strip()
)

links["Código SIAFI"] = (
    links["Código SIAFI"]
    .astype(str)
    .str.strip()
)


# =====================================================
# ATUALIZA PLANILHA PRINCIPAL
# =====================================================

if ATUALIZAR_PLANILHA:

    print("Atualizando dados_gerais_emendas.xlsx...")

    # Remove a coluna Link caso já exista
    dados = dados.drop(columns=["Link"], errors="ignore")

    # Cria um dicionário Código SIAFI -> Inteiro Teor
    mapa = dict(zip(
        links["Código SIAFI"].astype(str).str.strip(),
        links["Inteiro Teor"]
    ))

    # Cria a coluna Link
    dados["Link"] = (
        dados["CÓDIGO SIAFI"]
        .astype(str)
        .str.strip()
        .map(mapa)
    )

    # Salva o arquivo
    dados.to_excel(
        ARQ_DADOS,
        index=False
    )

    print(f"Planilha atualizada. {dados['Link'].notna().sum()} links encontrados.")

# =====================================================
# GITHUB
# =====================================================

if ATUALIZAR_GITHUB:

    print("Enviando alterações para o GitHub...")

    subprocess.run(["git", "add", "."], cwd=BASE)

    subprocess.run(
        [
            "git",
            "commit",
            "-m",
            "Atualização automática do Inteiro Teor"
        ],
        cwd=BASE
    )

    subprocess.run(
        ["git", "push"],
        cwd=BASE
    )

    print("GitHub atualizado.")


print("Concluído.")
