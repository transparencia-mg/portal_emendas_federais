"""
Atualiza a coluna Link (Inteiro Teor) de upload/dados_gerais_emendas.xlsx e confere
as três planilhas da pasta upload (não envia ao GitHub).

Fica na pasta scripts do repositório. Normalmente é chamado pelo atualizar.py
(na raiz), que também prepara as bases e envia tudo ao GitHub. Para rodar só esta etapa:

    python scripts/atualizar_links.py
"""
import re
import sys
from pathlib import Path

import pandas as pd


# =====================================================
# CONFIGURAÇÃO
# =====================================================

BASE = Path(__file__).resolve().parent.parent   # raiz do repositório (este arquivo fica em scripts/)
PASTA_UPLOAD = BASE / "upload"

ARQ_DADOS    = PASTA_UPLOAD / "dados_gerais_emendas.xlsx"
ARQ_EXECUCAO = PASTA_UPLOAD / "execucao_transferencias_especiais_pix.xlsx"
ARQ_PLANO    = PASTA_UPLOAD / "plano_execucao_emendas_pix.xlsx"

# Fonte dos links (Inteiro Teor), testada nesta ordem:
#   1) ARQ_LINKS, se você informar um arquivo próprio, ex.: BASE / "links_inteiro_teor.xlsx"
#   2) convenios.csv do repositório portal_convenios_entrada clonado ao lado deste
#   3) download pela internet (pode ser bloqueado pelo proxy da rede)
# Se nenhuma funcionar, os links que já estão na planilha são mantidos.
ARQ_LINKS = None
ARQ_LINKS_LOCAL = BASE.parent / "portal_convenios_entrada" / "dataset" / "data" / "convenios.csv"
URL_LINKS = "https://raw.githubusercontent.com/transparencia-mg/portal_convenios_entrada/main/dataset/data/convenios.csv"

ATUALIZAR_PLANILHA = True


# =====================================================
# FUNÇÕES AUXILIARES
# =====================================================

class ErroLinks(Exception):
    pass


def normalizar_siafi(valor):
    """9287341, '9287341', '9287341.0' e ' 9287341 ' viram '9287341'."""
    if pd.isna(valor):
        return ""
    s = str(valor).strip()
    s = re.sub(r"[.,]0+$", "", s)
    return re.sub(r"\D", "", s)


def achar_coluna(df, opcoes, arquivo):
    """Encontra a coluna aceitando variações de nome."""
    normal = {c.strip().lower(): c for c in df.columns}
    for op in opcoes:
        if op.lower() in normal:
            return normal[op.lower()]
    raise ErroLinks(f"Coluna {opcoes[0]!r} não encontrada em {arquivo}. Colunas: {list(df.columns)}")


def ler_links_arquivo(caminho):
    caminho = Path(caminho)
    if caminho.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(caminho)
    return pd.read_csv(caminho, sep=";", dtype=str)


def carregar_links():
    """Retorna (DataFrame, origem, coluna SIAFI, coluna Inteiro Teor) ou (None, ...) se não houver fonte."""
    links, origem = None, None
    if ARQ_LINKS and Path(ARQ_LINKS).exists():
        links, origem = ler_links_arquivo(ARQ_LINKS), Path(ARQ_LINKS).name
    elif ARQ_LINKS:
        print(f"  AVISO: ARQ_LINKS não encontrado: {ARQ_LINKS}")

    if links is None and ARQ_LINKS_LOCAL.exists():
        links, origem = ler_links_arquivo(ARQ_LINKS_LOCAL), f"{ARQ_LINKS_LOCAL} (cópia local)"

    if links is None:
        print("  Baixando links da base de convênios...")
        try:
            links, origem = pd.read_csv(URL_LINKS, sep=";", dtype=str), "convenios.csv (GitHub)"
        except Exception as erro:
            motivo = "o proxy da rede exige autenticação (erro 407)" if "407" in str(erro) else str(erro)
            print(f"  AVISO: não foi possível baixar os links: {motivo}.")
            print("  Para resolver, informe um arquivo em ARQ_LINKS ou atualize o repositório")
            print(f"  portal_convenios_entrada em {ARQ_LINKS_LOCAL.parent.parent.parent} (git pull).")
            return None, None, None, None

    print(f"  Links lidos de: {origem}")
    links.columns = links.columns.str.strip()
    col_siafi = achar_coluna(links, ["Código SIAFI", "codigo_siafi", "CÓDIGO SIAFI"], origem)
    col_teor  = achar_coluna(links, ["Inteiro Teor", "inteiro_teor"], origem)
    links[col_siafi] = links[col_siafi].map(normalizar_siafi)
    return links, origem, col_siafi, col_teor


# =====================================================
# ETAPAS
# =====================================================

def atualizar_links(dados):
    """Preenche a coluna Link. Retorna o DataFrame (com SIAFI normalizado)."""
    links, _, col_siafi, col_teor = carregar_links()
    if links is None:
        print("  Os links atuais da planilha serão mantidos; a atualização dos links foi pulada.")
        return dados

    # Código SIAFI -> Inteiro Teor (ignora links vazios; mantém o primeiro de cada código)
    validos = links[links[col_teor].notna() & (links[col_teor].astype(str).str.strip() != "")]
    mapa = (validos.drop_duplicates(col_siafi)
                   .set_index(col_siafi)[col_teor]
                   .astype(str).str.strip()
                   .to_dict())

    link_anterior = dados["Link"].copy() if "Link" in dados.columns else None
    dados = dados.drop(columns=["Link"], errors="ignore")
    dados["Link"] = dados["CÓDIGO SIAFI"].map(mapa)

    # Grava o SIAFI como número inteiro, como na planilha original
    siafi_original = dados["CÓDIGO SIAFI"]
    dados["CÓDIGO SIAFI"] = pd.to_numeric(dados["CÓDIGO SIAFI"], errors="coerce").astype("Int64")

    # Só regrava o Excel se algum link mudou (evita commits vazios a cada execução)
    if link_anterior is not None and link_anterior.fillna("").astype(str).tolist() == dados["Link"].fillna("").astype(str).tolist():
        print("  Links já estavam atualizados; planilha não foi regravada.")
    else:
        try:
            dados.to_excel(ARQ_DADOS, index=False)
        except PermissionError:
            raise ErroLinks(f"Não foi possível gravar {ARQ_DADOS.name}. Feche o arquivo no Excel e rode de novo.")
        print(f"  {ARQ_DADOS.name} atualizado.")

    com_link = dados["Link"].notna().sum()
    print(f"  Emendas com link: {com_link} de {len(dados)}.")
    sem_link = dados.loc[dados["Link"].isna(), "CÓDIGO SIAFI"].dropna().unique()
    if len(sem_link):
        print(f"  {len(sem_link)} Nº SIAFI sem link. Exemplos: {', '.join(map(str, sem_link[:10]))}")

    dados["CÓDIGO SIAFI"] = siafi_original
    return dados


def conferir_bases(dados, execucao, plano):
    """Só avisa sobre inconsistências entre as três bases; não altera nada."""
    print("Conferindo as bases...")
    siafi_g = set(dados["CÓDIGO SIAFI"].map(normalizar_siafi)) - {""}
    col_siafi_exec = achar_coluna(execucao, ["Nº SIAFI", "no_siafi"], ARQ_EXECUCAO.name)
    siafi_e = set(execucao[col_siafi_exec].map(normalizar_siafi)) - {""}

    try:
        col_siafi_plano = achar_coluna(plano, ["Nº SIAFI", "no_siafi"], ARQ_PLANO.name)
        siafi_p_serie = plano[col_siafi_plano].map(normalizar_siafi)
        siafi_p = set(siafi_p_serie) - {""}
        vazios = (siafi_p_serie == "").sum()
        if vazios:
            print(f"  AVISO: {vazios} linha(s) do plano sem Nº SIAFI.")
    except ErroLinks:
        siafi_p = set()
        print("  AVISO: o plano não tem a coluna 'Nº SIAFI'. O painel vai ligar o plano por estimativa.")

    for col in ("valor custeio", "valor capital", "valor da emenda"):
        if col in plano.columns:
            texto = plano[col].map(lambda v: isinstance(v, str)).sum()
            if texto:
                print(f"  AVISO: {texto} valor(es) gravado(s) como texto na coluna '{col}' do plano.")

    todos = siafi_g | siafi_e | siafi_p
    nas3 = len(siafi_g & siafi_e & siafi_p)
    em1 = sum(1 for k in todos if (k in siafi_g) + (k in siafi_e) + (k in siafi_p) == 1)
    print(f"  Nº SIAFI distintos: {len(todos)} | nas 3 bases: {nas3} | em 2: {len(todos) - nas3 - em1} | em 1: {em1}")
    fora = sorted((siafi_e | siafi_p) - siafi_g)
    if fora:
        print(f"  {len(fora)} Nº SIAFI na execução/plano sem cadastro em Emendas Geral: "
              f"{', '.join(fora[:12])}{' ...' if len(fora) > 12 else ''}")


def executar():
    """Atualiza os links e confere as bases. Lança ErroLinks se algo impedir a execução."""
    for arq in (ARQ_DADOS, ARQ_EXECUCAO, ARQ_PLANO):
        if not arq.exists():
            raise ErroLinks(f"Arquivo não encontrado: {arq}")

    print("Lendo planilhas...")
    try:
        dados    = pd.read_excel(ARQ_DADOS)
        execucao = pd.read_excel(ARQ_EXECUCAO)
        plano    = pd.read_excel(ARQ_PLANO)
    except PermissionError as erro:
        raise ErroLinks(f"Não foi possível abrir {Path(erro.filename).name}. Feche o arquivo no Excel e rode de novo.")
    for df in (dados, execucao, plano):
        df.columns = df.columns.astype(str).str.strip()

    achar_coluna(dados, ["CÓDIGO SIAFI"], ARQ_DADOS.name)
    dados["CÓDIGO SIAFI"] = dados["CÓDIGO SIAFI"].map(normalizar_siafi)

    if ATUALIZAR_PLANILHA:
        print("Atualizando links de dados_gerais_emendas.xlsx...")
        dados = atualizar_links(dados)

    conferir_bases(dados, execucao, plano)


def main():
    print(f"Repositório: {BASE}")
    try:
        executar()
    except ErroLinks as erro:
        sys.exit(f"ERRO: {erro}")
    print("Concluído (sem envio ao GitHub; para enviar, use: python atualizar.py).")


if __name__ == "__main__":
    main()
