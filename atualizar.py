"""
Atualização completa do painel de Emendas Federais.

Fica na RAIZ do repositório portal_emendas_federais. Rode:

    python atualizar.py

Etapas:
 1. git pull — sincroniza com o GitHub antes de mexer em qualquer arquivo.
 2. scripts/atualizar_bases.py — gera datas.xlsx e as planilhas da pasta upload a partir de
    EMENDASFEDERAIS.xlsx e de "Dados Emendas - ATUALIZADO <data>.xlsx" (se estiverem na pasta upload).
 3. scripts/atualizar_links.py — preenche a coluna Link (Inteiro Teor) de dados_gerais_emendas.xlsx
    e confere as bases.
 4. Apaga os arquivos de origem (só se as etapas 2 e 3 deram certo).
 5. git add / commit / push. No GitHub, a automação converte a pasta upload em CSV e publica os dados.

Se qualquer etapa falhar, o script para: os arquivos de origem não são apagados e nada é enviado.
"""
import subprocess
import sys
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE / "scripts"))
sys.dont_write_bytecode = True   # não cria scripts/__pycache__ (não vai parar no GitHub)

import atualizar_bases   # noqa: E402
import atualizar_links   # noqa: E402


# =====================================================
# CONFIGURAÇÃO
# =====================================================

ATUALIZAR_BASES   = True    # etapa 2
ATUALIZAR_LINKS   = True    # etapa 3
APAGAR_ORIGINAIS  = True    # etapa 4
ATUALIZAR_GITHUB  = True    # etapas 1 e 5
MENSAGEM_COMMIT   = "Atualização das emendas federais"


# =====================================================
# GIT
# =====================================================

def git(*args, mostrar=True):
    r = subprocess.run(["git", *args], cwd=BASE, capture_output=True, text=True)
    if mostrar and r.stdout.strip():
        print("    " + r.stdout.strip().replace("\n", "\n    "))
    if r.returncode != 0 and r.stderr.strip():
        print("    " + r.stderr.strip().replace("\n", "\n    "))
    return r


def parar(msg):
    print(f"\nERRO: {msg}")
    sys.exit("Processo interrompido. Os arquivos de origem NÃO foram apagados e nada foi enviado ao GitHub.")


# =====================================================
# EXECUÇÃO
# =====================================================

def main():
    print(f"Repositório: {BASE}")

    usar_git = ATUALIZAR_GITHUB and git("rev-parse", "--is-inside-work-tree", mostrar=False).returncode == 0
    if ATUALIZAR_GITHUB and not usar_git:
        print("  AVISO: esta pasta não é um repositório git; os arquivos serão preparados, mas não enviados.")

    # 1) Sincroniza antes de alterar qualquer arquivo
    if usar_git:
        print("\n[1/5] Sincronizando com o GitHub (git pull)...")
        if git("pull", "--rebase", "--autostash").returncode != 0:
            sys.exit("Falha no git pull. Resolva (ex.: conflito ou sem internet) e rode de novo. Nenhum arquivo foi alterado.")

    # 2) Bases
    origens = []
    if ATUALIZAR_BASES:
        print("\n[2/5] Preparando as bases...")
        try:
            _, origens = atualizar_bases.executar()
        except atualizar_bases.ErroProcessamento as erro:
            parar(erro)
        if not origens:
            print("  Nenhum arquivo novo na pasta upload; as bases atuais serão mantidas.")

    # 3) Links
    if ATUALIZAR_LINKS:
        print("\n[3/5] Atualizando links do Inteiro Teor e conferindo as bases...")
        try:
            atualizar_links.executar()
        except atualizar_links.ErroLinks as erro:
            parar(erro)

    # 4) Apaga os arquivos de origem
    if APAGAR_ORIGINAIS and origens:
        print("\n[4/5] Apagando arquivos de origem...")
        if not atualizar_bases.apagar_originais(origens) and usar_git:
            sys.exit("Feche o arquivo no Excel, apague-o da pasta upload e rode de novo "
                     "(ele não pode ir para o GitHub, senão a automação vai convertê-lo em CSV).")

    # 5) Commit e push
    if not usar_git:
        print("\nConcluído (sem envio ao GitHub).")
        return

    print("\n[5/5] Enviando alterações para o GitHub...")
    git("add", "-A")
    status = git("status", "--porcelain", mostrar=False).stdout.strip()
    if not status:
        print("  Nenhuma alteração para enviar (tudo já estava igual ao GitHub).")
        print("\nConcluído.")
        return
    print("  Alterações:")
    print("    " + status.replace("\n", "\n    "))

    if git("commit", "-m", f"{MENSAGEM_COMMIT} ({datetime.now():%d/%m/%Y %H:%M})").returncode != 0:
        sys.exit("Falha no commit.")
    if git("push").returncode != 0:
        # O GitHub pode ter commits novos (a automação grava os CSV): sincroniza e tenta de novo
        print("  Push recusado; sincronizando e tentando de novo...")
        if git("pull", "--rebase").returncode != 0 or git("push").returncode != 0:
            sys.exit("Falha no git push. O commit foi feito localmente; resolva e rode 'git push'.")
    print("  GitHub atualizado. A automação do repositório vai gerar os CSV e publicar os dados.")
    print("\nConcluído.")


if __name__ == "__main__":
    main()
