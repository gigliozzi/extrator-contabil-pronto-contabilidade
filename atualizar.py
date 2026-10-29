#!/usr/bin/env python3
"""
===============================================
Script : atualizar.py
Descrição : Atualiza os arquivos do projeto
            baixando a versão mais recente do
            repositório GitHub.
            Execute com duplo clique quando
            orientado pelo suporte.
Versão : v1.0 - 2026-10-29
===============================================
"""

import urllib.request
import shutil
import sys
from pathlib import Path

REPO_RAW = "https://raw.githubusercontent.com/gigliozzi/extrator-contabil-pronto-contabilidade/main"

ARQUIVOS = [
    ("scripts/extrator_visao_claude.py", "scripts/extrator_visao_claude.py"),
    ("scripts/interface.py",             "scripts/interface.py"),
]

BASE_DIR = Path(__file__).parent


def baixar(url, destino):
    try:
        with urllib.request.urlopen(url, timeout=30) as resp:
            conteudo = resp.read()
        with open(destino, "wb") as f:
            f.write(conteudo)
        return True
    except Exception as e:
        print(f"   ❌ Erro: {e}")
        return False


def main():
    print("=" * 55)
    print("  Atualizador — Extrator de Documentos Contábeis")
    print("  Studio Delta")
    print("=" * 55)
    print()

    sucessos = 0
    erros = 0

    for origem, destino_rel in ARQUIVOS:
        destino = BASE_DIR / destino_rel
        url = f"{REPO_RAW}/{origem}"

        print(f"  Atualizando: {destino_rel} ...", end=" ", flush=True)

        # Backup do arquivo atual
        if destino.exists():
            shutil.copy2(destino, str(destino) + ".bak")

        if baixar(url, destino):
            print("✅")
            sucessos += 1
        else:
            # Restaurar backup em caso de erro
            bak = Path(str(destino) + ".bak")
            if bak.exists():
                shutil.copy2(bak, destino)
                print("   ↩ Arquivo anterior restaurado.")
            erros += 1

    print()
    print("=" * 55)
    print(f"  Concluído: {sucessos} atualizado(s), {erros} erro(s)")
    print("=" * 55)

    if erros == 0:
        print("\n  ✅ Sistema atualizado com sucesso!")
    else:
        print("\n  ⚠️  Verifique os erros acima e contate o suporte.")

    print()
    input("  Pressione Enter para fechar...")


if __name__ == "__main__":
    main()
