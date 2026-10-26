#!/usr/bin/env python3
"""
===============================================
Script : instalar_dependencias.py
Descrição : Verifica e instala as dependências
            necessárias para o projeto OCR
Versão : v1.0 - 2026-05-04
===============================================

Execute este script UMA VEZ antes de usar o pipeline.
"""

import subprocess
import sys
import shutil

print("\n" + "="*60)
print("  INSTALAÇÃO DE DEPENDÊNCIAS — OCR de Documentos")
print("="*60)

# ── 1. Python packages ─────────────────────────────────────────
packages = [
    ("pymupdf",      "fitz",         "PyMuPDF — converte PDF em imagem"),
    ("pytesseract",  "pytesseract",  "pytesseract — interface Python para Tesseract"),
    ("pillow",       "PIL",          "Pillow — manipulação de imagens"),
    ("numpy",        "numpy",        "NumPy — análise de regiões da imagem"),
    ("openpyxl",     "openpyxl",     "openpyxl — geração de arquivos Excel"),
]

print("\n📦 Verificando pacotes Python...")
for pkg_install, pkg_import, descricao in packages:
    try:
        __import__(pkg_import)
        print(f"  ✅ {pkg_install:<15} já instalado")
    except ImportError:
        print(f"  ⬇️  Instalando {pkg_install} ({descricao})...")
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", pkg_install],
            capture_output=True, text=True
        )
        if result.returncode == 0:
            print(f"  ✅ {pkg_install} instalado com sucesso")
        else:
            print(f"  ❌ Erro ao instalar {pkg_install}:")
            print(f"     {result.stderr.strip()}")

# ── 2. Tesseract OCR ───────────────────────────────────────────
print("\n🔍 Verificando Tesseract OCR...")
tesseract_path = shutil.which("tesseract")
if tesseract_path:
    result = subprocess.run(["tesseract", "--version"],
                            capture_output=True, text=True)
    versao = result.stdout.splitlines()[0] if result.stdout else "desconhecida"
    print(f"  ✅ Tesseract encontrado: {versao}")
    print(f"     Caminho: {tesseract_path}")
else:
    print("  ❌ Tesseract NÃO encontrado!")
    print()
    print("  Para instalar no Windows:")
    print("  1. Acesse: https://github.com/UB-Mannheim/tesseract/wiki")
    print("  2. Baixe o instalador: tesseract-ocr-w64-setup-5.x.x.exe")
    print("  3. Durante a instalação, marque 'Additional language data'")
    print("     e selecione 'Portuguese'")
    print("  4. Adicione ao PATH: C:\\Program Files\\Tesseract-OCR")
    print()
    print("  Após instalar, execute este script novamente.")
    sys.exit(1)

# ── 3. Idioma Português ────────────────────────────────────────
print("\n🌐 Verificando suporte ao Português...")
result = subprocess.run(["tesseract", "--list-langs"],
                        capture_output=True, text=True)
langs = result.stdout + result.stderr
if "por" in langs:
    print("  ✅ Idioma Português (por) disponível")
else:
    print("  ❌ Idioma Português NÃO instalado!")
    print()
    print("  Para instalar no Windows:")
    print("  Execute novamente o instalador do Tesseract e")
    print("  marque o pacote de idioma 'Portuguese'.")
    print()
    print("  OU baixe manualmente:")
    print("  https://github.com/tesseract-ocr/tessdata/blob/main/por.traineddata")
    print("  Coloque o arquivo em: C:\\Program Files\\Tesseract-OCR\\tessdata\\")
    sys.exit(1)

# ── Resumo final ───────────────────────────────────────────────
print()
print("="*60)
print("  ✅ AMBIENTE PRONTO PARA USO")
print("="*60)
print()
print("  Para processar PDFs:")
print("  1. Coloque os PDFs na pasta:  entrada\\")
print("  2. Execute o pipeline:")
print("       python scripts\\pipeline.py")
print("  3. Para gerar o Excel junto:")
print("       python scripts\\pipeline.py --excel")
print()
print("  Estrutura de pastas:")
print("   entrada\\      → coloque os PDFs aqui")
print("   processados\\ → PDFs processados + JSONs de resultado")
print("   relatorios\\  → planilhas Excel geradas")
print("   erros\\       → PDFs que falharam no processamento")
print("   logs\\        → logs de execução por data")
print()
