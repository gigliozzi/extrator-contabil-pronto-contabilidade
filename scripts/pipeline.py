#!/usr/bin/env python3
"""
===============================================
Script : pipeline.py
Descrição : Orquestra o processamento de PDFs:
            1. Varre a pasta de entrada
            2. Detecta se o PDF é nativo ou escaneado
            3. Roteia para o parser correto:
               - RENAVAM (escaneado) → parser_renavam.py
               - Multi-documento (nativo) → parser_multi_doc.py
            4. Salva JSON + gera Excel
            5. Move PDF para processados ou erros
Versão : v2.0 - 2026-05-06
Autor  : William Gigliozzi
===============================================

Uso:
    python pipeline.py                    # processa todos os PDFs em entrada/
    python pipeline.py --excel            # processa e gera Excel ao final
    python pipeline.py --arquivo x.pdf   # processa um arquivo específico
"""

import os, sys, json, shutil, logging, argparse, traceback
from datetime import datetime
from pathlib import Path

# ── Configuração do Tesseract (necessário para PDFs escaneados) ──
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# ── Caminhos ──────────────────────────────────────────────────────
BASE_DIR        = Path(__file__).parent.parent
DIR_ENTRADA     = BASE_DIR / "entrada"
DIR_PROCESSADOS = BASE_DIR / "processados"
DIR_ERROS       = BASE_DIR / "erros"
DIR_LOGS        = BASE_DIR / "logs"
DIR_RELATORIOS  = BASE_DIR / "relatorios"

for d in [DIR_ENTRADA, DIR_PROCESSADOS, DIR_ERROS, DIR_LOGS, DIR_RELATORIOS]:
    d.mkdir(exist_ok=True)

# ── Log ───────────────────────────────────────────────────────────
DATA_HOJE   = datetime.now().strftime("%Y-%m-%d")
ARQUIVO_LOG = DIR_LOGS / f"pipeline_{DATA_HOJE}.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(ARQUIVO_LOG, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ]
)
log = logging.getLogger("pipeline")

# ── Importa parsers ───────────────────────────────────────────────
sys.path.insert(0, str(Path(__file__).parent))
from parser_renavam   import extrair          as extrair_renavam
from parser_multi_doc import processar_pdf    as extrair_multi
from parser_multi_doc import gerar_excel      as gerar_excel_multi


# ══════════════════════════════════════════════════════════════════
# DETECÇÃO DE TIPO DE PDF
# ══════════════════════════════════════════════════════════════════

def tem_texto_nativo(caminho_pdf: str, min_chars: int = 100) -> bool:
    """
    Verifica se o PDF tem texto nativo embutido.
    PDFs escaneados retornam texto vazio ou quase vazio.
    """
    import fitz
    try:
        doc = fitz.open(caminho_pdf)
        total = sum(len(doc[i].get_text().strip()) for i in range(min(3, len(doc))))
        doc.close()
        return total >= min_chars
    except Exception:
        return False


def detectar_tipo_escaneado(caminho_pdf: str) -> str:
    """
    Para PDFs escaneados (sem texto nativo), usa OCR rápido
    na primeira página para identificar o tipo.
    """
    import fitz, io
    from PIL import Image, ImageEnhance
    try:
        doc = fitz.open(caminho_pdf)
        pix = doc[0].get_pixmap(matrix=fitz.Matrix(150/72, 150/72))
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        doc.close()
        img = ImageEnhance.Contrast(img.convert("L")).enhance(2.0)
        texto = pytesseract.image_to_string(img, lang="por", config="--oem 3 --psm 3").upper()
        if "RENAVAM" in texto or "NADA CONSTA DO VE" in texto:
            return "RENAVAM"
        return "ESCANEADO_DESCONHECIDO"
    except Exception as e:
        log.error(f"Erro ao detectar tipo escaneado: {e}")
        return "ERRO"


# ══════════════════════════════════════════════════════════════════
# PROCESSAMENTO
# ══════════════════════════════════════════════════════════════════

def processar_pdf(caminho_pdf: Path):
    nome = caminho_pdf.name
    log.info(f"{'─'*60}")
    log.info(f"🔎 Iniciando: {nome}")

    try:
        # ── 1. Detecta se é nativo ou escaneado ───────────────────
        if tem_texto_nativo(str(caminho_pdf)):
            modo = "NATIVO"
            log.info("📄 PDF nativo (texto embutido) → parser_multi_doc")
        else:
            modo = detectar_tipo_escaneado(str(caminho_pdf))
            log.info(f"📄 PDF escaneado → tipo: {modo}")

        # ── 2. Roteia para o parser correto ───────────────────────
        if modo == "NATIVO":
            registros = extrair_multi(str(caminho_pdf))
            resultado = {
                "tipo_processamento" : "MULTI_DOCUMENTO_NATIVO",
                "arquivo"            : nome,
                "total_registros"    : len(registros),
                "registros"          : registros,
            }
            resumo = f"{len(registros)} registro(s)"

        elif modo == "RENAVAM":
            resultado = extrair_renavam(str(caminho_pdf))
            resultado["tipo_processamento"] = "RENAVAM_OCR"
            resumo = f"{resultado.get('total_infracoes', 0)} infração(ões) | {resultado['tempo_segundos']}s"

        else:
            log.warning(f"⚠️  Tipo '{modo}' sem parser implementado.")
            _mover(caminho_pdf, DIR_ERROS, sufixo="_tipo_desconhecido")
            return None

        # ── 3. Salva JSON de resultado ────────────────────────────
        nome_json    = caminho_pdf.stem + f"_{DATA_HOJE}.json"
        caminho_json = DIR_PROCESSADOS / nome_json
        with open(caminho_json, "w", encoding="utf-8") as f:
            json.dump(resultado, f, ensure_ascii=False, indent=2)

        # ── 4. Move PDF para processados ──────────────────────────
        _mover(caminho_pdf, DIR_PROCESSADOS)
        log.info(f"✅ Concluído: {nome} → {resumo}")
        return resultado

    except Exception as e:
        log.error(f"❌ ERRO ao processar {nome}: {e}")
        log.debug(traceback.format_exc())
        _mover(caminho_pdf, DIR_ERROS, sufixo="_erro")
        return None


def _mover(origem: Path, destino_dir: Path, sufixo: str = ""):
    nome    = origem.stem + sufixo + origem.suffix
    destino = destino_dir / nome
    contador = 1
    while destino.exists():
        nome    = f"{origem.stem}{sufixo}_{contador}{origem.suffix}"
        destino = destino_dir / nome
        contador += 1
    shutil.move(str(origem), str(destino))
    log.info(f"📁 Movido para: {destino.parent.name}/{destino.name}")


# ══════════════════════════════════════════════════════════════════
# GERAÇÃO DO EXCEL
# ══════════════════════════════════════════════════════════════════

def gerar_excel(resultados: list):
    """
    Consolida todos os resultados do dia em planilhas Excel.
    - PDFs nativos (multi-doc) → aba formato cliente (Nome/DATA/DOC/COMP/VALOR)
    - PDFs RENAVAM (OCR)       → aba RENAVAM (formato anterior)
    """
    from gerar_excel import gerar as gerar_renavam_excel

    # Separa por tipo de processamento
    registros_multi  = []
    resultados_renavam = []

    for r in resultados:
        if r.get("tipo_processamento") == "MULTI_DOCUMENTO_NATIVO":
            registros_multi.extend(r.get("registros", []))
        elif r.get("tipo_processamento") == "RENAVAM_OCR":
            resultados_renavam.append(r)

    # Gera Excel multi-documento (formato cliente)
    if registros_multi:
        nome_arquivo = f"relatorio_documentos_{DATA_HOJE}.xlsx"
        caminho      = DIR_RELATORIOS / nome_arquivo
        gerar_excel_multi(registros_multi, str(caminho))
        log.info(f"📊 Excel multi-doc: {caminho.name}")

    # Gera Excel RENAVAM (formato anterior com abas por tipo)
    if resultados_renavam:
        try:
            gerar_renavam_excel()
            log.info("📊 Excel RENAVAM gerado.")
        except Exception as e:
            log.error(f"Erro ao gerar Excel RENAVAM: {e}")


# ══════════════════════════════════════════════════════════════════
# PIPELINE PRINCIPAL
# ══════════════════════════════════════════════════════════════════

def executar(arquivo_especifico: str = None, gerar_excel_flag: bool = False):
    inicio = datetime.now()
    log.info("=" * 60)
    log.info(f"🚀 PIPELINE INICIADO — {inicio.strftime('%d/%m/%Y %H:%M:%S')}")
    log.info("=" * 60)

    # Lista PDFs
    if arquivo_especifico:
        pdfs = [DIR_ENTRADA / arquivo_especifico]
        pdfs = [p for p in pdfs if p.exists()]
        if not pdfs:
            log.error(f"❌ Arquivo não encontrado: {arquivo_especifico}")
            return
    else:
        vistos, pdfs = set(), []
        for p in sorted(DIR_ENTRADA.glob("*.pdf")) + sorted(DIR_ENTRADA.glob("*.PDF")):
            if p.name.lower() not in vistos:
                vistos.add(p.name.lower())
                pdfs.append(p)

    if not pdfs:
        log.info("📭 Nenhum PDF encontrado na pasta entrada/. Encerrando.")
        return

    log.info(f"📂 PDFs encontrados: {len(pdfs)}")

    # Processa
    sucessos, erros, resultados = 0, 0, []
    for pdf in pdfs:
        resultado = processar_pdf(pdf)
        if resultado:
            sucessos += 1
            resultados.append(resultado)
        else:
            erros += 1

    # Resumo
    duracao = round((datetime.now() - inicio).total_seconds(), 1)
    log.info("=" * 60)
    log.info(f"✅ Concluídos : {sucessos}")
    log.info(f"❌ Com erro   : {erros}")
    log.info(f"⏱️  Duração    : {duracao}s")
    log.info("=" * 60)

    # Excel
    if gerar_excel_flag and resultados:
        log.info("📊 Gerando planilhas Excel...")
        try:
            gerar_excel(resultados)
        except Exception as e:
            log.error(f"❌ Erro ao gerar Excel: {e}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Pipeline OCR de documentos")
    parser.add_argument("--arquivo", default=None,
                        help="Nome do PDF em entrada/ (ex: doc.pdf)")
    parser.add_argument("--excel", action="store_true",
                        help="Gera planilha Excel ao final")
    args = parser.parse_args()
    executar(arquivo_especifico=args.arquivo, gerar_excel_flag=args.excel)
