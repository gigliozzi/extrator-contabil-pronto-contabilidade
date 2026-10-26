#!/usr/bin/env python3
"""
===============================================
Script : parser_renavam.py
Descrição : OCR e extração de campos do documento
            "Nada Consta do Veículo (RENAVAM)"
Versão : v6.0 - 2026-05-04
Autor  : William Gigliozzi
===============================================
"""

import os
import re
import io
import json
import time
import logging

import fitz                        # PyMuPDF
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
import numpy as np
from PIL import Image, ImageEnhance, ImageFilter

# ── Configurações de OCR ───────────────────────────────────────
DPI    = 400      # resolução de conversão PDF → imagem
IDIOMA = "por"    # idioma do Tesseract

# ──────────────────────────────────────────────────────────────

log = logging.getLogger("parser_renavam")


# ── Utilitários de imagem ──────────────────────────────────────

def _prep(crop: Image.Image, contraste: float = 2.5) -> Image.Image:
    """Converte para escala de cinza e melhora contraste/nitidez."""
    crop = crop.convert("L")
    crop = ImageEnhance.Contrast(crop).enhance(contraste)
    return crop.filter(ImageFilter.SHARPEN)


def _ocr(img: Image.Image, y1: int, y2: int,
         psm: int = 4, contraste: float = 2.5,
         x1: int = 0, x2: int = None) -> str:
    """Recorta uma região da imagem e executa OCR."""
    if x2 is None:
        x2 = img.width
    y2 = min(y2, img.height)
    if y1 >= y2:
        return ""
    recorte = img.crop((x1, y1, x2, y2))
    return pytesseract.image_to_string(
        _prep(recorte, contraste),
        lang=IDIOMA,
        config=f"--oem 3 --psm {psm}"
    )


# ── Normalizadores de campo ────────────────────────────────────

def _norm_data(s: str) -> str:
    m = re.search(r"(\d{2}/\d{2}/\d{4})", s)
    return m.group(1) if m else ""


def _norm_hora(s: str) -> str:
    m = re.search(r"\b(\d{2}:\d{2})\b", s)
    if m:
        return m.group(1)
    # OCR às vezes remove os dois-pontos: "1840" → "18:40"
    m2 = re.search(r"\b(0\d|1\d|2[0-3])(\d{2})\b", s)
    if m2 and 0 <= int(m2.group(2)) <= 59:
        return f"{m2.group(1)}:{m2.group(2)}"
    return ""


def _norm_placa(s: str) -> str:
    s = re.sub(r"[^A-Za-z0-9]", "", s).upper()
    return s[:7] if len(s) >= 7 else s


def _norm_valor(s: str) -> str:
    """Normaliza valor monetário: garante formato NNN,NN."""
    m = re.search(r"(\d{1,4}[,.]\d{2})", s)
    if m:
        return m.group(1).replace(".", ",")
    m2 = re.search(r"\b(\d{3,6})\b", s)
    if m2:
        v = m2.group(1)
        return v[:-2] + "," + v[-2:]
    return ""


# ── Detecção de blocos de infração ────────────────────────────

def _detectar_blocos(img: Image.Image) -> list:
    """
    Retorna lista de (y_inicio, y_fim) de cada cabeçalho escuro
    (fundo cinza) que separa os blocos de infração.
    """
    arr = np.array(img.convert("L"))
    h, w = arr.shape
    blocos, dentro, ys = [], False, 0
    for y in range(0, h, 10):
        brilho = arr[y:y + 10, 100:w - 100].mean()
        if brilho < 170 and not dentro:
            dentro, ys = True, max(0, y - 20)
        elif brilho > 210 and dentro:
            dentro = False
            blocos.append((ys, y + 20))
    return blocos


# ── Extração de um bloco de infração ──────────────────────────

def _extrair_bloco(img: Image.Image, yc2: int, y_fim: int) -> dict:
    """
    Divide o espaço do bloco em zonas e extrai cada campo
    com o PSM mais adequado para aquela região.
    """
    esp = y_fim - yc2
    W   = img.width

    # Zonas proporcionais ao espaço disponível
    zA = (yc2,              yc2 + int(esp * 0.22))   # auto/renainf/data pgto
    zB = (zA[1],            zA[1] + int(esp * 0.18)) # data infração + hora
    zC = (zB[1],            zB[1] + int(esp * 0.22)) # descrição + placa
    zD = (zC[1],            zC[1] + int(esp * 0.30)) # local + valores
    zE = (zD[1],            y_fim)                   # status pagamento

    tA  = _ocr(img, *zA, psm=4, contraste=3.0)
    tB  = _ocr(img, *zB, psm=4, contraste=3.0)
    tC  = _ocr(img, *zC, psm=4, contraste=2.5)
    tD  = _ocr(img, *zD, psm=6, contraste=2.5)          # PSM 6 para linha de local+valores
    tD2 = _ocr(img, *zD, psm=4, contraste=3.0, x2=W//2) # fallback: metade esquerda
    tE  = _ocr(img, *zE, psm=4, contraste=2.5)

    # ── Auto de Infração / Renainf / Data pgto ─────────────────
    auto_inf = auto_ren = data_pgto = ""
    mr = re.search(r"\b([A-Z]{2}\d{7,10})\b", tA)
    mi = re.search(r"\b(\d{7,12})\b", tA)
    md = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", tA)
    auto_ren  = mr.group(1) if mr else ""
    if mi:
        v = mi.group(1)
        if not mr or v != mr.group(1):
            auto_inf = v
    data_pgto = md.group(1) if md else ""

    # ── Data da infração e hora ────────────────────────────────
    data_inf = hora = ""
    todos_ab  = tA + "\n" + tB
    datas     = re.findall(r"\d{2}/\d{2}/\d{4}", todos_ab)
    for d in datas:
        if d != data_pgto:
            data_inf = d
            break
    hora = _norm_hora(tB) or _norm_hora(tA)

    # ── Descrição e placa ──────────────────────────────────────
    descricao = placa = ""
    kw = r"(SINAL|ESTACIONAR|AVANC|VELOCIDADE|ULTRAPASS|PARAR|FAIXA|CINTO|CELULAR|MOVIMENTAC|SEMAF)"
    for txt in [tA, tB, tC]:
        for l in txt.splitlines():
            if re.search(kw, l, re.I):
                mp = re.search(r"\b([A-Z]{3}[0-9][A-Z0-9][0-9]{2})\b", l)
                if mp:
                    placa     = _norm_placa(mp.group(1))
                    descricao = l[:mp.start()].strip(" |=—.*'\"")
                else:
                    mp2 = re.search(r"\b([A-Z]{3}[0-9][A-Z0-9][0-9]{2})\b", txt)
                    if mp2:
                        placa = _norm_placa(mp2.group(1))
                    descricao = l.strip(" |=—.*'\"")
                if descricao:
                    break
        if descricao:
            break

    # ── Local da infração e valores monetários ─────────────────
    local = vo = vp = ""
    log_rx = r"(AVENIDA|AVENI|\bAV\b|AV\s|RUA\s|ESTRADA|MARACANA|ATLANTICA|PCBAN|BANDEIRA|RODOVIA|PRACA)"
    for txt in [tD, tD2, tA, tB, tC]:
        for l in txt.splitlines():
            if re.search(log_rx, l, re.I) and len(l.strip()) > 10:
                vals = re.findall(r"\d{2,6}[,.]?\d{2}", l)
                if len(vals) >= 2:
                    vo, vp = _norm_valor(vals[0]), _norm_valor(vals[1])
                elif len(vals) == 1:
                    vo = vp = _norm_valor(vals[0])
                local = re.sub(r'[\d,.\s|"\']+$', "", l).strip(" |iI\"'*.—")
                local = re.sub(r"\s{2,}", " ", local).strip()
                if local:
                    break
        if local:
            break

    # Fallback: busca valores em todo o bloco se ainda vazios
    if not vo:
        todos = tA + tB + tC + tD
        vals  = re.findall(r"\b(\d{1,4}[,.]\d{2})\b", todos)
        if len(vals) >= 2:
            vo, vp = _norm_valor(vals[0]), _norm_valor(vals[1])
        elif len(vals) == 1:
            vo = vp = _norm_valor(vals[0])

    # ── Status de pagamento ────────────────────────────────────
    status = ""
    for txt in [tA, tB, tC, tD, tE]:
        for l in txt.splitlines():
            if re.search(r"NÃO PAGO|NAOPAGO|N.O PAGO|INMOPAGO|NÃOPAGO", l, re.I):
                status = "NÃO PAGO"
                break
            elif re.search(r"\bPAGO\b", l, re.I):
                status = "PAGO"
                break
        if status:
            break

    return {
        "auto_infracao"       : auto_inf,
        "auto_renainf"        : auto_ren,
        "data_pagamento_desc" : data_pgto,
        "data_infracao"       : data_inf,
        "hora_infracao"       : hora,
        "placa_relacionada"   : placa,
        "descricao"           : " ".join(descricao.split()),
        "local_infracao"      : local,
        "valor_original_rs"   : vo,
        "valor_a_pagar_rs"    : vp,
        "status_pagamento"    : status,
    }


# ── Função pública principal ───────────────────────────────────

def extrair(caminho_pdf: str) -> dict:
    """
    Processa um PDF do tipo 'Nada Consta do Veículo (RENAVAM)'.
    Retorna dicionário com todos os campos extraídos.
    """
    t0  = time.time()
    doc = fitz.open(caminho_pdf)
    pix = doc[0].get_pixmap(
        matrix=fitz.Matrix(DPI / 72, DPI / 72),
        colorspace=fitz.csRGB
    )
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    doc.close()
    H = img.height

    # Cabeçalho do documento
    t_cab = _ocr(img, 200, 800, psm=4, contraste=2.5)
    m_dc  = re.search(r"Data da consulta[:\s]+(\d{2}[/\s]\d{2}[/\s]\d{4}[\s\d:]*)", t_cab, re.I)
    m_rv  = re.search(r"Renavam[:\s|O]+(\d[\d\s]{8,20})", t_cab, re.I)
    data_consulta = re.sub(r"\s+", " ", m_dc.group(1)).strip() if m_dc else ""
    renavam       = re.sub(r"\D", "", m_rv.group(1))[:11]       if m_rv else ""

    # Detecta e processa cada bloco de infração
    blocos    = _detectar_blocos(img)
    infracoes = []
    for idx, (yc1, yc2) in enumerate(blocos):
        y_fim = blocos[idx + 1][0] if idx + 1 < len(blocos) else min(yc2 + 2000, H)
        infracoes.append(_extrair_bloco(img, yc2, y_fim))

    return {
        "tipo_documento" : "RENAVAM - Nada Consta do Veículo",
        "arquivo"        : os.path.basename(caminho_pdf),
        "tempo_segundos" : round(time.time() - t0, 2),
        "data_consulta"  : data_consulta,
        "renavam"        : renavam,
        "total_infracoes": len(infracoes),
        "infracoes"      : infracoes,
    }
