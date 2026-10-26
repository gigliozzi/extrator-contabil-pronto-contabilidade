#!/usr/bin/env python3
"""
===============================================
Script : parser_multi_doc.py
Descrição : Extrai dados de PDFs com texto nativo
            contendo múltiplos tipos de documento
            e gera planilha Excel no formato do cliente
Versão : v1.0 - 2026-05-06
Colunas Excel: Nome | DATA | DOC | COMP | VALOR
===============================================
"""

import os, re, sys, json, time
from datetime import datetime
from pathlib import Path

import fitz  # PyMuPDF
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Caminhos ──────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent.parent
DIR_ENTRADA = BASE_DIR / "entrada"
DIR_PROCESSADOS = BASE_DIR / "processados"
DIR_RELATORIOS = BASE_DIR / "relatorios"
DIR_RELATORIOS.mkdir(exist_ok=True)

# ── Estilos Excel ─────────────────────────────────────────────
COR_HEADER = "1F4E79"
COR_PAR = "DEEAF1"
BORDA = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

# ══════════════════════════════════════════════════════════════
# UTILITÁRIOS
# ══════════════════════════════════════════════════════════════


def limpar(texto):
    """Remove espaços múltiplos e quebras desnecessárias."""
    return re.sub(r"\s+", " ", texto).strip()


def buscar(padrao, texto, grupo=1, flags=re.IGNORECASE):
    m = re.search(padrao, texto, flags)
    return limpar(m.group(grupo)) if m else ""


def norm_data(s):
    """Normaliza data para DD/MM/YYYY."""
    for pat in [r"(\d{2}/\d{2}/\d{4})", r"(\d{2}\.\d{2}\.\d{4})"]:
        m = re.search(pat, s)
        if m:
            try:
                return datetime.strptime(
                    m.group(1).replace(".", "/"), "%d/%m/%Y"
                ).strftime("%d/%m/%Y")
            except:
                pass
    return ""


def primeira_data(texto):
    """Retorna a primeira data encontrada no texto."""
    m = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", texto)
    return m.group(1) if m else ""


def norm_valor(texto):
    """
    Extrai o maior valor monetário do texto.
    Prefere formato N.NNN,NN sobre NNN,NN.
    """
    # Tenta padrão com milhar: 1.234,56
    vals_milhar = re.findall(r"\b\d{1,3}(?:\.\d{3})+,\d{2}\b", texto)
    if vals_milhar:
        # Retorna o maior
        def to_float(v):
            return float(v.replace(".", "").replace(",", "."))

        return max(vals_milhar, key=to_float)
    # Fallback: NNN,NN
    vals = re.findall(r"\b\d{2,6},\d{2}\b", texto)
    if vals:

        def to_float(v):
            return float(v.replace(",", "."))

        return max(vals, key=to_float)
    return ""


def linhas(texto):
    return [l.strip() for l in texto.splitlines() if l.strip()]


# ══════════════════════════════════════════════════════════════
# DETECTORES DE TIPO
# ══════════════════════════════════════════════════════════════


def detectar_tipo(texto):
    t = texto.upper()
    if re.search(r"GFD\s*-\s*GUIA DO FGTS DIGITAL|FGTS\s+DIGITAL", t):
        return "FGTS"
    if re.search(
        r"NIO.*CLIENTE.*FATURA|TOTAL DA SUA FATURA.*NIO|NIO.*INTERNET.*EMPRESAS|NIO\s*[-–]\s*CLIENT|SERVICOS DE REDE NORDESTE",
        t,
        re.DOTALL,
    ):
        return "FATURA_TELEFONE"
    if re.search(
        r"vivo\s*\*|Telefonica Brasil|TOTAL A PAGAR.*vivo|vivo.*TOTAL A PAGAR",
        t,
        re.DOTALL | re.I,
    ):
        return "FATURA_TELEFONE"
    # DAS: exige que seja o título principal, não apenas menção no rodapé
    if re.search(r"DOCUMENTO DE ARRECADA[ÇC][AÃ]O DO SIMPLES NACIONAL", t):
        # Confirma que não é um documento que apenas cita o Simples Nacional
        if not re.search(
            r"SANEAGO|EQUATORIAL|FATURA DE [ÁA]GUA|NFS-E|DANFSE|BOLETO|FICHA DE COMPENSA",
            t,
        ):
            return "DAS"
    if re.search(r"DOCUMENTO DE ARRECADA[ÇC][AÃ]O DE RECEITAS FEDERAIS|SENDA", t):
        return "DARF"
    if re.search(
        r"CONTRIBUI[ÇC][AÃ]O NEGOCIAL|SINDICATO.*AUTÔNOMOS|SEACOM.*CONTRIBUI|CONTRIBUI.*SEACOM",
        t,
    ):
        return "SINDICATO"
    if re.search(r"D\.U\.A\.M\.|DOCUMENTO UNICO DE ARRECADA", t):
        return "ISS_DUAM"
    if re.search(r"RCI\s*-\s*REMUNERA[ÇC][AÃ]O DE CONTRIBUINTE INDIVIDUAL", t):
        return "RCI"
    if re.search(
        r"DEMONSTRATIVO DE PAGAMENTO|HOLERITE|SALARIO FIXO.*INSS", t, re.DOTALL
    ):
        return "HOLERITE"
    if re.search(r"RECIBO DE F[ÉE]RIAS|AVISO\s+PR[ÉE]VIO\s+DE\s+F[ÉE]RIAS", t):
        return "FERIAS"
    if re.search(r"AVISO PR[ÉE]VIO DO EMPREGADOR.*?TRABALHADO", t, re.DOTALL):
        return "AVISO_PREVIO"
    if re.search(r"TERMO DE RESCIS[AÃ]O|TERMO DE QUITA[ÇC][AÃ]O DE RESCIS[AÃ]O", t):
        return "RESCISAO"
    if re.search(r"AVISO PR[ÉE]VIO\s+DO EMPREGADOR", t):
        return "AVISO_PREVIO"
    if re.search(r"SANEAGO|FATURA DE [ÁA]GUA", t):
        return "AGUA_SANEAGO"
    if re.search(r"EQUATORIAL\s+GOI[ÁA]S|ENERGIA EL[ÉE]TRICA.*DISTRIBUIDORA", t):
        return "ENERGIA_EQUATORIAL"
    if re.search(r"DOCUMENTO AUXILIAR DA NFS-E|DANFSE|NFS-E\b", t):
        return "NFSe"
    if re.search(
        r"NOTA FISCAL DE SERVI[ÇC]OS ELETR[ÔO]NICA|NFS-E CAMPINAS|NFS-E\s+CURITIBA|NFS-E\s+GOIANIA",
        t,
    ):
        return "NFSe"
    if re.search(r"DOCUMENTO AUXILIAR DA NOTA FISCAL DE CONSUMIDOR|NFCe\s+N[°º]", t):
        return "NFCe"
    if re.search(
        r"BOLETO|FICHA DE COMPENSA[ÇC][AÃ]O|NOSSO N[ÚU]MERO|PAGÁVEL EM QUALQUER BANCO",
        t,
    ):
        return "BOLETO"
    return "DESCONHECIDO"


# ══════════════════════════════════════════════════════════════
# PARSERS POR TIPO
# Retornam dict: {nome, data, doc, comp, valor}
# ══════════════════════════════════════════════════════════════


def _vazio():
    return {"nome": "", "data": "", "doc": "", "comp": "", "valor": ""}


# ── BOLETO ────────────────────────────────────────────────────
def parse_boleto(texto):
    r = _vazio()
    # Beneficiário: busca nome após a linha "Beneficiário" ignorando linhas com CNPJ/label
    # Inter: "Quem vai receber:\nNOME" — aparece antes de Beneficiário
    m_quem = re.search(r"Quem vai receber:\s*\n([^\n]+)", texto, re.I)
    if m_quem:
        r["nome"] = limpar(m_quem.group(1))

    linhas_txt = [l.strip() for l in texto.splitlines() if l.strip()]
    if not r["nome"]:
        for i, l in enumerate(linhas_txt):
            if re.match(r"Benefici[áa]rio\s*$", l, re.I):
                # Pega próxima linha que não seja endereço, CNPJ ou label
                for j in range(i + 1, min(i + 6, len(linhas_txt))):
                    candidato = linhas_txt[j]
                    # Ignora linhas com CNPJ, endereço, labels, números ou logradouros
                    if re.match(
                        r"^(\d{2}\.\d{3}|\d{2}/\d{2}|Endere|Paga|Local|Venci|Ag[êe]n|Data|RUA|AV\.|AVEN|CEP|GO\s|Goiânia)",
                        candidato,
                        re.I,
                    ):
                        continue
                    # Remove CNPJ inline se houver
                    nome = re.sub(
                        r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\s*[-–\-]\s*", "", candidato
                    ).strip()
                    # Remove sufixos como "Efi Pagamentos - CNPJ"
                    nome = re.sub(r"\s*[|\-–]\s*\d{2}\.\d{3}.*$", "", nome).strip()
                    nome = re.sub(
                        r"\s+(?:Efi Pagamentos|Agência|Ag[êe]ncia).*$",
                        "",
                        nome,
                        flags=re.I,
                    ).strip()
                    if len(nome) > 4:
                        r["nome"] = nome
                        break
                break

    # Fallback: busca linha com padrão de nome de pessoa/empresa
    if not r["nome"]:
        m = re.search(
            r"Benefici[áa]rio\s+(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\s*[-\u2013]\s*)?([A-Z\xc1\xc9\xcd\xd3\xda\xc2\xca\xd4\xc3\xd5\xc7][^\n]{5,60})",
            texto,
            re.I,
        )
        if m:
            r["nome"] = limpar(m.group(2))

    # Vencimento
    m_v = re.search(r"Vencimento\s*\n?\s*(\d{2}/\d{2}/\d{4})", texto, re.I)
    r["data"] = m_v.group(1) if m_v else ""

    # Número do documento — prioriza formatos longos (ex: 00019/112/9054657304-2)
    # antes de cair no padrão genérico
    m_d = re.search(
        r"N[°º\.]\s*(?:do\s+)?[Dd]ocumento\s*\n?\s*(\d{5}/\d{3}/\d+[-\d]*)", texto, re.I
    )
    if not m_d:
        m_d = re.search(
            r"N[°º\.]\s*(?:do\s+)?[Dd]ocumento\s*\n?\s*([A-Z0-9/\-\.]{4,30})",
            texto,
            re.I,
        )
    if m_d:
        val_doc = limpar(m_d.group(1))
        # Rejeita se capturou uma data no formato MM/YYYY ou DD/MM/YYYY
        if not re.match(r"^\d{2}/\d{4}$", val_doc) and not re.match(
            r"^\d{2}/\d{2}/\d{4}$", val_doc
        ):
            r["doc"] = val_doc

    # Demonstrativo / descrição
    m_dem = re.search(r"Demonstrativo[:\s]+([^\n]{3,60})", texto, re.I)
    if m_dem:
        comp = limpar(m_dem.group(1))
        # Ignora se capturou apenas número de NF
        if not re.match(r"^NFE:\s*\d+$", comp, re.I):
            r["comp"] = comp

    # Valor do documento
    m_val = re.search(r"Valor\s+do\s+[Dd]ocumento\s*\n?\s*([\d\.,]+)", texto, re.I)
    if m_val:
        v = limpar(m_val.group(1))
        # Corrige valores sem vírgula: "3.29385" → "3.293,85"
        if "," not in v:
            v = v.replace(".", "") if v.count(".") > 1 else v
            if re.match(r"^\d+$", v.replace(".", "")) and len(v.replace(".", "")) >= 3:
                raw = v.replace(".", "")
                v = raw[:-2] + "," + raw[-2:]
        r["valor"] = v
    else:
        r["valor"] = norm_valor(texto)

    return r


# ── NFS-e ─────────────────────────────────────────────────────
def parse_nfse(texto):
    r = _vazio()
    # Nome do prestador: primeira ocorrência de Nome/Nome Empresarial com valor real
    for pat in [
        r"Nome[!/]?\s*/?\s*Nome Empresarial\s*\n([A-ZÁÉÍÓÚÂÊÔÃÕÇ][^\n]{3,80})",
        r"Nome\s*/\s*Nome\s+Empresarial\s+([A-Z][^\n]{3,60})",
        r"EMITENTE.*?Nome[!/]?\s*/?\s*Nome Empresarial\s*\n?\s*([A-Z][^\n]{3,60})",
        r"Raz[ãa]o Social.*?Prestador.*?\n([A-Z][^\n]{3,60})",
    ]:
        m = re.search(pat, texto, re.I | re.DOTALL)
        if m:
            nome = limpar(m.group(1))
            # Descarta se capturou label, e-mail, CPF/CNPJ inline ou lixo OCR
            if not re.match(
                r"^(E-mail|Email|CNPJ|CPF|Inscri|Telef|Endere|Munic|urna|uma|\d{2,3}[\.\d])",
                nome,
                re.I,
            ):
                r["nome"] = nome
                break

    # Se nome ainda vazio, tenta "Dados do Prestador de Serviço" (NFS-e municipal)
    if not r["nome"]:
        m_prest = re.search(
            r"Dados do Prestador.*?\n([A-Z][^\n]{4,60})", texto, re.I | re.DOTALL
        )
        if m_prest:
            cand = limpar(m_prest.group(1))
            if not re.match(r"^(CNPJ|CPF|Inscri|Endere|CEP|Fone|Rua|Av\.)", cand, re.I):
                r["nome"] = cand
    # Fallback: razão social do prestador em NFS-e Goiânia
    if not r["nome"]:
        m_rs = re.search(r"Nome Empresarial\s*\n([A-Z][^\n]{4,80})", texto, re.I)
        if m_rs:
            r["nome"] = limpar(m_rs.group(1))

    # Fallback adicional: nome de pessoa física com CPF inline (ex: "RAQUEL ... 71399088181")
    if not r["nome"]:
        m_pf = re.search(
            r"([A-ZÁÉÍÓÚÂÊÔÃÕÇ][A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]{8,60})\s+(\d{11})\b", texto, re.I
        )
        if m_pf:
            r["nome"] = limpar(m_pf.group(1))

    # Data de emissão
    m_d = re.search(
        r"Data e Hora da emiss[ãa]o da NFS-e\s*\n?\s*(\d{2}/\d{2}/\d{4})", texto, re.I
    )
    if not m_d:
        m_d = re.search(
            r"Data\s*(?:e\s*[Hh]ora)?\s*de\s*[Gg]era[çc][ãa]o\s*(?:da\s*NFS-e)?\s*\n?\s*(\d{2}/\d{2}/\d{4})",
            texto,
            re.I,
        )
    r["data"] = m_d.group(1) if m_d else primeira_data(texto)

    # Número da NFS-e
    m_n = re.search(r"N[úu]mero da NFS-e\s*\n?\s*(\d+)", texto, re.I)
    r["doc"] = m_n.group(1) if m_n else ""

    # Descrição do serviço
    m_desc = re.search(r"Descri[çc][aã]o do Servi[çc]o\s*\n([^\n]{5,120})", texto, re.I)
    if m_desc:
        r["comp"] = limpar(m_desc.group(1))

    # Valor do serviço
    m_val = re.search(r"Valor do Servi[çc]o\s*\n?\s*R\$?\s*([\d\.,]+)", texto, re.I)
    if m_val:
        r["valor"] = limpar(m_val.group(1))
    else:
        r["valor"] = norm_valor(texto)

    return r


# ── NFCe ──────────────────────────────────────────────────────
def parse_nfce(texto):
    r = _vazio()
    # NFCe: "CNPJ: XX/XX - NOME -" ou terceira linha com razão social
    m = re.search(r"CNPJ:\s*[\d./\-]+\s*[-–]\s*([^\n]{5,80}?)\s*[-–\n]", texto, re.I)
    if m:
        r["nome"] = limpar(m.group(1)).rstrip("-– ")
    if not r["nome"]:
        # Padrão CNPJ partido em duas linhas (pág 88: "CNPJ: 14.524\n5/0001-78 - NOME")
        m2 = re.search(
            r"CNPJ:\s*[\d./\-]+\n[\d./\-]+\s*[-–]\s*([^\n]{5,80}?)\s*[-–\n]",
            texto,
            re.I,
        )
        if m2:
            r["nome"] = limpar(m2.group(1)).rstrip("-– ")
    if not r["nome"]:
        linhas_nfce = [l.strip() for l in texto.splitlines() if l.strip()]
        for l in linhas_nfce[:6]:
            if re.match(
                r"^[A-ZÁÉÍÓÚÂÊÔÃÕÇ][A-Za-záéíóúâêôãõçÁ\s]+(?:LTDA|ME|SA|S\.A\.|Gas|Gás|Papelaria|BRASIL)",
                l,
                re.I,
            ):
                r["nome"] = l
                break

    # Data e número: "NFCe nº 28963 Série 1 22/12/2025"
    m_nfce = re.search(
        r"NFCe\s+n[°º]\s*(\d+)\s+S[ée]rie\s+\d+\s+(\d{2}/\d{2}/\d{4})", texto, re.I
    )
    if m_nfce:
        r["doc"] = m_nfce.group(1)
        r["data"] = m_nfce.group(2)
    else:
        m_n = re.search(r"NFCe\s+n[°º]\s*(\d+)", texto, re.I)
        r["doc"] = m_n.group(1) if m_n else ""
        r["data"] = primeira_data(texto)

    # Descrição dos itens (primeira descrição de produto)
    m_desc = re.search(r"Descri[çc][aã]o\s*\n([^\n]{5,80})", texto, re.I)
    if m_desc:
        r["comp"] = limpar(m_desc.group(1))

    # Valor a pagar
    m_val = re.search(r"Valor a Pagar R?\$?\s*([\d\.,]+)", texto, re.I)
    if m_val:
        r["valor"] = limpar(m_val.group(1))
    else:
        r["valor"] = norm_valor(texto)

    return r


# ── DARF ──────────────────────────────────────────────────────
def parse_darf(texto):
    r = _vazio()
    m = re.search(r"Raz[ãa]o Social\s*\n([A-ZÁÉÍÓÚÂÊÔÃÕÇ][^\n]{3,60})", texto, re.I)
    if m:
        nome = limpar(m.group(1))
        # Descarta se capturou label em vez de nome
        r["nome"] = (
            nome
            if not re.match(r"^(CNPJ|Período|Data|Número|Observ)", nome, re.I)
            else "RECEITA FEDERAL"
        )
    else:
        r["nome"] = "RECEITA FEDERAL"

    # "DatadeVencimento" sem espaço é padrão frequente no DARF
    m_d = re.search(
        r"Data\s*de\s*Venc(?:imento)?\s*\n?\s*(\d{2}/\d{2}/\d{4})", texto, re.I
    )
    if not m_d:
        m_d = re.search(r"Pagar.*?at[ée]\s*(\d{2}/\d{2}/\d{4})", texto, re.I)
    if not m_d:
        m_d = re.search(r"\b(\d{2}/\d{2}/20\d{2})\b", texto)  # qualquer data
    r["data"] = m_d.group(1) if m_d else ""

    m_n = re.search(r"N[úu]mero\s*do\s*[Dd]ocumento\s*\n([^\n]+)", texto, re.I)
    r["doc"] = limpar(m_n.group(1)) if m_n else ""

    # Denominações (INSS, IRRF, etc.)
    denoms = re.findall(
        r"\b(IRRF|INSS|CSLL|COFINS|PIS|CP SEGURADOS|CONTRIBUI[ÇC][AÃ]O PREV)\b",
        texto,
        re.I,
    )
    r["comp"] = (
        " / ".join(dict.fromkeys(d.upper() for d in denoms)) if denoms else "DARF"
    )

    m_val = re.search(r"Valor Total do Documento\s*\n?\s*([\d\.,]+)", texto, re.I)
    if m_val:
        r["valor"] = limpar(m_val.group(1))
    else:
        r["valor"] = norm_valor(texto)

    return r


# ── FGTS ──────────────────────────────────────────────────────
def parse_fgts(texto):
    r = _vazio()
    # Nome: busca linha após "Nome/Razão Social do Empregador", ignora "CPF/CNPJ"
    linhas_fgts = [l.strip() for l in texto.splitlines() if l.strip()]
    for idx_f, l_f in enumerate(linhas_fgts):
        if re.search(r"Nome/Raz[ãa]o Social do Empregador", l_f, re.I):
            for jj in range(idx_f + 1, min(idx_f + 4, len(linhas_fgts))):
                cand = linhas_fgts[jj]
                if re.match(r"^(CPF|CNPJ|\d|Núm|Ident)", cand, re.I):
                    continue
                if len(cand) > 4:
                    r["nome"] = cand
                    break
            break
    if not r["nome"]:
        m2 = re.search(r"([A-Z][A-Z\s]+(?:LTDA|ME|SA)\b)", texto, re.I)
        r["nome"] = limpar(m2.group(1)) if m2 else "FGTS DIGITAL"

    m_d = re.search(
        r"Pagar.*?documento.*?at[ée]\s*\n?\s*(\d{2}/\d{2}/\d{4})",
        texto,
        re.I | re.DOTALL,
    )
    r["data"] = m_d.group(1) if m_d else primeira_data(texto)

    m_id = re.search(r"Identificador\s*\n([^\n]+)", texto, re.I)
    r["doc"] = ""  # deixar em branco conforme solicitado

    m_tag = re.search(r"Tag\s*\n([^\n]+)", texto, re.I)
    r["comp"] = limpar(m_tag.group(1)) if m_tag else "FGTS"

    m_val = re.search(r"Valor a recolher\s*\n?\s*([\d\.,]+)", texto, re.I)
    if m_val:
        r["valor"] = limpar(m_val.group(1))
    else:
        r["valor"] = norm_valor(texto)

    return r


# ── DAS ───────────────────────────────────────────────────────
def parse_das(texto):
    r = _vazio()
    # Razão Social: padrão "Razão Social\nNOME" ou inline
    m = re.search(r"Raz[ãa]o\s*So[c]?[ia]+l[\s\n\\]+([A-Z][^\n]{3,60})", texto, re.I)
    if not m:
        # Padrão "CNPJ\nXXXX\nRazão Social\nNOME" frequente no DAS
        partes = re.split(r"Raz[ãa]o\s*So[c]?[ia]+l", texto, flags=re.I)
        if len(partes) > 1:
            proxima = [l.strip() for l in partes[1].splitlines() if l.strip()]
            for p in proxima[:3]:
                if re.match(r"^[A-Z][A-Z\s]{3,}", p) and not re.match(
                    r"^(CNPJ|Data|Período|Número|Obs)", p, re.I
                ):
                    r["nome"] = p
                    break
    if not r["nome"]:
        if m:
            r["nome"] = limpar(m.group(1))
        else:
            r["nome"] = (
                buscar(r"([A-Z][A-Z\s]+(?:LTDA|ME|SA|S\.A\.))", texto)
                or "SIMPLES NACIONAL"
            )

    m_d = re.search(
        r"Data\s*de\s*Venc(?:imento)?\s*\n?\s*(\d{2}/\d{2}/\d{4})", texto, re.I
    )
    r["data"] = m_d.group(1) if m_d else ""

    m_n = re.search(r"N[úu]mero\s*do\s*Documento\s*\n([^\n]+)", texto, re.I)
    r["doc"] = limpar(m_n.group(1)) if m_n else ""

    m_per = re.search(r"Per[íi]odo\s*de\s*Apura[çc][ãa]o\s*\n([^\n]+)", texto, re.I)
    r["comp"] = f"DAS {limpar(m_per.group(1))}" if m_per else "SIMPLES NACIONAL"

    m_val = re.search(r"Valor Total do Documento\s*\n?\s*([\d\.,]+)", texto, re.I)
    if m_val:
        v = limpar(m_val.group(1))
        # Garante formato correto: "1.18255" → "1.182,55"
        if "," not in v and len(v) >= 3:
            v = v[:-2] + "," + v[-2:]
        r["valor"] = v
    else:
        r["valor"] = norm_valor(texto)

    return r


# ── HOLERITE ──────────────────────────────────────────────────
def parse_holerite(texto):
    r = _vazio()
    # Empresa
    m = re.search(
        r"^([A-Z][A-Z\s]+LTDA|[A-Z][A-Z\s]+ME|[A-Z][A-Z\s]+S\.?A\.?)\b", texto, re.M
    )
    r["nome"] = limpar(m.group(1)) if m else "PRONTO CONTABILIDADE"

    # Data do recibo
    m_d = re.search(r"Data:\s*(\d{2}[\s/]\d{2}[\s/]\d{2,4})", texto, re.I)
    if m_d:
        raw = re.sub(r"\s+", "/", m_d.group(1).strip())
        r["data"] = norm_data(raw) or raw
    else:
        # Tenta mês de referência
        meses = {
            "JANEIRO": "01",
            "FEVEREIRO": "02",
            "MARÇO": "03",
            "ABRIL": "04",
            "MAIO": "05",
            "JUNHO": "06",
            "JULHO": "07",
            "AGOSTO": "08",
            "SETEMBRO": "09",
            "OUTUBRO": "10",
            "NOVEMBRO": "11",
            "DEZEMBRO": "12",
        }
        m_mes = re.search(
            r"(JANEIRO|FEVEREIRO|MAR[ÇC]O|ABRIL|MAIO|JUNHO|JULHO|AGOSTO|SETEMBRO|OUTUBRO|NOVEMBRO|DEZEMBRO)/(\d{4})",
            texto,
            re.I,
        )
        if m_mes:
            mes = meses.get(
                m_mes.group(1).upper().replace("Ç", "C").replace("Ã", "A"), "01"
            )
            r["data"] = f"01/{mes}/{m_mes.group(2)}"

    r["doc"] = ""  # em branco

    # Nome do funcionário e mês
    m_func = re.search(r"Nome do Empregado\s*\n([^\n]+)", texto, re.I)
    m_ref = re.search(
        r"(JANEIRO|FEVEREIRO|MAR[ÇC]O|ABRIL|MAIO|JUNHO|JULHO|AGOSTO|SETEMBRO|OUTUBRO|NOVEMBRO|DEZEMBRO)/(\d{4})",
        texto,
        re.I,
    )
    func = limpar(m_func.group(1)) if m_func else ""
    ref = m_ref.group(0) if m_ref else ""
    r["comp"] = f"HOLERITE {func} {ref}".strip()

    m_liq = re.search(r"Valor\s+L[íi]quido\s*\n?\s*([\d\.,]+)", texto, re.I)
    if m_liq:
        r["valor"] = limpar(m_liq.group(1))
    else:
        r["valor"] = norm_valor(texto)

    return r


# ── RCI ───────────────────────────────────────────────────────
def parse_rci(texto):
    r = _vazio()
    m = re.search(r"Tomador[:\s]*\n([A-ZÁÉÍÓÚÂÊÔÃÕÇ][^\n]{3,60})", texto, re.I)
    if m:
        r["nome"] = limpar(m.group(1))
    else:
        m2 = re.search(
            r"Tomador.*?\n.*?\n([A-Z][A-Z\s]+(?:LTDA|ME)\b)", texto, re.I | re.DOTALL
        )
        r["nome"] = limpar(m2.group(1)) if m2 else ""

    m_ref = re.search(r"M[êe]s/Refer[êe]ncia[:\s]*\n([^\n]+)", texto, re.I)
    ref = limpar(m_ref.group(1)) if m_ref else ""

    # Data: fim do mês de referência
    meses = {
        "JANEIRO": "31/01",
        "FEVEREIRO": "28/02",
        "MARÇO": "31/03",
        "ABRIL": "30/04",
        "MAIO": "31/05",
        "JUNHO": "30/06",
        "JULHO": "31/07",
        "AGOSTO": "31/08",
        "SETEMBRO": "30/09",
        "OUTUBRO": "31/10",
        "NOVEMBRO": "30/11",
        "DEZEMBRO": "31/12",
    }
    m_mes = re.search(
        r"(JANEIRO|FEVEREIRO|MAR[ÇC]O|ABRIL|MAIO|JUNHO|JULHO|AGOSTO|SETEMBRO|OUTUBRO|NOVEMBRO|DEZEMBRO)/(\d{4})",
        ref,
        re.I,
    )
    if m_mes:
        chave = m_mes.group(1).upper().replace("Ç", "C").replace("Ã", "A")
        r["data"] = f"{meses.get(chave,'30')}/{m_mes.group(2)}"
    else:
        r["data"] = primeira_data(texto)

    r["doc"] = ""

    m_prest = re.search(r"Prestador[:\s]*\n([^\n]+)", texto, re.I)
    prest = limpar(m_prest.group(1)) if m_prest else ""
    # Limpa labels espúrios do prestador
    prest = re.sub(r"\b(Tomador|Prestador|CNPJ)[:\s]*", "", prest, flags=re.I).strip()
    r["comp"] = f"PRO-LABORE {prest} {ref}".strip()

    m_val = re.search(r"Valor\s+L[íi]quido\s*\n?\s*([\d\.,]+)", texto, re.I)
    if m_val:
        r["valor"] = limpar(m_val.group(1))
    else:
        r["valor"] = norm_valor(texto)

    return r


# ── ÁGUA SANEAGO ──────────────────────────────────────────────
def parse_agua(texto):
    r = _vazio()
    r["nome"] = "SANEAGO"

    m_v = re.search(r"VENCIMENTO[:\s]*\n?\s*(\d{2}/\d{2}/\d{4})", texto, re.I)
    r["data"] = m_v.group(1) if m_v else ""

    m_d = re.search(r"DOCUMENTO\s+N[°º][:\s]*([\d]+)", texto, re.I)
    r["doc"] = m_d.group(1) if m_d else ""

    m_ref = re.search(r"REFER[ÊEÉÈ]NCIA[:\s]*\n?\s*(\w+/\d{4})", texto, re.I)
    r["comp"] = f"ÁGUA {limpar(m_ref.group(1))}" if m_ref else "ÁGUA SANEAGO"

    m_val = re.search(r"VALOR TOTAL\s*\(R\$\)\s*\n?\s*([\d\.,]+)", texto, re.I)
    if m_val:
        r["valor"] = limpar(m_val.group(1))
    else:
        r["valor"] = norm_valor(texto)

    return r


# ── ENERGIA EQUATORIAL ────────────────────────────────────────
def parse_energia(texto):
    r = _vazio()
    r["nome"] = "EQUATORIAL GOIÁS"

    m_v = re.search(r"Vencimento\s*\n?\s*(\d{2}/\d{2}/\d{4})", texto, re.I)
    r["data"] = m_v.group(1) if m_v else ""

    m_d = re.search(r"NOTA FISCAL\s+N[°º]?\s*([\d]+)", texto, re.I)
    r["doc"] = m_d.group(1) if m_d else ""

    m_ref = re.search(r"Conta\s+m[êe]s\s*\n?\s*(\w+/\d{4})", texto, re.I)
    r["comp"] = f"ENERGIA {limpar(m_ref.group(1))}" if m_ref else "ENERGIA ELÉTRICA"

    # Energia: múltiplos padrões de valor
    m_val = re.search(
        r"(?:Total\s+a\s+pagar|Total\s+a\s+1)\s*\n?\s*R?\$?\s*(\d[\d\.]*,\d{2})\b",
        texto,
        re.I,
    )
    if not m_val:
        m_val = re.search(r"R\$?\s*\n\s*(\d[\d\.]*,\d{2})\b", texto, re.I)
    if not m_val:
        # "N5\n94.02" ou "RS\n94,02" — OCR lê R$ como N5/NS
        m_val = re.search(r"(?:N[5S]|RS)\s*\n\s*(\d+[,.]\d{2})\b", texto, re.I)
        if m_val:
            v = m_val.group(1).replace(".", ",")
            r["valor"] = v
            m_val = None  # já setou
    if m_val:
        r["valor"] = limpar(m_val.group(1))
    if not r["valor"]:
        r["valor"] = norm_valor(texto)

    return r


# ── ISS / DUAM ────────────────────────────────────────────────
def parse_iss_duam(texto):
    r = _vazio()
    r["nome"] = "PREFEITURA DE GOIÂNIA"

    m_v = re.search(r"VENCIMENTO[:\s]*\n?\s*(\d{2}/\d{2}/\d{4})", texto, re.I)
    r["data"] = m_v.group(1) if m_v else ""

    r["doc"] = ""  # em branco conforme solicitado

    m_trib = re.search(r"\d+\s*-\s*(I\.?S\.?S\.?\s*\w+)", texto, re.I)
    r["comp"] = limpar(m_trib.group(1)) if m_trib else "ISS"

    m_val = re.search(r"TOTAL\s*A\s*PAGAR[:\s]*\n?\s*[\*\s]*([\d\.,]+)", texto, re.I)
    if m_val:
        r["valor"] = limpar(m_val.group(1))
    else:
        r["valor"] = norm_valor(texto)

    return r


# ── FATURA TELEFONE (NIO, Vivo) ─────────────────────────────
def parse_fatura_telefone(texto):
    r = _vazio()
    # Nome: operadora
    if re.search(r"vivo|Telefonica Brasil", texto, re.I):
        r["nome"] = "VIVO"
    elif re.search(
        r"NIO\s*[-–]\s*CLIENT\s+CO\s+SERVICOS|SERVICOS DE REDE NORDESTE", texto, re.I
    ):
        # Captura razão social completa da NIO
        m_nio = re.search(
            r"(NIO\s*[-–]\s*CLIENT\s+CO\s+SERVICOS\s+DE\s+REDE\s+NORDESTE\s+S\.?A\.?)",
            texto,
            re.I,
        )
        r["nome"] = (
            limpar(m_nio.group(1))
            if m_nio
            else "NIO - CLIENT CO SERVICOS DE REDE NORDESTE S.A."
        )
    elif re.search(r"NIO.*INTERNET|NIO.*CLIENTE", texto, re.I):
        r["nome"] = "NIO INTERNET"
    else:
        r["nome"] = "OPERADORA"

    # Data de vencimento
    m_v = re.search(r"Vencimento\s*\n?\s*(\d{2}/\d{2}/\d{4})", texto, re.I)
    r["data"] = m_v.group(1) if m_v else primeira_data(texto)

    # Número da fatura
    m_n = re.search(r"N[ÚU]MERO DA FATURA[:\s]*([\d]+)", texto, re.I)
    r["doc"] = m_n.group(1) if m_n else ""

    # Competência
    m_c = re.search(r"(?:Mês|Per[íi]odo|eferenc[íi]a)[:\s]*(\w+/\d{4})", texto, re.I)
    r["comp"] = limpar(m_c.group(1)) if m_c else "FATURA TELEFONE"

    # Valor
    m_val = re.search(
        r"TOTAL\s+(?:DA\s+SUA\s+)?(?:A\s+)?FATURA\s*\n?\s*(?:II\s*)?(\d[\d\.,]+)",
        texto,
        re.I,
    )
    if not m_val:
        m_val = re.search(r"TOTAL\s+A\s+PAGAR\s*[-–]?\s*R\$\s*([\d\.,]+)", texto, re.I)
    r["valor"] = limpar(m_val.group(1)) if m_val else norm_valor(texto)
    return r


# ── SINDICATO / CONTRIBUIÇÃO NEGOCIAL ────────────────────────
def parse_sindicato(texto):
    r = _vazio()
    # Nome: linha com "SINDICATO" ou "Beneficiário\nNOME"
    m_sind = re.search(r"(SINDICATO[^\n]{5,120})", texto, re.I)
    if m_sind:
        r["nome"] = limpar(m_sind.group(1))[:80]
    else:
        m = re.search(r"Benefici[\u00e1a]rio\s*[:\n]+([A-Z][^\n]{5,80})", texto, re.I)
        r["nome"] = (
            re.sub(
                r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\s*[-\u2013]\s*",
                "",
                limpar(m.group(1)),
            ).strip()
            if m
            else "SINDICATO"
        )
    # Sindicato: data de vencimento da guia (não data de admissão)
    m_d = re.search(r"Vencimento\s*\n?\s*(\d{2}/\d{2}/\d{4})", texto, re.I)
    if not m_d:
        m_d = re.search(
            r"(?:Data|Goiânia|GOIANIA)[,\s]+\d+[\s/]+\d+[\s/]+(\d{4})", texto, re.I
        )
    if m_d:
        r["data"] = m_d.group(1) if "/" in m_d.group(1) else primeira_data(texto)
    else:
        # Pega primeira data que não seja de admissão (evita 02/01/2024)
        datas = re.findall(r"\b(\d{2}/\d{2}/20\d{2})\b", texto)
        r["data"] = datas[0] if datas else ""
    r["doc"] = ""
    m_comp = re.search(
        r"(CONTRIBUI[\u00c7C][\u00c3A]O\s+(?:NEGOCIAL|ASSISTENCIAL)[^\n]*)", texto, re.I
    )
    r["comp"] = limpar(m_comp.group(1)) if m_comp else "CONTRIBUI\u00c7\u00c3O SINDICAL"
    # Valor: preferência pela guia total, não salário base
    # Valor: campos específicos da guia sindical, ignora código CBO (4131.10)
    m_val = re.search(r"Valor\s+Original\s+da\s+Guia[:\s]*(\d[\d\.,]+)", texto, re.I)
    if not m_val:
        m_val = re.search(r"Total\s+a\s+Recolher[:\s]*(\d[\d\.,]+)", texto, re.I)
    if not m_val:
        # Busca valores no formato XX,XX (não 4131.10 que é CBO)
        vals = re.findall(r"\b(\d{1,4},\d{2})\b", texto)
        if vals:

            def to_f(v):
                return float(v.replace(",", "."))

            # Pega o menor valor que seja > 0 (valor por funcionário)
            positivos = [v for v in vals if to_f(v) > 0]
            r["valor"] = min(positivos, key=to_f) if positivos else ""
    if not r["valor"] and m_val:
        r["valor"] = limpar(m_val.group(1))
    return r


# ── Documentos sem valor financeiro ──────────────────────────
def parse_rh(texto, tipo):
    """Férias, rescisão, aviso prévio — sem valor financeiro direto."""
    r = _vazio()
    m = re.search(r"^([A-Z][A-Z\s]+LTDA|[A-Z][A-Z\s]+ME)\b", texto, re.M)
    r["nome"] = limpar(m.group(1)) if m else "PRONTO CONTABILIDADE"
    r["data"] = primeira_data(texto)
    r["doc"] = ""
    r["comp"] = tipo
    r["valor"] = norm_valor(texto)
    return r


# ══════════════════════════════════════════════════════════════
# ROTEADOR
# ══════════════════════════════════════════════════════════════

PARSERS = {
    "BOLETO": parse_boleto,
    "NFSe": parse_nfse,
    "NFCe": parse_nfce,
    "DARF": parse_darf,
    "FGTS": parse_fgts,
    "DAS": parse_das,
    "HOLERITE": parse_holerite,
    "RCI": parse_rci,
    "AGUA_SANEAGO": parse_agua,
    "ENERGIA_EQUATORIAL": parse_energia,
    "ISS_DUAM": parse_iss_duam,
    "SINDICATO": parse_sindicato,
    "FATURA_TELEFONE": parse_fatura_telefone,
}


def extrair_pagina(texto):
    tipo = detectar_tipo(texto)
    if tipo in PARSERS:
        dados = PARSERS[tipo](texto)
    elif tipo in ("FERIAS", "RESCISAO", "AVISO_PREVIO"):
        dados = parse_rh(texto, tipo)
    else:
        dados = _vazio()
    dados["tipo"] = tipo
    return dados


# ══════════════════════════════════════════════════════════════
# GERAÇÃO DO EXCEL
# ══════════════════════════════════════════════════════════════

COLUNAS = [
    ("Nome", 40),
    ("DATA", 14),
    ("DOC", 22),
    ("COMP", 45),
    ("VALOR", 14),
]


def gerar_excel(registros, caminho_saida):
    wb = Workbook()
    ws = wb.active
    ws.title = "Documentos"
    ws.freeze_panes = "A2"

    # Cabeçalho
    for col, (nome, larg) in enumerate(COLUNAS, 1):
        cel = ws.cell(row=1, column=col, value=nome)
        cel.font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        cel.fill = PatternFill("solid", fgColor=COR_HEADER)
        cel.alignment = Alignment(horizontal="center", vertical="center")
        cel.border = BORDA
        ws.column_dimensions[get_column_letter(col)].width = larg
    ws.row_dimensions[1].height = 20

    # Dados
    for i, reg in enumerate(registros, 2):
        par = i % 2 == 0
        valores = [
            reg.get("nome", ""),
            reg.get("data", ""),
            reg.get("doc", ""),
            reg.get("comp", ""),
            reg.get("valor", ""),
        ]
        for col, val in enumerate(valores, 1):
            cel = ws.cell(row=i, column=col, value=val)
            cel.font = Font(name="Arial", size=10)
            cel.fill = PatternFill("solid", fgColor=COR_PAR if par else "FFFFFF")
            cel.alignment = Alignment(vertical="center")
            cel.border = BORDA

    wb.save(caminho_saida)
    print(f"✅ Excel salvo: {caminho_saida}")


# ══════════════════════════════════════════════════════════════
# PIPELINE PRINCIPAL
# ══════════════════════════════════════════════════════════════


def processar_pdf(caminho_pdf):
    t0 = time.time()
    nome = Path(caminho_pdf).stem
    doc = fitz.open(caminho_pdf)

    registros = []
    tipos_encontrados = {}

    for i, pagina in enumerate(doc):
        texto = pagina.get_text()
        if not texto.strip():
            continue
        reg = extrair_pagina(texto)
        reg["pagina"] = i + 1
        registros.append(reg)
        tipos_encontrados[reg["tipo"]] = tipos_encontrados.get(reg["tipo"], 0) + 1

    doc.close()

    # Remove duplicatas de boleto (frente + verso da mesma página)
    registros = deduplicar(registros)

    tempo = round(time.time() - t0, 1)

    # Resumo
    print(f"\n{'='*60}")
    print(f"  Arquivo  : {Path(caminho_pdf).name}")
    print(f"  Páginas  : {len(registros)}")
    print(f"  Tempo    : {tempo}s")
    print(f"{'='*60}")
    for tipo, qtd in sorted(tipos_encontrados.items()):
        print(f"  {tipo:<25}: {qtd}")
    print(f"{'='*60}\n")

    return registros


def deduplicar(registros):
    """
    Remove registros duplicados: boletos têm recibo do pagador
    e ficha de compensação na mesma página — geram duplicata de tipo.
    Mantém o primeiro registro por (tipo + valor + data), descarta
    DESCONHECIDO com todos os campos vazios.
    """
    vistos = set()
    resultado = []
    for r in registros:
        # Descarta linhas completamente vazias
        if not any([r.get("nome"), r.get("data"), r.get("valor")]):
            continue
        chave = (r["tipo"], r.get("valor", ""), r.get("data", ""))
        if chave not in vistos:
            vistos.add(chave)
            resultado.append(r)
    return resultado


# ── Entrada de linha de comando ───────────────────────────────


def main():
    if len(sys.argv) < 2:
        print("Uso: python parser_multi_doc.py <arquivo.pdf> [saida.xlsx]")
        print(
            "     python parser_multi_doc.py pasta/entrada/  # processa todos os PDFs"
        )
        sys.exit(0)

    entrada = Path(sys.argv[1])

    if entrada.is_dir():
        pdfs = list(entrada.glob("*.pdf")) + list(entrada.glob("*.PDF"))
    elif entrada.is_file():
        pdfs = [entrada]
    else:
        print(f"❌ Caminho não encontrado: {entrada}")
        sys.exit(1)

    todos_registros = []
    for pdf in pdfs:
        regs = processar_pdf(str(pdf))
        todos_registros.extend(regs)

    if not todos_registros:
        print("Nenhum registro extraído.")
        sys.exit(0)

    # Saída
    if len(sys.argv) >= 3:
        saida = Path(sys.argv[2])
    else:
        data_hoje = datetime.now().strftime("%Y-%m-%d")
        saida = DIR_RELATORIOS / f"relatorio_{data_hoje}.xlsx"

    gerar_excel(todos_registros, str(saida))

    # Salva JSON para auditoria
    json_saida = saida.with_suffix(".json")
    with open(json_saida, "w", encoding="utf-8") as f:
        json.dump(todos_registros, f, ensure_ascii=False, indent=2)
    print(f"📋 JSON de auditoria: {json_saida}")


if __name__ == "__main__":
    main()
