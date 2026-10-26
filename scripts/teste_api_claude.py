#!/usr/bin/env python3
"""
Teste comparativo: Parser regex vs API Claude
Extrai campos de páginas do PDF e compara os resultados.
"""

import json, time, re
import fitz
import urllib.request

PDF_PATH = r"C:\dev\clients\extract-pdf-ocr\entrada\CEMAVE_2025_10.pdf"
API_URL = "https://api.anthropic.com/v1/messages"
MODEL = "claude-sonnet-4-6"
API_KEY = "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"  # ← substitua pela sua chave em console.anthropic.com

PAGINAS_TESTE = [5, 11, 51, 61, 65, 80, 81, 89, 101, 105, 15, 25, 36, 130]

PROMPT_SISTEMA = """Você é um extrator de dados de documentos contábeis brasileiros.
Dado o texto de uma página de documento, extraia exatamente os 5 campos abaixo e retorne SOMENTE um JSON válido, sem explicações.

Campos:
- tipo: tipo do documento (ex: BOLETO, NFSe, DAS, DARF, HOLERITE, RCI, FATURA_TELEFONE, ENERGIA, COMPROVANTE_PIX, COMPROVANTE_CARTAO, CHEQUE, EXTRATO, RELATORIO, DESCONHECIDO)
- nome: nome do beneficiário, fornecedor ou empresa emissora
- data: data de vencimento ou emissão no formato DD/MM/YYYY
- doc: número do documento (nota fiscal, boleto, etc.)
- comp: descrição resumida do que é o pagamento (competência, serviço, produto)
- valor: valor monetário principal no formato brasileiro (ex: 1.234,56)

Regras:
- Se um campo não existir no documento, retorne string vazia ""
- Retorne SOMENTE o JSON, sem markdown, sem explicações
- Exemplo: {"tipo":"BOLETO","nome":"SANEAGO","data":"10/11/2025","doc":"123456","comp":"ÁGUA NOV/2025","valor":"187,43"}"""


# ── Detector de tipo do parser regex (simplificado) ───────────
def detectar_tipo_regex(texto):
    t = texto.upper()
    if re.search(r"GFD\s*-\s*GUIA DO FGTS DIGITAL|FGTS\s+DIGITAL", t):
        return "FGTS"
    if re.search(
        r"NIO.*CLIENTE.*FATURA|NIO\s*[-–]\s*CLIENT|SERVICOS DE REDE NORDESTE",
        t,
        re.DOTALL,
    ):
        return "FATURA_TELEFONE"
    if re.search(r"vivo\s*\*|Telefonica Brasil", t, re.I):
        return "FATURA_TELEFONE"
    if re.search(r"DOCUMENTO DE ARRECADA[ÇC][AÃ]O DO SIMPLES NACIONAL", t):
        if not re.search(r"SANEAGO|EQUATORIAL|NFS-E|BOLETO|FICHA DE COMPENSA", t):
            return "DAS"
    if re.search(r"DOCUMENTO DE ARRECADA[ÇC][AÃ]O DE RECEITAS FEDERAIS", t):
        return "DARF"
    if re.search(r"RCI\s*-\s*REMUNERA[ÇC][AÃ]O DE CONTRIBUINTE INDIVIDUAL", t):
        return "RCI"
    if re.search(
        r"DEMONSTRATIVO DE PAGAMENTO|HOLERITE|SALARIO FIXO.*INSS", t, re.DOTALL
    ):
        return "HOLERITE"
    if re.search(r"SANEAGO|FATURA DE [ÁA]GUA", t):
        return "AGUA_SANEAGO"
    if re.search(r"EQUATORIAL\s+GOI[ÁA]S|ENERGIA EL[ÉE]TRICA.*DISTRIBUIDORA", t):
        return "ENERGIA_EQUATORIAL"
    if re.search(r"DOCUMENTO AUXILIAR DA NFS-E|DANFSE|NFS-E\b", t):
        return "NFSe"
    if re.search(
        r"BOLETO|FICHA DE COMPENSA[ÇC][AÃ]O|NOSSO N[ÚU]MERO|PAGÁVEL EM QUALQUER BANCO",
        t,
    ):
        return "BOLETO"
    return "DESCONHECIDO"


# ── Chamada à API ─────────────────────────────────────────────
def chamar_api(texto_pagina):
    payload = json.dumps(
        {
            "model": MODEL,
            "max_tokens": 1000,
            "system": PROMPT_SISTEMA,
            "messages": [{"role": "user", "content": texto_pagina[:3000]}],
        }
    ).encode("utf-8")

    req = urllib.request.Request(
        API_URL,
        data=payload,
        headers={
            "Content-Type": "application/json",
            "x-api-key": API_KEY,
            "anthropic-version": "2023-06-01",
        },
        method="POST",
    )

    with urllib.request.urlopen(req) as resp:
        data = json.loads(resp.read())

    resposta_texto = data["content"][0]["text"].strip()
    # Remove possível markdown residual
    resposta_texto = re.sub(r"^```json\s*|\s*```$", "", resposta_texto.strip())
    return json.loads(resposta_texto)


# ── Pipeline principal ────────────────────────────────────────
def main():
    doc = fitz.open(PDF_PATH)
    resultados = []

    print(f"\n{'='*70}")
    print(f"  TESTE COMPARATIVO: REGEX vs API CLAUDE")
    print(f"  PDF: CEMAVE_2025_10.pdf  |  Páginas: {len(PAGINAS_TESTE)}")
    print(f"{'='*70}\n")

    for pg_num in PAGINAS_TESTE:
        texto = doc[pg_num - 1].get_text()
        tipo_regex = detectar_tipo_regex(texto) if texto.strip() else "SEM_TEXTO"

        print(f"Pág {pg_num:3d} | Regex: {tipo_regex:<20} ", end="", flush=True)

        t0 = time.time()
        try:
            resultado_api = chamar_api(texto)
            tempo = round(time.time() - t0, 1)
            print(f"| API: {resultado_api.get('tipo','?'):<20} | {tempo}s")
        except Exception as e:
            resultado_api = {"erro": str(e)}
            tempo = round(time.time() - t0, 1)
            print(f"| ERRO: {str(e)[:40]} | {tempo}s")

        resultados.append(
            {
                "pagina": pg_num,
                "tipo_regex": tipo_regex,
                "texto_amostra": texto[:200].replace("\n", " "),
                "api": resultado_api,
            }
        )

        time.sleep(0.3)  # evita rate limit

    doc.close()

    # ── Relatório detalhado ───────────────────────────────────
    print(f"\n{'='*70}")
    print(f"  RESULTADOS DETALHADOS")
    print(f"{'='*70}")

    for r in resultados:
        print(f"\n{'─'*70}")
        print(f"  PÁGINA {r['pagina']}  |  Regex: {r['tipo_regex']}")
        print(f"  Texto: {r['texto_amostra'][:120]}...")
        if "erro" in r["api"]:
            print(f"  API ERRO: {r['api']['erro']}")
        else:
            api = r["api"]
            print(f"  ┌─ Tipo  : {api.get('tipo','')}")
            print(f"  ├─ Nome  : {api.get('nome','')}")
            print(f"  ├─ Data  : {api.get('data','')}")
            print(f"  ├─ Doc   : {api.get('doc','')}")
            print(f"  ├─ Comp  : {api.get('comp','')}")
            print(f"  └─ Valor : {api.get('valor','')}")

    # ── Salva JSON completo ───────────────────────────────────
    saida = r"C:\dev\clients\extract-pdf-ocr\scripts\resultado_teste_api.json"
    with open(saida, "w", encoding="utf-8") as f:
        json.dump(resultados, f, ensure_ascii=False, indent=2)
    print(f"\n📋 JSON completo salvo em: {saida}")


if __name__ == "__main__":
    main()
