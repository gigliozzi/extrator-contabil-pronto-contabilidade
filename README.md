# Extrator de Documentos Contábeis
### Studio Delta — William Gigliozzi

Sistema de extração automatizada de dados de documentos contábeis em PDF, utilizando visão computacional via API Claude. Converte documentos escaneados, boletos, notas fiscais, holerites e outros documentos contábeis em planilha Excel organizada.

---

## Estrutura de pastas

```
EXTRATOR-PDF\
│
├── entrada\              ← Coloque os PDFs aqui para processar
├── excel_gerado\         ← Planilhas Excel geradas automaticamente
├── logs\                 ← JSONs de auditoria por processamento
│
├── scripts\
│   ├── interface.py              ← Interface gráfica (ponto de entrada)
│   └── extrator_visao_claude.py  ← Motor de extração via API Claude
│
├── instalar_dependencias.py  ← Execute UMA VEZ após instalar o Python
├── set-env.txt               ← Instruções para configurar a chave de API
└── README.md
```

---

## Instalação (fazer apenas uma vez)

### 1. Instalar Python 3.11.8

Baixe o instalador em:
```
https://www.python.org/downloads/release/python-3118/
```
Escolha **Windows installer (64-bit)**.

> ⚠️ Na tela de instalação, marque obrigatoriamente **"Add Python to PATH"** antes de clicar em Install Now.

### 2. Instalar dependências

Abra o PowerShell na pasta do projeto e execute:
```powershell
python instalar_dependencias.py
```

Ou manualmente:
```powershell
pip install pymupdf openpyxl
```

### 3. Configurar a chave de API

Abra o PowerShell e execute o comando abaixo, substituindo pelo valor fornecido pelo suporte:
```powershell
[System.Environment]::SetEnvironmentVariable("ANTHROPIC_API_KEY", "sua-chave-aqui", "User")
```

> A chave fica salva no perfil do usuário Windows. Não é necessário repetir esse passo a cada uso.

---

## Uso diário

1. Abra o PowerShell na pasta `scripts\`
2. Execute:
```powershell
python interface.py
```
3. Na interface, clique em **Carregar PDF** e selecione o arquivo
4. Clique em **Processar PDF** e aguarde o processamento
5. Ao concluir, clique em **Abrir Excel** para visualizar o resultado

A planilha Excel é gerada automaticamente em `excel_gerado\` com o nome do arquivo original e a data/hora do processamento.

---

## Tipos de documento suportados

| Tipo | Descrição |
|------|-----------|
| BOLETO | Boletos bancários |
| NFSe / NFCe | Notas fiscais de serviço e consumidor |
| DAS / DARF | Guias de recolhimento federal |
| FGTS | Guias de recolhimento do FGTS |
| HOLERITE | Contracheques e recibos de salário |
| RCI | Recibo de pagamento autônomo |
| CHEQUE | Cheques (inclusive manuscritos) |
| COMPROVANTE_PIX | Comprovantes de transferência PIX |
| COMPROVANTE_CARTAO | Comprovantes de cartão de crédito/débito |
| FATURA_TELEFONE | Faturas de telefonia |
| ENERGIA / AGUA | Contas de energia e água |
| ISS / SINDICATO | Guias de ISS e contribuições sindicais |
| FERIAS / RESCISAO | Recibos de férias e rescisão |
| EXTRATO_CARTAO | Extratos de cartão |
| RELATORIO | Relatórios diversos |

---

## Colunas da planilha Excel gerada

| Coluna | Descrição |
|--------|-----------|
| NOME | Nome do beneficiário, fornecedor ou empresa emissora |
| DATA | Data de vencimento ou emissão (DD/MM/YYYY) |
| DOC | Número do documento |
| COMPL | Descrição resumida do pagamento e competência |
| VALOR | Valor monetário principal |

---

## Limitações conhecidas

- A precisão da extração depende da qualidade e legibilidade do PDF
- Documentos com baixa resolução, manchas ou texto muito pequeno podem apresentar erros
- O processamento exige conexão com a internet (API Claude)
- Em caso de instabilidade na API, o sistema tenta automaticamente até 3 vezes antes de registrar o erro

---

## Suporte

Em caso de problemas:
1. Verifique o log JSON em `logs\` correspondente ao processamento
2. Entre em contato informando o nome do arquivo e o conteúdo do log

**Studio Delta — William Gigliozzi**
gigliozzi@gmail.com
