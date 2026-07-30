
# ============================================================
# Importador de Borderos - Positivo DataSync
# Autor: Pedro Henrique Forsch
# Contato: pedrohenriqueforsch@gmail.com
# ============================================================

import pdfplumber
import re


def ler_pdf(caminho):

    texto = ""

    with pdfplumber.open(caminho) as pdf:
        for pagina in pdf.pages:
            conteudo = pagina.extract_text()

            if conteudo:
                texto += conteudo + "\n"

    return texto


def extrair_cabecalho(texto):

    dados = {}

    m = re.search(r"Cedente:\s*(.+)", texto)
    dados["cliente"] = m.group(1).strip() if m else ""

    m = re.search(r"N\.Border[oô]:\s*(\d+)", texto)
    dados["lote"] = m.group(1).strip() if m else ""

    m = re.search(r"Data:\s*(\d{2}/\d{2}/\d{4})", texto)
    dados["data"] = m.group(1).strip() if m else ""

    return dados


PADRAO_TITULO = re.compile(
    r"^(DM|DS|CHQ)\s+"      # tipo
    r"(\d+)\s+"             # documento (CPF/CNPJ do sacado)
    r"(.+?)\s+"             # sacado (nome, não-guloso)
    r"(\S+)\s+"             # numero do titulo
    r"(\d{2}/\d{2}/\d{4})\s+"  # vencimento
    r"([\d\.,]+)\s+"        # valor (face)
    r"(\d+)\s+"             # dias
    r"([\d\.,]+)\s+"        # fator/adv
    r"([\d\.,]+)\s+"        # impostos
    r"([\d\.,]+)\s+"        # tarifas
    r"([\d\.,]+)\s*$"       # liquido (valor pago)
)

# Detecta linhas que PARECEM ser um título (começam com um dos tipos
# conhecidos) mas que não bateram com o PADRAO_TITULO completo. Serve de
# rede de segurança para avisar quando um PDF vier com layout quebrado
# (ex: título dividido entre duas páginas) em vez de simplesmente ignorar
# a linha sem avisar ninguém.
PADRAO_TIPO_TITULO = re.compile(r"^(DM|DS|CHQ)\s")


def extrair_cheques(texto, cabecalho):
    """
    Retorna uma tupla (cheques, nao_reconhecidas).

    - cheques: lista de dicts com os títulos extraídos com sucesso.
    - nao_reconhecidas: lista de linhas que pareciam ser um título
      (começam com DM/DS/CHQ) mas não bateram com o padrão completo.
      Quem chama esta função decide o que fazer com isso (ex: imprimir
      na hora, acumular num resumo final, etc.).
    """

    cheques = []
    nao_reconhecidas = []

    linhas = texto.splitlines()

    for linha in linhas:

        linha_limpa = linha.strip()

        m = PADRAO_TITULO.match(linha_limpa)

        if not m:
            if PADRAO_TIPO_TITULO.match(linha_limpa):
                nao_reconhecidas.append(linha_limpa)
            continue

        (
            tipo,
            documento,
            sacado,
            numero,
            vencimento,
            valor,
            dias,
            fator_adv,
            impostos,
            tarifas,
            liquido,
        ) = m.groups()

        cheques.append({
            "cliente": cabecalho["cliente"],
            "data": cabecalho["data"],
            "lote": cabecalho["lote"],
            "tipo": tipo,
            "documento": documento,
            "sacado": sacado.strip(),
            "numero": numero,
            "vencimento": vencimento,
            "valor_face": valor,
            "dias": dias,
            "valor_pago": liquido,
        })

    return cheques, nao_reconhecidas
