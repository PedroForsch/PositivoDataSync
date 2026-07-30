
# ============================================================
# Importador de Borderos - Positivo DataSync
# Autor: Pedro Henrique Forsch
# Contato: pedrohenriqueforsch@gmail.com
# ============================================================

MESES = {
    1: "jan",
    2: "fev",
    3: "mar",
    4: "abr",
    5: "mai",
    6: "jun",
    7: "jul",
    8: "ago",
    9: "set",
    10: "out",
    11: "nov",
    12: "dez",
}

# Nomes completos, sem acento (a comparação sempre remove acentos antes),
# usados para aceitar abas escritas como "Julho 2026" e para montar a
# mensagem de erro "Nenhuma aba encontrada para julho/2026".
MESES_COMPLETO = {
    1: "janeiro",
    2: "fevereiro",
    3: "marco",
    4: "abril",
    5: "maio",
    6: "junho",
    7: "julho",
    8: "agosto",
    9: "setembro",
    10: "outubro",
    11: "novembro",
    12: "dezembro",
}

COLUNAS = {
    "cliente": "A",
    "data": "B",
    "lote": "C",
    "dias": "D",
    "valor_face": "E",
    "valor_pago": "F",
}
