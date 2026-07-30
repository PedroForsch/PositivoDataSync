
# ============================================================
# Importador de Borderos - Positivo DataSync
# Autor: Pedro Henrique Forsch
# Contato: pedrohenriqueforsch@gmail.com
# ============================================================

from datetime import datetime
from pathlib import Path
import re

import xlwings as xw

from config import COLUNAS, MESES, MESES_COMPLETO


class AbaNaoEncontrada(Exception):
    """Levantado quando nenhuma aba da planilha corresponde ao mês/ano do cheque."""
    pass


def _normalizar(texto):
    """Deixa minúsculo, remove acentos comuns e espaços duplicados, para
    comparar nomes de aba sem se importar com maiúsculas/acentuação."""
    texto = texto.strip().lower()

    for com_acento, sem_acento in (
        ("ç", "c"), ("ã", "a"), ("õ", "o"), ("á", "a"), ("â", "a"),
        ("é", "e"), ("ê", "e"), ("í", "i"), ("ó", "o"), ("ô", "o"), ("ú", "u"),
    ):
        texto = texto.replace(com_acento, sem_acento)

    return " ".join(texto.split())


_NOME_MES_PARA_NUMERO = {}
for _numero, _abrev in MESES.items():
    _NOME_MES_PARA_NUMERO[_abrev] = _numero
for _numero, _completo in MESES_COMPLETO.items():
    _NOME_MES_PARA_NUMERO[_normalizar(_completo)] = _numero


def _mes_ano_da_aba(nome_aba):
    """
    Tenta interpretar o nome de uma aba como um mês/ano de operação,
    aceitando formatos como "jul 26", "Jul 26", "JUL 26", "Julho 2026",
    "07/2026", "07-2026". Retorna (mes, ano) com ano de 4 dígitos, ou None
    se o nome da aba não bater com nenhum formato conhecido (ex: "Empresas").
    """
    texto = _normalizar(nome_aba)

    m = re.fullmatch(r"(\d{1,2})[/\-](\d{4})", texto)
    if m:
        mes, ano = int(m.group(1)), int(m.group(2))
        return (mes, ano) if 1 <= mes <= 12 else None

    m = re.fullmatch(r"([a-z]+)[\s/\-]*(\d{2,4})", texto)
    if m:
        nome_mes, ano_texto = m.groups()
        mes = _NOME_MES_PARA_NUMERO.get(nome_mes)
        if mes is None:
            return None
        ano = int(ano_texto)
        if ano < 100:
            ano += 2000
        return mes, ano

    return None


class ExcelWriter:
    def __init__(self):
        self.app = None
        self.wb = None

    def abrir(self , caminho=None):
        if caminho is None:
            pasta_planilha = (
                Path(__file__).resolve().parent / "planilha"
            )

            arquivos = []

            for extensao in ("*.xlsm", "*.xlsx", "*.xls"):
                arquivos.extend(
                    pasta_planilha.glob(extensao)
                )

            arquivos = [
                arquivo
                for arquivo in arquivos
                if not arquivo.name.startswith("~$")
            ]

            if not arquivos:
                raise FileNotFoundError(
                    f"Nenhuma planilha encontrada em:\n"
                    f"{pasta_planilha}"
                )

            if len(arquivos) > 1:
                nomes = "\n".join(
                    f"- {arquivo.name}"
                    for arquivo in arquivos
                )

                raise RuntimeError(
                    "Foi encontrada mais de uma planilha:\n"
                    f"{nomes}\n\n"
                    "Deixe apenas uma planilha na pasta."
                )

            planilha = arquivos[0]
        else:
            planilha = Path(caminho)

            if not planilha.exists():
                raise FileNotFoundError(
                    f"Planilha não encontrada\n{planilha}"
                )

        print(
            f"Planilha encontrada: {planilha.name}"
        )

        self.app = xw.App(visible=False)

        self.app.display_alerts = False
        self.app.screen_updating = False

        self.wb = self.app.books.open(
            str(planilha)
        )

    def fechar(self):
        if self.wb:
            self.wb.save()
            self.wb.close()

        if self.app:
            self.app.quit()

    def nome_aba(self, data_operacao):
        data = datetime.strptime(
            data_operacao,
            "%d/%m/%Y"
        )

        mes = MESES[data.month]

        ano = str(data.year)[2:]

        return f"{mes} {ano}"

    def obter_aba(self, data_operacao):
        """
        Lê todas as abas da planilha e procura a que corresponde ao mês/ano
        do cheque (ignorando maiúsculas/minúsculas e aceitando vários
        formatos: "jul 26", "Julho 2026", "07/2026", etc). Levanta
        AbaNaoEncontrada com a lista de abas existentes se não encontrar.
        """
        data = datetime.strptime(data_operacao, "%d/%m/%Y")
        mes_alvo, ano_alvo = data.month, data.year

        for aba in self.wb.sheets:
            if _mes_ano_da_aba(aba.name) == (mes_alvo, ano_alvo):
                return aba

        nome_mes = MESES_COMPLETO[mes_alvo]
        abas_existentes = "\n".join(aba.name for aba in self.wb.sheets)

        raise AbaNaoEncontrada(
            f"Nenhuma aba encontrada para {nome_mes}/{ano_alvo}\n\n"
            f"Abas existentes:\n{abas_existentes}"
        )

    def primeira_linha_vazia(self, aba, limite=10000):
        # Leitura em lote: uma única chamada COM traz a coluna inteira,
        # em vez de uma chamada por célula. Isso evita deixar o processo
        # do Excel "ocupado" por muito tempo durante o loop, que era o
        # que causava o "Call was rejected by callee" / travamentos
        # quando o usuário mexia no computador antes do processo terminar.
        coluna = COLUNAS['cliente']

        valores = aba.range(
            f"{coluna}1:{coluna}{limite}"
        ).value

        for indice, valor in enumerate(valores, start=1):
            if valor is None:
                return indice

            if str(valor).strip() == "":
                return indice

        # Nenhuma linha vazia encontrada dentro do limite: assume a
        # próxima linha após o limite em vez de travar silenciosamente.
        return limite + 1

    def escrever_cheque(
        self,
        aba,
        linha,
        cheque
    ):
        aba.range(
            f"{COLUNAS['cliente']}{linha}"
        ).value = cheque["cliente"]

        data = datetime.strptime(
        cheque["data"],
        "%d/%m/%Y"
        )

        celula_data = aba.range(
        f"{COLUNAS['data']}{linha}"
        )

        celula_data.value = data
        celula_data.number_format = "dd/mmm"

        aba.range(
            f"{COLUNAS['lote']}{linha}"
        ).value = cheque["lote"]

        aba.range(
            f"{COLUNAS['dias']}{linha}"
        ).value = int(
            cheque["dias"]
        )

        valor_face = float(
            cheque["valor_face"]
            .replace(".", "")
            .replace(",", ".")
        )

        valor_pago = float(
            cheque["valor_pago"]
            .replace(".", "")
            .replace(",", ".")
        )

        aba.range(
            f"{COLUNAS['valor_face']}{linha}"
        ).value = valor_face

        aba.range(
            f"{COLUNAS['valor_pago']}{linha}"
        ).value = valor_pago

    def importar_cheques(self, cheques):
        total_importado = 0
        resumo_abas = {}
        erros_detalhados = []
        abas_ja_anunciadas = set()

        for cheque in cheques:
            linha = None

            try:
                aba = self.obter_aba(cheque["data"])

                if cheque["data"] not in abas_ja_anunciadas:
                    print(f"Data do cheque : {cheque['data']}")
                    print(f"Aba encontrada : {aba.name}")
                    abas_ja_anunciadas.add(cheque["data"])

                resumo_abas.setdefault(aba.name, {"total": 0, "importados": 0})
                resumo_abas[aba.name]["total"] += 1

                linha = self.primeira_linha_vazia(aba)
                self.escrever_cheque(aba, linha, cheque)

                print(f"[OK] {cheque['cliente']} -> {aba.name} (linha {linha})")

                total_importado += 1
                resumo_abas[aba.name]["importados"] += 1

            except Exception as erro:
                if isinstance(erro, AbaNaoEncontrada):
                    print()
                    print("ERRO")
                    print(str(erro))
                else:
                    print()
                    print("=" * 60)
                    print("ERRO AO IMPORTAR")
                    print("=" * 60)
                    print(f"Cliente : {cheque['cliente']}")
                    print(f"Lote    : {cheque['lote']}")
                    print(f"Data    : {cheque['data']}")
                    print(f"Linha   : {linha if linha is not None else 'N/A'}")
                    print()
                    print(erro)
                    print("=" * 60)

                erros_detalhados.append({
                    "arquivo": cheque.get("arquivo", "?"),
                    "cliente": cheque["cliente"],
                    "lote": cheque["lote"],
                    "linha": linha if linha is not None else "N/A",
                    "motivo": str(erro),
                })

        print()
        for nome_aba, dados in resumo_abas.items():
            print("=" * 60)
            print(f"ABA        : {nome_aba}")
            print(f"CHEQUES    : {dados['total']}")
            print(f"IMPORTADOS : {dados['importados']}")
            print("=" * 60)

        if erros_detalhados:
            print()
            print("RESUMO DE ERROS")
            print("-" * 60)
            for erro_item in erros_detalhados:
                print(f"Arquivo : {erro_item['arquivo']}")
                print(f"Cliente : {erro_item['cliente']}")
                print(f"Lote    : {erro_item['lote']}")
                print(f"Linha   : {erro_item['linha']}")
                print(f"Motivo  : {erro_item['motivo']}")
                print("-" * 60)

        print()
        print("=" * 60)
        print(f"TOTAL IMPORTADO: {total_importado} de {len(cheques)}")
        print("=" * 60)


def importar_para_excel(cheques):
    excel = ExcelWriter()

    try:
        excel.abrir()
        excel.importar_cheques(cheques)

    finally:
        excel.fechar()
