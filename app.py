
# ============================================================
# Importador de Borderos - Positivo DataSync
# Autor: Pedro Henrique Forsch
# Contato: pedrohenriqueforsch@gmail.com
# ============================================================

import datetime
import os
import sys
import traceback
from pathlib import Path


CAMINHO_LOG_ERRO = Path.home() / "Desktop" / "ImportadorBorderos_erro.log"


def _checkpoint(mensagem):
    """Grava uma linha de progresso no log, mesmo antes de tudo estar importado.
    Serve para saber até onde o app chegou quando ele crasha sem gerar
    uma exceção Python capturável (ex: falha nativa da janela)."""
    try:
        with open(CAMINHO_LOG_ERRO, "a", encoding="utf-8") as arquivo:
            arquivo.write(f"[{datetime.datetime.now()}] {mensagem}\n")
    except Exception:
        pass


def registrar_erro_fatal(erro):
    try:
        with open(CAMINHO_LOG_ERRO, "a", encoding="utf-8") as arquivo:
            arquivo.write("=" * 60 + "\n")
            arquivo.write(f"{datetime.datetime.now()} - ERRO FATAL\n")
            arquivo.write("=" * 60 + "\n")
            arquivo.write("".join(traceback.format_exception(type(erro), erro, erro.__traceback__)))
            arquivo.write("\n")
    except Exception:
        pass


# Em apps empacotados com --noconsole, sys.stdout/sys.stderr podem vir como
# None. Se qualquer biblioteca (ou nosso próprio código) tentar usar print(),
# isso crasha o app inteiro sem deixar rastro nenhum. Corrigimos isso ANTES
# de importar qualquer outra coisa.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

# O motor de renderização do WebView2 (usado pelo pywebview no Windows) às
# vezes engasga por alguns segundos a cada interação quando a aceleração
# por GPU tem problema com o driver de vídeo da máquina (comum em notebooks
# com placa de vídeo integrada mais antiga, ou com antivírus escaneando os
# processos auxiliares que o WebView2 cria a cada repintura). Desativar a
# GPU resolve esse tipo de travamento, ao custo de uma renderização
# levemente mais pesada para a CPU — imperceptível numa tela simples como
# a nossa.
os.environ["WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS"] = (
    "--disable-gpu --disable-software-rasterizer --disable-gpu-compositing"
)

_checkpoint("Iniciando aplicativo.")

try:
    import json
    import threading
    import shutil
    import hashlib

    _checkpoint("Bibliotecas padrao do Python importadas com sucesso.")

    import webview

    _checkpoint("pywebview importado com sucesso.")

    from pdf_reader import ler_pdf, extrair_cabecalho, extrair_cheques

    _checkpoint("pdf_reader importado com sucesso.")

    from excel_writer import ExcelWriter, AbaNaoEncontrada

    _checkpoint("excel_writer importado com sucesso (xlwings carregado).")

except Exception as erro:
    registrar_erro_fatal(erro)
    raise


def caminho_recurso(relativo):
    """Resolve caminhos tanto rodando via python quanto empacotado pelo PyInstaller."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return str(base / relativo)


def pasta_app():
    """Pasta onde o .exe (ou o app.py, rodando via python puro) está.
    Diferente de caminho_recurso: essa pasta é gravável (a _MEIPASS do
    PyInstaller, usada só para ler recursos como o index.html, pode ser
    uma pasta temporária)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def fazer_backup_planilha(caminho_planilha):
    """Copia a planilha para uma subpasta 'backups' dentro da pasta do
    aplicativo, com data/hora no nome, antes de qualquer escrita."""
    origem = Path(caminho_planilha)

    pasta_backups = pasta_app() / "backups"
    pasta_backups.mkdir(exist_ok=True)

    carimbo = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    destino = pasta_backups / f"{origem.stem} - backup {carimbo}{origem.suffix}"

    shutil.copy2(origem, destino)
    return destino


def hash_arquivo(caminho, tamanho_bloco=1024 * 1024):
    """Calcula o hash SHA-256 do conteúdo de um arquivo, para detectar
    dois PDFs com conteúdo idêntico mesmo que tenham nomes diferentes."""
    sha256 = hashlib.sha256()

    with open(caminho, "rb") as arquivo:
        for bloco in iter(lambda: arquivo.read(tamanho_bloco), b""):
            sha256.update(bloco)

    return sha256.hexdigest()


class Api:
    def __init__(self):
        self.window = None
        self.pdfs = []
        self.planilha = None

    # ---------- seleção de arquivos ----------

    def selecionar_pdfs(self):
        arquivos = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=True,
            file_types=("Arquivos PDF (*.pdf)",),
        )

        if arquivos:
            for caminho in arquivos:
                if caminho not in self.pdfs:
                    self.pdfs.append(caminho)

        return [Path(p).name for p in self.pdfs]

    def remover_pdf(self, indice):
        if 0 <= indice < len(self.pdfs):
            self.pdfs.pop(indice)

        return [Path(p).name for p in self.pdfs]

    def selecionar_planilha(self):
        arquivos = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Planilhas Excel (*.xlsm;*.xlsx;*.xls)",),
        )

        if arquivos:
            self.planilha = arquivos[0]
            return Path(self.planilha).name

        return None

    def remover_planilha(self):
        self.planilha = None
        return None

    # ---------- processamento ----------

    def _log(self, mensagem):
        if self.window:
            self.window.evaluate_js(
                f"appendLog({json.dumps(mensagem)})"
            )

    def _progresso(self, atual, total, rotulo):
        if self.window:
            self.window.evaluate_js(
                f"atualizarProgresso({atual}, {total}, {json.dumps(rotulo)})"
            )

    def _esconder_progresso(self):
        if self.window:
            self.window.evaluate_js("esconderProgresso()")

    def processar(self):
        if not self.pdfs:
            return {"ok": False, "mensagem": "Selecione ao menos um PDF na aba 'Arquivos PDF'."}

        if not self.planilha:
            return {"ok": False, "mensagem": "Selecione a planilha Excel na aba 'Planilha Excel'."}

        thread = threading.Thread(target=self._processar_thread, daemon=True)
        thread.start()

        return {"ok": True}

    def _processar_thread(self):
        try:
            self._log("=" * 60)
            self._log(f"{len(self.pdfs)} PDF(s) para processar")
            self._log("=" * 60)

            todos_cheques = []
            erros_leitura = 0
            arquivos_sem_titulo = []
            hashes_vistos = {}

            total_pdfs = len(self.pdfs)

            for indice, caminho in enumerate(self.pdfs, start=1):
                nome = Path(caminho).name

                self._progresso(indice, total_pdfs, "Lendo PDFs")

                try:
                    hash_atual = hash_arquivo(caminho)
                except Exception:
                    hash_atual = None

                if hash_atual and hash_atual in hashes_vistos:
                    self._log(
                        f"[AVISO] {nome}: conteúdo idêntico a "
                        f"{hashes_vistos[hash_atual]} - PDF duplicado, não foi processado."
                    )
                    continue

                if hash_atual:
                    hashes_vistos[hash_atual] = nome

                try:
                    texto = ler_pdf(caminho)
                    cabecalho = extrair_cabecalho(texto)
                    cheques, nao_reconhecidas = extrair_cheques(texto, cabecalho)

                    if not cheques:
                        self._log(
                            f"[ERRO] {nome}: nenhum título foi lido "
                            f"(esperado pelo menos 1). Confira o PDF manualmente."
                        )
                        arquivos_sem_titulo.append(nome)
                    else:
                        self._log(
                            f"{nome}: {len(cheques)} título(s) "
                            f"({cabecalho['cliente']})"
                        )

                    if nao_reconhecidas:
                        self._log(
                            f"  [AVISO] {len(nao_reconhecidas)} linha(s) "
                            f"nao reconhecida(s) neste arquivo"
                        )

                    for cheque in cheques:
                        cheque["arquivo"] = nome

                    todos_cheques.extend(cheques)

                except Exception as erro:
                    self._log(f"[ERRO] {nome}: {erro}")
                    erros_leitura += 1
                    arquivos_sem_titulo.append(nome)

            if not todos_cheques:
                self._log("")
                self._log("Nenhum título encontrado nos PDFs.")

                if arquivos_sem_titulo:
                    self._log("")
                    self._log("Arquivo(s) sem nenhum título lido:")
                    for nome_arquivo in arquivos_sem_titulo:
                        self._log(f"  - {nome_arquivo}")

                self._esconder_progresso()

                self.window.evaluate_js(
                    "processoFinalizado("
                    + json.dumps({
                        "ok": False,
                        "mensagem": "Nenhum título encontrado nos PDFs.",
                        "arquivos_sem_titulo": arquivos_sem_titulo,
                    })
                    + ")"
                )
                return

            self._log("")
            self._log(f"Total de títulos a importar: {len(todos_cheques)}")

            try:
                destino_backup = fazer_backup_planilha(self.planilha)
                self._log(f"Backup da planilha salvo em: backups/{destino_backup.name}")
            except Exception as erro_backup:
                self._log(f"[AVISO] Falha ao fazer backup da planilha: {erro_backup}")

            self._log("Abrindo planilha...")

            excel = ExcelWriter()
            excel.abrir(self.planilha)

            total_importado = 0
            erros_importacao = 0
            resumo_abas = {}
            erros_detalhados = []
            cache_abas = {}
            abas_anunciadas = set()

            try:
                total_cheques = len(todos_cheques)

                for indice, cheque in enumerate(todos_cheques, start=1):
                    linha = None

                    self._progresso(indice, total_cheques, "Gravando na planilha")

                    try:
                        data_str = cheque["data"]

                        if data_str not in cache_abas:
                            try:
                                aba_encontrada = excel.obter_aba(data_str)
                                cache_abas[data_str] = aba_encontrada

                                if aba_encontrada.name not in abas_anunciadas:
                                    self._log(f"Aba: {aba_encontrada.name}")
                                    abas_anunciadas.add(aba_encontrada.name)

                            except AbaNaoEncontrada as erro_aba:
                                cache_abas[data_str] = erro_aba
                                chave_erro = str(erro_aba)

                                if chave_erro not in abas_anunciadas:
                                    self._log("[ERRO]")
                                    for linha_erro in chave_erro.splitlines():
                                        self._log(f"  {linha_erro}")
                                    abas_anunciadas.add(chave_erro)

                        resultado = cache_abas[data_str]

                        if isinstance(resultado, Exception):
                            raise resultado

                        aba = resultado

                        resumo_abas.setdefault(aba.name, {"total": 0, "importados": 0})
                        resumo_abas[aba.name]["total"] += 1

                        linha = excel.primeira_linha_vazia(aba)
                        excel.escrever_cheque(aba, linha, cheque)

                        total_importado += 1
                        resumo_abas[aba.name]["importados"] += 1

                    except Exception as erro:
                        linha_str = linha if linha is not None else "N/A"

                        self._log(
                            f"[ERRO] {cheque['cliente']} "
                            f"({cheque.get('arquivo', '?')}), linha {linha_str}: {erro}"
                        )

                        erros_importacao += 1

                        erros_detalhados.append({
                            "arquivo": cheque.get("arquivo", "?"),
                            "cliente": cheque["cliente"],
                            "lote": cheque["lote"],
                            "linha": linha_str,
                            "motivo": str(erro),
                        })

            finally:
                excel.fechar()

            self._log("")

            for nome_aba, dados in resumo_abas.items():
                self._log("=" * 60)
                self._log(f"PLANILHA   : {Path(self.planilha).name}")
                self._log(f"ABA        : {nome_aba}")
                self._log(f"CHEQUES    : {dados['total']}")
                self._log(f"IMPORTADOS : {dados['importados']}")
                self._log("=" * 60)

            if erros_detalhados:
                self._log("")
                self._log("RESUMO DE ERROS")
                self._log("-" * 60)

                for erro_item in erros_detalhados:
                    self._log(f"Arquivo : {erro_item['arquivo']}")
                    self._log(f"Cliente : {erro_item['cliente']}")
                    self._log(f"Lote    : {erro_item['lote']}")
                    self._log(f"Linha   : {erro_item['linha']}")
                    self._log(f"Motivo  : {erro_item['motivo']}")
                    self._log("-" * 60)

            self._log("")
            self._log("=" * 60)
            self._log("RESUMO GERAL")
            self._log(f"Arquivos processados   : {len(self.pdfs)}")
            self._log(f"Titulos encontrados    : {len(todos_cheques)}")
            self._log(f"Titulos importados     : {total_importado}")
            self._log(f"Erros de leitura       : {erros_leitura}")
            self._log(f"Erros de importacao    : {erros_importacao}")
            self._log("=" * 60)

            if arquivos_sem_titulo:
                self._log("")
                self._log("[ATENCAO] Arquivo(s) sem nenhum título lido "
                           "(confira manualmente, pode haver título perdido):")
                for nome_arquivo in arquivos_sem_titulo:
                    self._log(f"  - {nome_arquivo}")

            resultado = {
                "ok": True,
                "total": len(todos_cheques),
                "importado": total_importado,
                "erros_leitura": erros_leitura,
                "erros_importacao": erros_importacao,
                "arquivos_sem_titulo": arquivos_sem_titulo,
            }

            self._esconder_progresso()

            self.window.evaluate_js(
                f"processoFinalizado({json.dumps(resultado)})"
            )

        except Exception as erro:
            self._log(f"[ERRO GERAL] {erro}")
            self._esconder_progresso()

            self.window.evaluate_js(
                "processoFinalizado("
                + json.dumps({"ok": False, "mensagem": str(erro)})
                + ")"
            )


def main():
    _checkpoint("Criando instancia da Api.")
    api = Api()

    _checkpoint("Chamando webview.create_window()...")
    window = webview.create_window(
        "Positivo DataSync - Importador de Borderôs",
        caminho_recurso("index.html"),
        js_api=api,
        width=1080,
        height=760,
        min_size=(860, 620),
    )
    _checkpoint("Janela criada com sucesso.")

    api.window = window

    _checkpoint("Chamando webview.start()...")
    webview.start(gui="edgechromium")
    _checkpoint("webview.start() encerrado normalmente (janela fechada pelo usuario).")


if __name__ == "__main__":
    try:
        main()
    except Exception as erro:
        registrar_erro_fatal(erro)
        raise
