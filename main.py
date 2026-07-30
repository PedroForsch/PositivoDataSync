
# ============================================================
# Importador de Borderos - Positivo DataSync
# Autor: Pedro Henrique Forsch
# Contato: pedrohenriqueforsch@gmail.com
# ============================================================

from pathlib import Path

from pdf_reader import (
    ler_pdf,
    extrair_cabecalho,
    extrair_cheques
)
from excel_writer import importar_para_excel
PASTA_ENTRADA = Path("entrada")
def carregar_pdfs():

    todos_cheques = []
    resumo_nao_reconhecidas = []

    arquivos = list(PASTA_ENTRADA.glob("*.pdf"))

    if not arquivos:

        print("Nenhum PDF encontrado.")

        return [], []

    print()
    print("=" * 60)
    print(f"{len(arquivos)} PDF(s) encontrado(s)")
    print("=" * 60)

    for pdf in arquivos:

        print()
        print(f"Lendo: {pdf.name}")

        try:

            texto = ler_pdf(pdf)

            cabecalho = extrair_cabecalho(texto)

            cheques, nao_reconhecidas = extrair_cheques(
                texto,
                cabecalho
            )

            print(
                f"{len(cheques)} cheque(s) encontrado(s)"
            )

            if nao_reconhecidas:

                print(
                    f"[AVISO] {len(nao_reconhecidas)} linha(s) "
                    f"nao reconhecida(s) neste arquivo"
                )

                resumo_nao_reconhecidas.append({
                    "arquivo": pdf.name,
                    "lote": cabecalho.get("lote", "?"),
                    "cliente": cabecalho.get("cliente", "?"),
                    "linhas": nao_reconhecidas,
                })

            todos_cheques.extend(
                cheques
            )

        except Exception as erro:

            print()

            print("=" * 60)
            print("ERRO")
            print("=" * 60)

            print(pdf.name)

            print()

            print(erro)

    return todos_cheques, resumo_nao_reconhecidas


def exibir_resumo_nao_reconhecidas(resumo):

    if not resumo:
        return

    print()
    print("!" * 60)
    print("ATENCAO: borderô(s) com linha(s) de titulo NAO reconhecida(s)")
    print("Esses titulos NAO foram importados. Confira manualmente:")
    print("!" * 60)

    for item in resumo:

        print()
        print(f"Arquivo  : {item['arquivo']}")
        print(f"Borderô  : {item['lote']}")
        print(f"Cliente  : {item['cliente']}")

        for linha in item["linhas"]:
            print(f"  - {linha!r}")

    print()
    print("!" * 60)


def main():

    cheques, resumo_nao_reconhecidas = carregar_pdfs()

    exibir_resumo_nao_reconhecidas(
        resumo_nao_reconhecidas
    )

    if not cheques:

        print()

        print("Nenhum cheque encontrado.")

        return

    print()

    print("=" * 60)
    print(f"TOTAL DE CHEQUES: {len(cheques)}")
    print("=" * 60)

    print()

    resposta = input(
        "Deseja importar para o Excel? (S/N): "
    ).strip().upper()

    if resposta != "S":

        print("Operação cancelada.")

        return

    importar_para_excel(
        cheques
    )

    print()

    print("=" * 60)
    print("IMPORTAÇÃO CONCLUÍDA")
    print("=" * 60)

    if resumo_nao_reconhecidas:
        print()
        print(
            f"Lembrete: {len(resumo_nao_reconhecidas)} arquivo(s) "
            f"tiveram título(s) não reconhecido(s) — veja o aviso acima."
        )


if __name__ == "__main__":

    main()
