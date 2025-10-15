from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from pathlib import Path

def salvar_txt(linhas, path_saida):
    """Salva o relatório em formato TXT.

    Parâmetros:
    - linhas: lista de strings que representam as linhas do relatório
    - path_saida: caminho do arquivo de saída (string ou Path)

    Observação: sobrescreve o arquivo se já existir.
    """
    with open(path_saida, "w", encoding="utf-8") as f:
        f.write("\n".join(linhas))


def salvar_pdf(linhas, path_saida):
    """Salva o relatório em formato PDF.

    Nota: a função existe para compatibilidade, mas o fluxo atual da aplicação
    pode não chamar a geração de PDF (dependendo da configuração do controller).
    """
    c = canvas.Canvas(path_saida, pagesize=A4)
    largura, altura = A4
    y = altura - 50
    c.setFont("Helvetica", 11)

    for linha in linhas:
        if y < 50:
            c.showPage()
            c.setFont("Helvetica", 11)
            y = altura - 50
        c.drawString(50, y, linha)
        y -= 18

    c.save()
