from utils.excel_utils import ler_planilha, aplicar_filtros
from utils.file_utils import salvar_txt
from pathlib import Path

class RelatorioService:
    """Camada de negócio responsável pela lógica de geração dos relatórios."""

    def gerar_relatorio_completo(self, excel_path, saida_txt, ano, mes, tipo):
        df = ler_planilha(excel_path)
        df_filtrado = aplicar_filtros(df, ano, mes, tipo)

        if df_filtrado.empty:
            return False, f"Nenhum resultado encontrado para o filtro: {tipo}"

        # Agrupar e montar texto
        linhas = []
        for (ano, mes_num), grupo in df_filtrado.groupby(["ANO", "MES_NUM"]):
            mes_nome = grupo["MES_NOME"].iloc[0]
            linhas.append("-----------------------------")
            linhas.append(f"{mes_nome} ({ano})")
            for _, linha in grupo.iterrows():
                linhas.append(f"{linha['MATRICULA']} - {linha['NOME ALUNO']}")
            linhas.append("")
        linhas.append("-----------------------------")

        # Salvar arquivo TXT
        salvar_txt(linhas, saida_txt)

        return True, f"Relatório gerado:\n\nTXT → {Path(saida_txt).resolve()}"
