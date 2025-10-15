from services.relatorio_service import RelatorioService
from pathlib import Path

class RelatorioController:
    """Controlador que faz a ponte entre a interface gráfica (GUI) e a camada de serviço.

    Responsabilidades:
    - validar entradas vindas da GUI (arquivo Excel, filtros, destino)
    - construir caminhos padrão quando necessário
    - chamar o service responsável pela geração do relatório

    O controller não realiza processamento pesado nem I/O direto (exceto determinar
    caminhos). Ele captura exceções do service e retorna tuplas (bool, mensagem)
    para a camada de apresentação exibir feedback ao usuário.
    """

    def __init__(self):
        self.service = RelatorioService()

    def processar_relatorio(self, excel_path, saida_txt, ano, mes, tipo):
        """Orquestra o fluxo entre GUI e Service, validando entradas."""
        if not excel_path:
            return False, "Nenhum arquivo Excel selecionado."

        # Define caminho padrão se não informado
        if not saida_txt or saida_txt.strip() == "":
            saida_txt = Path(excel_path).with_name("relatorio_alunos_por_mes.txt")

        try:
            return self.service.gerar_relatorio_completo(excel_path, saida_txt, ano, mes, tipo)
        except Exception as e:
            return False, f"Erro inesperado: {e}"
