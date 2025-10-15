import ttkbootstrap as ttk
from ttkbootstrap.constants import INFO, SECONDARY, SUCCESS
from tkinter import filedialog, messagebox
from controllers.relatorio_controller import RelatorioController
from pathlib import Path
import os
import tkinter as tk
import webbrowser


class RelatorioApp:
    """Aplicação GUI principal para geração e visualização de relatórios.

    Responsabilidades principais:
    - montar a interface gráfica com abas (gerar / visualizar)
    - coletar parâmetros do usuário (arquivo Excel, filtros, destino)
    - delegar a geração do relatório para o controller

    A GUI usa ttkbootstrap para aparência; objetos tkinter principais são
    armazenados como atributos da instância (por ex. `self.text_area`).
    """
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Gerador de Relatório Mensal - Interatech")
        self.root.geometry("900x600")
        self.root.resizable(False, False)

        self.controller = RelatorioController()

        # Variáveis
        self.caminho_excel = ttk.StringVar()
        self.caminho_saida = ttk.StringVar(value="relatorio_alunos_por_mes.txt")
        self.ano_filtrado = ttk.StringVar(value="Todos")
        self.mes_filtrado = ttk.StringVar(value="Todos")
        self.tipo_filtro = ttk.StringVar(value="Todos")

        self._montar_interface()

    def _montar_interface(self):
        # Cria o container de abas e duas abas principais: gerar e visualizar
        self.tabs = ttk.Notebook(self.root, bootstyle="primary")
        self.tabs.pack(fill=tk.BOTH, expand=tk.YES)

        self.frame_gerar = ttk.Frame(self.tabs, padding=20)
        self.frame_visualizar = ttk.Frame(self.tabs, padding=10)
        self.tabs.add(self.frame_gerar, text="🧾 Gerar Relatório")
        self.tabs.add(self.frame_visualizar, text="📄 Relatório Gerado")

        self._montar_aba_gerar()
        self._montar_aba_visualizar()

    # --------------------------
    # ABA 1 - GERAÇÃO
    # --------------------------
    def _montar_aba_gerar(self):
        frame = self.frame_gerar
        # Área de seleção do arquivo Excel e filtros
        ttk.Label(frame, text="📘 Arquivo Excel:", font=("Segoe UI", 11, "bold")).pack(anchor=tk.W)
        ttk.Entry(frame, textvariable=self.caminho_excel, width=80).pack(pady=5)
        ttk.Button(frame, text="Selecionar arquivo", bootstyle=INFO, command=self._escolher_arquivo).pack()

        ttk.Separator(frame, bootstyle=SECONDARY).pack(fill=tk.X, pady=10)

    # Filtros
        filtro_frame = ttk.Frame(frame)
        filtro_frame.pack(pady=5)

        meses = ["Todos", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                 "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        anos = ["Todos"] + [str(ano) for ano in range(2015, 2031)]
        tipos = ["Todos", "Compra de Crédito", "Outros"]

        ttk.Label(filtro_frame, text="📅 Mês:").grid(row=0, column=0, padx=5)
        ttk.Combobox(filtro_frame, textvariable=self.mes_filtrado, values=meses, width=20).grid(row=0, column=1)

        ttk.Label(filtro_frame, text="Ano:").grid(row=0, column=2, padx=10)
        ttk.Combobox(filtro_frame, textvariable=self.ano_filtrado, values=anos, width=15).grid(row=0, column=3)

        ttk.Label(frame, text="🎯 Tipo de filtro:").pack(anchor=tk.W, pady=(10, 2))
        ttk.Combobox(frame, textvariable=self.tipo_filtro, values=tipos, width=25).pack()

        # Separador e área de destino/ação
        ttk.Separator(frame, bootstyle=SECONDARY).pack(fill=tk.X, pady=10)

        ttk.Label(frame, text="💾 Local de saída:").pack(anchor=tk.W)
        ttk.Entry(frame, textvariable=self.caminho_saida, width=80).pack(pady=5)
        ttk.Button(frame, text="Selecionar destino", bootstyle=INFO, command=self._escolher_destino).pack()

        ttk.Button(frame, text="🚀 Gerar Relatório", bootstyle=SUCCESS, command=self._gerar_relatorio).pack(pady=15)
        ttk.Button(frame, text="📂 Abrir Pasta", bootstyle=SECONDARY, command=self._abrir_pasta).pack(pady=5)

    # --------------------------
    # ABA 2 - VISUALIZAÇÃO
    # --------------------------
    def _montar_aba_visualizar(self):
        # Aba para visualizar o conteúdo do TXT gerado
        ttk.Label(self.frame_visualizar, text="📄 Visualização do Relatório Gerado",
                  font=("Segoe UI", 12, "bold")).pack(anchor=tk.W, pady=5)

        self.text_area = tk.Text(self.frame_visualizar, height=25, width=110,
                                 wrap="word", bg="#1e1e1e", fg="white", insertbackground="white")
        self.text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=tk.YES, padx=5, pady=10)

        scroll = ttk.Scrollbar(self.frame_visualizar, command=self.text_area.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.text_area.config(yscrollcommand=scroll.set)

        ttk.Button(self.frame_visualizar, text="🔍 Abrir no Bloco de Notas",
                   bootstyle=INFO, command=self._abrir_txt).pack(pady=5)

    # --------------------------
    # FUNÇÕES DE CONTROLE
    # --------------------------
    def _escolher_arquivo(self):
        arquivo = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx *.xls")])
        if arquivo:
            self.caminho_excel.set(arquivo)

    def _escolher_destino(self):
        destino = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Arquivo de texto", "*.txt")])
        if destino:
            self.caminho_saida.set(destino)

    def _gerar_relatorio(self):
        excel = self.caminho_excel.get()
        saida_txt = self.caminho_saida.get()
        ano, mes, tipo = self.ano_filtrado.get(), self.mes_filtrado.get(), self.tipo_filtro.get()

        if not excel:
            messagebox.showwarning("Aviso", "Selecione o arquivo Excel primeiro!")
            return

        # Se o usuário não definiu destino, cria um arquivo padrão ao lado do Excel
        if not saida_txt or saida_txt.strip() == "":
            saida_txt = str(Path(excel).with_name("relatorio_alunos_por_mes.txt"))
            self.caminho_saida.set(saida_txt)

        sucesso, mensagem = self.controller.processar_relatorio(excel, saida_txt, ano, mes, tipo)

        if sucesso:
            messagebox.showinfo("Sucesso", mensagem)
            self._mostrar_relatorio(saida_txt)
            self.tabs.select(self.frame_visualizar)
        else:
            messagebox.showerror("Erro", mensagem)

    def _mostrar_relatorio(self, caminho_txt):
        # Lê o arquivo TXT gerado e exibe na text_area. Mostra diálogo em caso de erro.
        try:
            with open(caminho_txt, "r", encoding="utf-8") as f:
                conteudo = f.read()
            self.text_area.delete(1.0, "end")
            self.text_area.insert("end", conteudo)
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def _abrir_txt(self):
        if os.path.exists(self.caminho_saida.get()):
            os.startfile(self.caminho_saida.get())
        else:
            messagebox.showwarning("Aviso", "O arquivo ainda não foi gerado.")

    def _abrir_pasta(self):
        pasta = Path(self.caminho_saida.get()).parent
        if pasta.exists():
            webbrowser.open(str(pasta))


if __name__ == "__main__":
    app = ttk.Window(themename="superhero")
    RelatorioApp(app)
    app.mainloop()
