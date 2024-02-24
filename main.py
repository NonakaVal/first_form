import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedStyle
import os
import subprocess
from openpyxl import Workbook


class AplicacaoCliente:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Gerenciamento")

        # Crie um estilo temático
        style = ThemedStyle(self.root)

        # Defina a cor da barra superior (substitua "blue" pela cor desejada)
        style.configure("TFrame", background="#RRGGBB")

        try:
            imagem = tk.PhotoImage(file=icone_path)
            self.root.tk.call('wm', 'iconphoto', self.root._w, imagem)
        except Exception as e:
            print(f"Erro ao definir o ícone na aba na barra: {e}")

        self.nome_var = tk.StringVar()
        self.email_var = tk.StringVar()
        self.telefone_var = tk.StringVar()
        self.endereco_var = tk.StringVar()
        self.pesquisa_var = tk.StringVar()
        self.clientes = []

        self.criar_formulario_cliente()
        self.criar_tabela_cliente()

        abrir_excel_button = ttk.Button(self.root, text="Abrir Excel na Pasta", command=self.abrir_excel)
        abrir_excel_button.grid(row=1, column=0, pady=10, padx=98, sticky="w")

        self.criar_planilha_excel()

    def criar_planilha_excel(self):
        self.workbook = Workbook()
        self.planilha = self.workbook.active

        # Adicione os cabeçalhos
        colunas_cliente = ["Nome", "E-mail", "Telefone", "Endereço"]
        self.planilha.append(colunas_cliente)

    def salvar_planilha_excel(self):
        planilha_path = "clientes.xlsx"
        self.workbook.save(planilha_path)

    def criar_formulario_cliente(self):
        formulario_frame = ttk.LabelFrame(self.root, text="Dados do Cliente", style="Accent.TLabelframe",
                                          padding=(10, 5))
        formulario_frame.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        # Nome
        nome_label = ttk.Label(formulario_frame, text="Nome:", style="Accent.TLabel")
        nome_label.grid(row=0, column=0, padx=5, pady=10, sticky="w")
        nome_entry = ttk.Entry(formulario_frame, textvariable=self.nome_var, width=30)
        nome_entry.grid(row=0, column=1, padx=5, pady=10, sticky="w")

        # E-mail
        email_label = ttk.Label(formulario_frame, text="E-mail:", style="Accent.TLabel")
        email_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        email_entry = ttk.Entry(formulario_frame, textvariable=self.email_var, width=30)
        email_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # Telefone
        telefone_label = ttk.Label(formulario_frame, text="Telefone:", style="Accent.TLabel")
        telefone_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        telefone_entry = ttk.Entry(formulario_frame, textvariable=self.telefone_var, width=30)
        telefone_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        # Endereço
        endereco_label = ttk.Label(formulario_frame, text="Endereço:", style="Accent.TLabel")
        endereco_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")
        endereco_entry = ttk.Entry(formulario_frame, textvariable=self.endereco_var, width=30)
        endereco_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        # Botão Adicionar Cliente
        adicionar_cliente_button = ttk.Button(formulario_frame, text="Adicionar Cliente",
                                              command=self.adicionar_cliente, style="Accent.TButton")
        adicionar_cliente_button.grid(row=4, columnspan=2, pady=10)

        # Entry para pesquisa
        pesquisa_entry = ttk.Entry(formulario_frame, textvariable=self.pesquisa_var, width=30)
        pesquisa_entry.grid(row=5, column=1, padx=5, pady=5, sticky="w")

        # Botão Pesquisar Cliente
        pesquisar_cliente_button = ttk.Button(formulario_frame, text="Pesquisar Cliente",
                                              command=self.pesquisar_cliente, style="Accent.TButton")
        pesquisar_cliente_button.grid(row=6, columnspan=2, pady=10)

    def criar_tabela_cliente(self):
        tabela_frame = ttk.LabelFrame(self.root, text="Lista de Clientes", style="Accent.TLabelframe",
                                      padding=(10, 5))
        tabela_frame.grid(row=0, column=1, padx=10, pady=10, rowspan=2, sticky="nsew")

        # Criar a tabela de clientes
        colunas_cliente = ("Nome", "E-mail", "Telefone", "Endereço")
        self.tabela_cliente = ttk.Treeview(tabela_frame, columns=colunas_cliente, show="headings",
                                           style="Treeview.Heading")

        # Configurar as colunas
        for coluna in colunas_cliente:
            self.tabela_cliente.heading(coluna, text=coluna)
            self.tabela_cliente.column(coluna, width=150)

        self.tabela_cliente.grid(row=0, column=0, sticky="nsew")

        # Adicionar barra de rolagem
        scrollbar_cliente = ttk.Scrollbar(tabela_frame, orient="vertical", command=self.tabela_cliente.yview)
        scrollbar_cliente.grid(row=0, column=1, sticky="ns")
        self.tabela_cliente.configure(yscrollcommand=scrollbar_cliente.set)

    def adicionar_cliente(self):
        nome = self.nome_var.get()
        email = self.email_var.get()
        telefone = self.telefone_var.get()
        endereco = self.endereco_var.get()

        novo_cliente = {"Nome": nome, "E-mail": email, "Telefone": telefone, "Endereço": endereco}
        self.clientes.append(novo_cliente)
        self.tabela_cliente.insert("", "end", values=list(novo_cliente.values()))
        self.planilha.append(list(novo_cliente.values()))
        self.salvar_planilha_excel()
        self.nome_var.set("")
        self.email_var.set("")
        self.telefone_var.set("")
        self.endereco_var.set("")

    def atualizar_tabela_cliente(self, data=None):
        for row in self.tabela_cliente.get_children():
            self.tabela_cliente.delete(row)

        data = data or self.clientes
        for cliente in data:
            self.tabela_cliente.insert("", "end", values=list(cliente.values()))

    def pesquisar_cliente(self):
        pesquisa = self.pesquisa_var.get().lower()

        clientes_filtrados = [cliente for cliente in self.clientes if pesquisa in str(cliente).lower()]
        self.atualizar_tabela_cliente(clientes_filtrados)

    def abrir_excel(self):

        diretorio_atual = os.getcwd()
        arquivos_excel = [arquivo for arquivo in os.listdir(diretorio_atual) if arquivo.endswith(".xlsx")]

        if arquivos_excel:
            arquivo_excel = arquivos_excel[0]
            caminho_completo = os.path.join(diretorio_atual, arquivo_excel)

            try:
                subprocess.Popen(["start", "", caminho_completo], shell=True)
            except FileNotFoundError:
                print("Erro ao abrir o Excel. Certifique-se de que o Excel está instalado e configure o PATH corretamente.")
        else:
            print("Nenhum arquivo Excel encontrado na pasta.")

if __name__ == "__main__":
    root = tk.Tk()
    app = AplicacaoCliente(root)
    root.grid_rowconfigure(2, weight=1)
    root.grid_columnconfigure(1, weight=1)

    root.mainloop()
