import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry  
from openpyxl import load_workbook
from datetime import datetime

class TermoEntregaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Preenchimento de Termo de Entrega")
        
        # Variáveis para armazenar os dados
        self.controle_var = tk.StringVar()
        self.local_saida_var = tk.StringVar(value="ANTONELLY SEDE - ESTOQUE T.I")
        self.local_destino_var = tk.StringVar()
        self.motivo_var = tk.StringVar()
        self.data_saida_var = tk.StringVar()
        self.data_retorno_var = tk.StringVar()
        self.nome_var = tk.StringVar()
        self.cpf_var = tk.StringVar()
        self.setor_var = tk.StringVar()
        self.cargo_var = tk.StringVar()
        self.responsavel_setor_var = tk.StringVar()
        self.descricao_var = tk.StringVar()
        self.quantidade_var = tk.StringVar()
        self.unidade_var = tk.StringVar()
        self.estoque_var = tk.StringVar(value="SEDE")
        self.observacao_var = tk.StringVar()
        
        
        # Lista para armazenar múltiplos equipamentos
        self.equipamentos = []
        
        # Criar a interface
        self.create_widgets()
        
    def create_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Seção de cabeçalho
        header_frame = ttk.LabelFrame(main_frame, text="Informações do Termo", padding="10")
        header_frame.grid(row=0, column=0, sticky="ew", pady=5)
        
        ttk.Label(header_frame, text="Nº de Controle:").grid(row=0, column=0, sticky="e")
        ttk.Entry(header_frame, textvariable=self.controle_var).grid(row=0, column=1)
        
        ttk.Label(header_frame, text="Local de Saída:").grid(row=1, column=0, sticky="e")
        ttk.Entry(header_frame, textvariable=self.local_saida_var).grid(row=1, column=1)
        
        ttk.Label(header_frame, text="Local de Destino:").grid(row=2, column=0, sticky="e")
        ttk.Entry(header_frame, textvariable=self.local_destino_var).grid(row=2, column=1)
        
        ttk.Label(header_frame, text="Motivo:").grid(row=3, column=0, sticky="e")
        ttk.Entry(header_frame, textvariable=self.motivo_var).grid(row=3, column=1)
        
        ttk.Label(header_frame, text="Data de Saída:").grid(row=4, column=0, sticky="e")
        ttk.Entry(header_frame, textvariable=self.data_saida_var).grid(row=4, column=1)
        
        ttk.Label(header_frame, text="Data de Retorno:").grid(row=5, column=0, sticky="e")
        ttk.Entry(header_frame, textvariable=self.data_retorno_var).grid(row=5, column=1)
        
        # Seção do Responsável
        responsavel_frame = ttk.LabelFrame(main_frame, text="Responsável", padding="10")
        responsavel_frame.grid(row=1, column=0, sticky="ew", pady=5)
        
        ttk.Label(responsavel_frame, text="Nome:").grid(row=0, column=0, sticky="e")
        ttk.Entry(responsavel_frame, textvariable=self.nome_var).grid(row=0, column=1)
        
        ttk.Label(responsavel_frame, text="CPF:").grid(row=1, column=0, sticky="e")
        ttk.Entry(responsavel_frame, textvariable=self.cpf_var).grid(row=1, column=1)
        
        ttk.Label(responsavel_frame, text="Setor:").grid(row=2, column=0, sticky="e")
        ttk.Entry(responsavel_frame, textvariable=self.setor_var).grid(row=2, column=1)
        
        ttk.Label(responsavel_frame, text="Cargo:").grid(row=3, column=0, sticky="e")
        ttk.Entry(responsavel_frame, textvariable=self.cargo_var).grid(row=3, column=1)
        
        ttk.Label(responsavel_frame, text="Responsável do Setor:").grid(row=4, column=0, sticky="e")
        ttk.Entry(responsavel_frame, textvariable=self.responsavel_setor_var).grid(row=4, column=1)
        
        # Seção do Equipamento (agora com botão para adicionar mais)
        equipamento_frame = ttk.LabelFrame(main_frame, text="Detalhamento do Equipamento", padding="10")
        equipamento_frame.grid(row=2, column=0, sticky="ew", pady=5)
        
        ttk.Label(equipamento_frame, text="Descrição:").grid(row=0, column=0, sticky="e")
        ttk.Entry(equipamento_frame, textvariable=self.descricao_var, width=50).grid(row=0, column=1)
        
        ttk.Label(equipamento_frame, text="Quantidade:").grid(row=1, column=0, sticky="e")
        ttk.Entry(equipamento_frame, textvariable=self.quantidade_var).grid(row=1, column=1)
        
        ttk.Label(equipamento_frame, text="Unidade:").grid(row=2, column=0, sticky="e")
        ttk.Entry(equipamento_frame, textvariable=self.unidade_var).grid(row=2, column=1)
        
        ttk.Label(equipamento_frame, text="Estoque:").grid(row=3, column=0, sticky="e")
        ttk.Entry(equipamento_frame, textvariable=self.estoque_var).grid(row=3, column=1)
        
        # Botão para adicionar equipamento à lista
        ttk.Button(equipamento_frame, text="Adicionar Equipamento", command=self.adicionar_equipamento).grid(row=4, column=0, columnspan=2, pady=5)
        
        # Lista de equipamentos adicionados (visualização)
        self.lista_equipamentos = tk.Listbox(equipamento_frame, height=4, width=60)
        self.lista_equipamentos.grid(row=5, column=0, columnspan=2, pady=5)
        
        # Botão para remover equipamento selecionado
        ttk.Button(equipamento_frame, text="Remover Selecionado", command=self.remover_equipamento).grid(row=6, column=0, columnspan=2, pady=5)
        
        # Observação
        ttk.Label(equipamento_frame, text="Observação:").grid(row=7, column=0, sticky="e")
        ttk.Entry(equipamento_frame, textvariable=self.observacao_var, width=50).grid(row=7, column=1)
        
        # Botões principais
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, pady=10)
        
        ttk.Button(button_frame, text="Preencher Termo", command=self.preencher_termo).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Limpar", command=self.limpar_campos).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Sair", command=self.root.quit).pack(side=tk.LEFT, padx=5)
        
    def adicionar_equipamento(self):
        """Adiciona um equipamento à lista"""
        descricao = self.descricao_var.get()
        quantidade = self.quantidade_var.get()
        unidade = self.unidade_var.get()
        estoque = self.estoque_var.get()
        
        if not descricao:
            messagebox.showwarning("Aviso", "Digite uma descrição para o equipamento!")
            return
        
        equipamento = {
            "descricao": descricao,
            "quantidade": quantidade,
            "unidade": unidade,
            "estoque": estoque
        }
        
        self.equipamentos.append(equipamento)
        self.lista_equipamentos.insert(tk.END, f"{quantidade} {unidade} - {descricao} ({estoque})")
        
        # Limpa os campos para um novo equipamento
        self.descricao_var.set("")
        self.quantidade_var.set("1")
        self.unidade_var.set("UND")
        self.estoque_var.set("SEDE")
        
    def remover_equipamento(self):
        """Remove o equipamento selecionado da lista"""
        selecionado = self.lista_equipamentos.curselection()
        if selecionado:
            self.equipamentos.pop(selecionado[0])
            self.lista_equipamentos.delete(selecionado[0])
        
    def limpar_campos(self):
        """Limpa todos os campos"""
        self.controle_var.set("")
        self.nome_var.set("")
        self.cpf_var.set("")
        self.descricao_var.set("")
        self.quantidade_var.set("1")
        self.unidade_var.set("UND")
        self.estoque_var.set("SEDE")
        self.observacao_var.set("")
        self.data_saida_var.set("")
        self.data_retorno_var.set("")
        self.equipamentos = []
        self.lista_equipamentos.delete(0, tk.END)
        
    def preencher_termo(self):
        """Preenche o termo no Excel com todos os equipamentos"""
        if not self.equipamentos:
            messagebox.showwarning("Aviso", "Adicione pelo menos um equipamento!")
            return
            
        try:
            # Carrega o template do Excel
            template_path = "TERMO_DE_ENTREGA_DE_EQUIPAMENTO_151.xlsx"
            wb = load_workbook(template_path)
            ws = wb["ORDEM DE RETIRADA DE ESTOQUE"]
            
            # Preenche os dados gerais
            ws['G2'] = self.controle_var.get()
            ws['C3'] = f"{self.local_saida_var.get()}"
            ws['C4'] = f"{self.local_destino_var.get()}"
            ws['F4'] = self.data_saida_var.get()
            ws['C5'] = f"{self.motivo_var.get()}"
            ws['F5'] = self.data_retorno_var.get()
            ws['B7'] = f"{self.nome_var.get()}"
            ws['E7'] = f"{self.cpf_var.get()}"
            ws['B8'] = f"{self.setor_var.get()}"
            ws['E8'] = f"{self.cargo_var.get()}"
            ws['B9'] = f"{self.responsavel_setor_var.get()}"
            ws['A23'] = f"{self.observacao_var.get()}"
            
            # Preenche os equipamentos (a partir da linha 12)
            linha = 12
            for equipamento in self.equipamentos:
                ws[f'A{linha}'] = equipamento["descricao"]
                ws[f'E{linha}'] = equipamento["quantidade"]
                ws[f'F{linha}'] = equipamento["unidade"]
                ws[f'G{linha}'] = equipamento["estoque"]
                linha += 1
            
            # Data atual para as assinaturas
            data_atual = datetime.now().strftime("%Y-%m-%d")
            ws['A32'] = f"DATA: {data_atual}"
            ws['C32'] = f"DATA: {data_atual}"
            ws['D32'] = f"DATA: {data_atual}"
            ws['F32'] = f"DATA: {data_atual}"
            
            # Salva o arquivo com um novo nome
            nome_arquivo = f"Termo_{self.nome_var.get()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            wb.save(nome_arquivo)
            
            messagebox.showinfo("Sucesso", f"Termo preenchido e salvo como:\n{nome_arquivo}")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao preencher o termo:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = TermoEntregaApp(root)
    root.mainloop()