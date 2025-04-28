import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry  
from openpyxl import load_workbook
from datetime import datetime
import os

class TermoEntregaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Preenchimento de Termo de Entrega")
        
        # Variáveis para armazenar os dados
        self.controle_var = tk.StringVar(value=self.get_next_control())
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

    def load_last_control(self):
        """Lê o último número de controle de arquivo ou retorna None"""
        path = 'last_control.txt'
        if os.path.exists(path):
            with open(path, 'r') as f:
                return f.read().strip()
        return None

    def save_last_control(self, control):
        """Salva o último número de controle em arquivo"""
        with open('last_control.txt', 'w') as f:
            f.write(control)

    def get_next_control(self):
        """Gera o próximo número de controle no formato NNNN/YYYY"""
        last = self.load_last_control()
        year = datetime.now().year
        if last and '/' in last:
            num_str, y_str = last.split('/')
            try:
                num = int(num_str)
                y = int(y_str)
            except ValueError:
                num, y = 0, year
            if y == year:
                next_num = num + 1
            else:
                next_num = 1
        else:
            next_num = 1
        return f"{next_num:04d}/{year}"

    def create_widgets(self):
        # Frame principal com duas colunas
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Seção de cabeçalho (coluna 0)
        header_frame = ttk.LabelFrame(main_frame, text="Informações do Termo", padding=10)
        header_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        header_frame.columnconfigure(1, weight=1)
        
        # Campos do termo
        labels = ["Nº de Controle:", "Local de Saída:", "Local de Destino:", "Motivo:"]
        vars = [self.controle_var, self.local_saida_var, self.local_destino_var, self.motivo_var]
        for i, (lbl, var) in enumerate(zip(labels, vars)):
            ttk.Label(header_frame, text=lbl).grid(row=i, column=0, sticky="e", pady=2)
            ttk.Entry(header_frame, textvariable=var).grid(row=i, column=1, sticky="ew", pady=2)
        
        # DatePickers
        ttk.Label(header_frame, text="Data de Saída:").grid(row=4, column=0, sticky="e", pady=2)
        DateEntry(header_frame, textvariable=self.data_saida_var, date_pattern='yyyy-MM-dd').grid(row=4, column=1, sticky="w", pady=2)
        ttk.Label(header_frame, text="Data de Retorno:").grid(row=5, column=0, sticky="e", pady=2)
        DateEntry(header_frame, textvariable=self.data_retorno_var, date_pattern='yyyy-MM-dd').grid(row=5, column=1, sticky="w", pady=2)
        
        # Seção do Responsável (coluna 1)
        responsavel_frame = ttk.LabelFrame(main_frame, text="Responsável", padding=10)
        responsavel_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        responsavel_frame.columnconfigure(1, weight=1)
        
        # Campos do responsável
        labels_r = ["Nome:", "CPF:", "Setor:", "Cargo:", "Responsável do Setor:"]
        vars_r = [self.nome_var, self.cpf_var, self.setor_var, self.cargo_var, self.responsavel_setor_var]
        for i, (lbl, var) in enumerate(zip(labels_r, vars_r)):
            ttk.Label(responsavel_frame, text=lbl).grid(row=i, column=0, sticky="e", pady=2)
            ttk.Entry(responsavel_frame, textvariable=var).grid(row=i, column=1, sticky="ew", pady=2)
        
        # Seção de Equipamento (abaixo, ocupa duas colunas)
        equipamento_frame = ttk.LabelFrame(main_frame, text="Detalhamento do Equipamento", padding=10)
        equipamento_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        equipamento_frame.columnconfigure(1, weight=1)
        
        # Campos de equipamento
        eq_labels = ["Descrição:", "Quantidade:", "Unidade:", "Estoque:"]
        eq_vars = [self.descricao_var, self.quantidade_var, self.unidade_var, self.estoque_var]
        for i, (lbl, var) in enumerate(zip(eq_labels, eq_vars)):
            ttk.Label(equipamento_frame, text=lbl).grid(row=i, column=0, sticky="e", pady=2)
            ttk.Entry(equipamento_frame, textvariable=var, width=50 if lbl=="Descrição:" else None).grid(row=i, column=1, sticky="ew", pady=2)
        
        # Botões de adicionar/remover
        btn_frame = ttk.Frame(equipamento_frame)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=5)
        ttk.Button(btn_frame, text="Adicionar Equipamento", command=self.adicionar_equipamento).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Remover Selecionado", command=self.remover_equipamento).pack(side=tk.LEFT, padx=5)
        
        # Lista de equipamentos
        self.lista_equipamentos = tk.Listbox(equipamento_frame, height=4)
        self.lista_equipamentos.grid(row=5, column=0, columnspan=2, sticky="ew", pady=5)
        
        # Observação
        ttk.Label(equipamento_frame, text="Observação:").grid(row=6, column=0, sticky="e", pady=2)
        ttk.Entry(equipamento_frame, textvariable=self.observacao_var, width=50).grid(row=6, column=1, sticky="ew", pady=2)
        
        # Botões principais
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=2, column=0, columnspan=2, pady=10)
        ttk.Button(action_frame, text="Preencher Termo", command=self.preencher_termo).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Limpar", command=self.limpar_campos).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Sair", command=self.root.quit).pack(side=tk.LEFT, padx=5)
        
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