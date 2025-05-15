import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry

class Widgets:
    def __init__(self, instancia):
        self.instancia = instancia

    def execute(self):
        """
        root        : tk.Tk()
        vars        : dict com todas as tk.StringVar e tk.IntVar
        callbacks   : dict com funções add, remove, fill, clear, open, exit
        equip_listbox: objeto que tenha método `.lista_equipamentos` (pode ser o controller)
        """

        main = ttk.Frame(self.instancia.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)
        main.columnconfigure((0,1), weight=1)

        # --- Header ---
        header = ttk.LabelFrame(main, text="Informações do Termo", padding=10)
        header.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        header.columnconfigure(1, weight=1)

        ttk.Label(header, text="Nº Controle:").grid(row=0, column=0, sticky="e")
        ttk.Entry(header, textvariable=self.instancia.vars['controle']).grid(row=0, column=1, sticky="ew")
        ttk.Label(header, text="Local de Saída:").grid(row=1, column=0, sticky="e")
        ttk.Entry(header, textvariable=self.instancia.vars['local_saida']).grid(row=1, column=1, sticky="ew")
        ttk.Label(header, text="Local de Destino:").grid(row=2, column=0, sticky="e")
        dest = ttk.Combobox(header, textvariable=self.instancia.vars['local_destino'])
        dest['values'] = ["MACAPÁ","DNIT","HOME - OFFICE","OBRA","VIAGEM","OUTROS"]
        dest.state(['!readonly']); dest.grid(row=2, column=1, sticky="ew")
        ttk.Label(header, text="Motivo:").grid(row=3, column=0, sticky="e")
        ttk.Entry(header, textvariable=self.instancia.vars['motivo']).grid(row=3, column=1, sticky="ew")
        ttk.Label(header, text="Data de Saída:").grid(row=4, column=0, sticky="e")
        DateEntry(header, textvariable=self.instancia.vars['data_saida'], date_pattern='dd-MM-yyyy').grid(row=4, column=1, sticky="w")
        frame_retorno = ttk.Frame(header)
        frame_retorno.grid(row=5, column=0, sticky="e", padx=(0, 5))
        ttk.Label(frame_retorno, text="Possui Retorno?").pack(side=tk.LEFT)
        chk = ttk.Checkbutton(
            frame_retorno,
            variable=self.instancia.vars['tem_retorno'],
            command=self.instancia.on_toggle_return
        )
        chk.pack(side=tk.LEFT, padx=(4, 0))
        date_ret = DateEntry(header, textvariable=self.instancia.vars['data_retorno'], date_pattern='dd-MM-yyyy')
        date_ret.grid(row=5, column=1, sticky="w")
        
        # ctx_obj = self.instancia
        # if ctx_obj:
        #     ctx_obj.data_retorno_widget = date_ret
        self.instancia.data_retorno_widget = date_ret

        # --- Responsável ---
        resp = ttk.LabelFrame(main, text="Responsável", padding=10)
        resp.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        resp.columnconfigure(1, weight=1)
        labels = ["Nome:","CPF:","Setor:","Cargo:","Resp. do Setor:"]
        for i,text in enumerate(labels):
            ttk.Label(resp, text=text).grid(row=i, column=0, sticky="e")
        ttk.Entry(resp, textvariable=self.instancia.vars['nome']).grid(row=0, column=1, sticky="ew")
        ttk.Entry(resp, textvariable=self.instancia.vars['cpf']).grid(row=1, column=1, sticky="ew")
        setor = ttk.Combobox(resp, textvariable=self.instancia.vars['setor'])
        setor['values'] = ["ENGENHARIA","FINANCEIRO","PAGAMENTOS","COMPRAS","JURÍDICO","RH",
                        "MARKETING","LOGÍSTICA","PATRIMÔNIO","IP4","OUTROS","SESMET",
                        "DIRETORIA","PROCESSOS","QUALIDADE","TI","DNIT"]
        setor.state(['!readonly']); setor.grid(row=2, column=1, sticky="ew")
        cargo = ttk.Combobox(resp, textvariable=self.instancia.vars['cargo'])
        cargo['values'] = ["ASSISTENTE ADMINISTRATIVO","AUXILIAR ADMINISTRATIVO",
                        "AUXILIAR DE ENGENHARIA","ANALISTA DE TI","COORDENADOR DE PROJETOS",
                        "DIRETOR FINANCEIRO","GERENTE DE COMPRAS","ADVOGADO",
                        "AUXILIAR DE RH","ANALISTA DE MARKETING"]
        cargo.state(['!readonly']); cargo.grid(row=3, column=1, sticky="ew")
        ttk.Entry(resp, textvariable=self.instancia.vars['responsavel_setor']).grid(row=4, column=1, sticky="ew")

        # --- Equipamentos ---
        eqf = ttk.LabelFrame(main, text="Equipamentos", padding=10)
        eqf.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        eqf.columnconfigure(1, weight=1)
        labels_eq = ["Descrição:","Quantidade:","Unidade:","Estoque:"]
        keys_eq   = ['descricao','quantidade','unidade','estoque']
        for i,(lbl,k) in enumerate(zip(labels_eq,keys_eq)):
            ttk.Label(eqf, text=lbl).grid(row=i, column=0, sticky="e")
            if k=='quantidade':
                ttk.Spinbox(eqf, from_=1, to=9999, textvariable=self.instancia.vars[k]).grid(row=i, column=1, sticky="w")
            else:
                ttk.Entry(eqf, textvariable=self.instancia.vars[k]).grid(row=i, column=1, sticky="ew")
        
        btnf = ttk.Frame(eqf)
        btnf.grid(row=4, column=0, columnspan=2, pady=5)
        ttk.Button(btnf, text="Adicionar", command=self.instancia.adicionar_equipamento).pack(side=tk.LEFT, padx=5)
        ttk.Button(btnf, text="Remover",  command=self.instancia.remover_equipamento).pack(side=tk.LEFT, padx=5)
        lst = tk.Listbox(eqf, height=4)
        lst.grid(row=5, column=0, columnspan=2, sticky="ew", pady=5)
        
        # o controller receberá esta Listbox para atualizá-la
        self.instancia.lista_equipamentos = lst
        
        ttk.Label(eqf, text="Observação:").grid(row=6, column=0, sticky="e")
        ttk.Entry(eqf, textvariable=self.instancia.vars['observacao']).grid(row=6, column=1, sticky="ew")

        # --- Ações ---
        act = ttk.Frame(main, padding=10)
        act.grid(row=2, column=0, columnspan=2)
        ttk.Button(act, text="Preencher", command=self.instancia.preencher_termo).pack(side=tk.LEFT, padx=5)
        ttk.Button(act, text="Limpar",    command=self.instancia.limpar_campos).pack(side=tk.LEFT, padx=5)
        ttk.Button(act, text="Abrir Pasta", command=self.instancia.open_folder).pack(side=tk.LEFT, padx=5)
        ttk.Button(act, text="Sair",  command=self.instancia.root.quit).pack(side=tk.LEFT, padx=5)

        # Assinatura
        ttk.Label(main, text="Criado por Filipe Oliveira - v1.0",
                font=("Arial",8), foreground="gray").grid(row=3, column=1, sticky="e")
