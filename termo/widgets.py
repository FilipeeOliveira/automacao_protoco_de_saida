import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry

def get_widget(ctx: dict):
    """
    ctx['root']    : tk.Tk()
    ctx['vars']    : dict com todas as tk.StringVar e tk.IntVar
    ctx['callbacks']: dict com funções add, remove, fill, clear, open, exit
    ctx['equip_listbox']: objeto que tenha método `.lista_equipamentos` (pode ser o controller)
    """
    root = ctx['root']
    vars = ctx['vars']
    cb  = ctx['callbacks']

    main = ttk.Frame(root, padding=10)
    main.pack(fill=tk.BOTH, expand=True)
    main.columnconfigure((0,1), weight=1)

    # --- Header ---
    header = ttk.LabelFrame(main, text="Informações do Termo", padding=10)
    header.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
    header.columnconfigure(1, weight=1)

    ttk.Label(header, text="Nº Controle:").grid(row=0, column=0, sticky="e")
    ttk.Entry(header, textvariable=vars['controle']).grid(row=0, column=1, sticky="ew")
    ttk.Label(header, text="Local de Saída:").grid(row=1, column=0, sticky="e")
    ttk.Entry(header, textvariable=vars['local_saida']).grid(row=1, column=1, sticky="ew")
    ttk.Label(header, text="Local de Destino:").grid(row=2, column=0, sticky="e")
    dest = ttk.Combobox(header, textvariable=vars['local_destino'])
    dest['values'] = ["MACAPÁ","DNIT","HOME - OFFICE","OBRA","VIAGEM","OUTROS"]
    dest.state(['!readonly']); dest.grid(row=2, column=1, sticky="ew")
    ttk.Label(header, text="Motivo:").grid(row=3, column=0, sticky="e")
    ttk.Entry(header, textvariable=vars['motivo']).grid(row=3, column=1, sticky="ew")
    ttk.Label(header, text="Data de Saída:").grid(row=4, column=0, sticky="e")
    DateEntry(header, textvariable=vars['data_saida'], date_pattern='dd-MM-yyyy').grid(row=4, column=1, sticky="w")
    ttk.Label(header, text="Data de Retorno:").grid(row=5, column=0, sticky="e")
    date_ret = DateEntry(header, textvariable=vars['data_retorno'], date_pattern='dd-MM-yyyy')
    date_ret.grid(row=5, column=1, sticky="w")
      # Checkbutton para Sim/Não
    chk = ttk.Checkbutton(
        header,
        text="Possui Retorno?",
        variable=vars['tem_retorno'],
        command=cb['toggle_return']
    )
    chk.grid(row=5, column=0, padx=(10,0))
    
    # Expõe o widget de data no controller para habilitar/desabilitar
    ctx_obj = ctx.get('equip_listbox')
    if ctx_obj:
        ctx_obj.data_retorno_widget = date_ret

    # --- Responsável ---
    resp = ttk.LabelFrame(main, text="Responsável", padding=10)
    resp.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
    resp.columnconfigure(1, weight=1)
    labels = ["Nome:","CPF:","Setor:","Cargo:","Resp. do Setor:"]
    for i,text in enumerate(labels):
        ttk.Label(resp, text=text).grid(row=i, column=0, sticky="e")
    ttk.Entry(resp, textvariable=vars['nome']).grid(row=0, column=1, sticky="ew")
    ttk.Entry(resp, textvariable=vars['cpf']).grid(row=1, column=1, sticky="ew")
    setor = ttk.Combobox(resp, textvariable=vars['setor'])
    setor['values'] = ["ENGENHARIA","FINANCEIRO","PAGAMENTOS","COMPRAS","JURÍDICO","RH",
                       "MARKETING","LOGÍSTICA","PATRIMÔNIO","IP4","OUTROS","SESMET",
                       "DIRETORIA","PROCESSOS","QUALIDADE","TI","DNIT"]
    setor.state(['!readonly']); setor.grid(row=2, column=1, sticky="ew")
    cargo = ttk.Combobox(resp, textvariable=vars['cargo'])
    cargo['values'] = ["ASSISTENTE ADMINISTRATIVO","AUXILIAR ADMINISTRATIVO",
                       "AUXILIAR DE ENGENHARIA","ANALISTA DE TI","COORDENADOR DE PROJETOS",
                       "DIRETOR FINANCEIRO","GERENTE DE COMPRAS","ADVOGADO",
                       "AUXILIAR DE RH","ANALISTA DE MARKETING"]
    cargo.state(['!readonly']); cargo.grid(row=3, column=1, sticky="ew")
    ttk.Entry(resp, textvariable=vars['responsavel_setor']).grid(row=4, column=1, sticky="ew")

    # --- Equipamentos ---
    eqf = ttk.LabelFrame(main, text="Equipamentos", padding=10)
    eqf.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
    eqf.columnconfigure(1, weight=1)
    labels_eq = ["Descrição:","Quantidade:","Unidade:","Estoque:"]
    keys_eq   = ['descricao','quantidade','unidade','estoque']
    for i,(lbl,k) in enumerate(zip(labels_eq,keys_eq)):
        ttk.Label(eqf, text=lbl).grid(row=i, column=0, sticky="e")
        if k=='quantidade':
            ttk.Spinbox(eqf, from_=1, to=9999, textvariable=vars[k]).grid(row=i, column=1, sticky="w")
        else:
            ttk.Entry(eqf, textvariable=vars[k]).grid(row=i, column=1, sticky="ew")
    btnf = ttk.Frame(eqf)
    btnf.grid(row=4, column=0, columnspan=2, pady=5)
    ttk.Button(btnf, text="Adicionar", command=cb['add']).pack(side=tk.LEFT, padx=5)
    ttk.Button(btnf, text="Remover",  command=cb['remove']).pack(side=tk.LEFT, padx=5)
    lst = tk.Listbox(eqf, height=4)
    lst.grid(row=5, column=0, columnspan=2, sticky="ew", pady=5)
    # o controller receberá esta Listbox para atualizá-la
    ctx_obj = ctx.get('equip_listbox')
    if ctx_obj:
        ctx_obj.lista_equipamentos = lst
    ttk.Label(eqf, text="Observação:").grid(row=6, column=0, sticky="e")
    ttk.Entry(eqf, textvariable=vars['observacao']).grid(row=6, column=1, sticky="ew")

    # --- Ações ---
    act = ttk.Frame(main, padding=10)
    act.grid(row=2, column=0, columnspan=2)
    ttk.Button(act, text="Preencher", command=cb['fill']).pack(side=tk.LEFT, padx=5)
    ttk.Button(act, text="Limpar",    command=cb['clear']).pack(side=tk.LEFT, padx=5)
    ttk.Button(act, text="Abrir Pasta", command=cb['open']).pack(side=tk.LEFT, padx=5)
    ttk.Button(act, text="Sair",      command=cb['exit']).pack(side=tk.LEFT, padx=5)

    # Assinatura
    ttk.Label(main, text="Criado por Filipe Oliveira - v1.0",
              font=("Arial",8), foreground="gray").grid(row=3, column=1, sticky="e")
