import tkinter as tk
from tkinter import messagebox
from src.widgets import get_widget
from src.generate import get_next_control, fill_term

class TermoEntregaApp:
    def __init__(self, root, output_dir="termos_salvos"):
        self.root = root
        self.output_dir = output_dir

        # Variáveis tk
        self.vars = {
            'controle'           : tk.StringVar(value=get_next_control()),
            'local_saida'        : tk.StringVar(value="ANTONELLY SEDE - ESTOQUE T.I"),
            'local_destino'      : tk.StringVar(),
            'motivo'             : tk.StringVar(),
            'data_saida'         : tk.StringVar(),
            'data_retorno'       : tk.StringVar(),
            'tem_retorno'        : tk.BooleanVar(value=True),
            'nome'               : tk.StringVar(),
            'cpf'                : tk.StringVar(),
            'setor'              : tk.StringVar(),
            'cargo'              : tk.StringVar(),
            'responsavel_setor'  : tk.StringVar(),
            'descricao'          : tk.StringVar(),
            'quantidade'         : tk.IntVar(value=1),
            'unidade'            : tk.StringVar(value="UND"),
            'estoque'            : tk.StringVar(value="SEDE"),
            'observacao'         : tk.StringVar(),
        }
        self.equipamentos = []

        # Callbacks
        self.callbacks = {
            'add'    : self.adicionar_equipamento,
            'remove' : self.remover_equipamento,
            'fill'   : self.preencher_termo,
            'clear'  : self.limpar_campos,
            'open'   : self.open_folder,
            'exit'   : self.root.quit,
            'toggle_return': self.on_toggle_return
        }

        # Constrói UI
        get_widget({
            'root': self.root,
            'vars': self.vars,
            'callbacks': self.callbacks,
            'equip_listbox': self
        })
        
    def on_toggle_return(self):
        """ Habilita/desabilita (e limpa) o campo de data_retorno na UI """
        entry = getattr(self, 'data_retorno_widget', None)
        if not self.vars['tem_retorno'].get():
            # desliga e limpa
            self.vars['data_retorno'].set("")
            entry.config(state='disabled')
        else:
            entry.config(state='normal')

    def adicionar_equipamento(self):
        desc = self.vars['descricao'].get().strip()
        if not desc:
            messagebox.showwarning("Aviso","Digite descrição")
            return
        eq = {
            'descricao': desc,
            'quantidade': self.vars['quantidade'].get(),
            'unidade': self.vars['unidade'].get(),
            'estoque': self.vars['estoque'].get()
        }
        self.equipamentos.append(eq)
        self.lista_equipamentos.insert(tk.END,
            f"{eq['quantidade']} {eq['unidade']} - {eq['descricao']} ({eq['estoque']})"
        )
        # limpa campos
        self.vars['descricao'].set("")
        self.vars['quantidade'].set(1)
        self.vars['unidade'].set("UND")
        self.vars['estoque'].set("SEDE")

    def remover_equipamento(self):
        sel = self.lista_equipamentos.curselection()
        if not sel: return
        idx = sel[0]
        self.equipamentos.pop(idx)
        self.lista_equipamentos.delete(idx)

    def preencher_termo(self):
        if not self.vars['tem_retorno'].get():
            self.vars['data_retorno'].set("")
        if not self.equipamentos:
            messagebox.showwarning("Aviso","Adicione ao menos um equipamento")
            return
        data = {k: v.get() for k,v in self.vars.items()}
        try:
            path = fill_term(data, self.equipamentos, self.output_dir)
            messagebox.showinfo("Sucesso", f"Termo salvo em:\n{path}")
            self.limpar_campos()
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def limpar_campos(self):
        for k, v in self.vars.items():
            if isinstance(v, tk.IntVar):
                v.set(1)
            elif isinstance(v, tk.BooleanVar):
                v.set(True) 
            else:
                v.set("")
        self.vars['controle'].set(get_next_control())
        self.equipamentos.clear()
        self.lista_equipamentos.delete(0, tk.END)
        self.on_toggle_return()

    def open_folder(self):
        import os, sys, subprocess
        p = os.path.abspath(self.output_dir)
        try:
            if sys.platform.startswith('darwin'):
                subprocess.call(['open', p])
            elif sys.platform.startswith('linux'):
                subprocess.call(['xdg-open', p])
            else:
                os.startfile(p)
        except Exception as e:
            messagebox.showerror("Erro", str(e))
