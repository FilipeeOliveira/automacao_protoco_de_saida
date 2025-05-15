from tkinter import messagebox
import tkinter as tk

class AddEquipment:
    def __init__(self, instancia):
        self.instancia = instancia
        
    def execute(self):
        desc = self.instancia.vars['descricao'].get().strip()
        if not desc:
            messagebox.showwarning("Aviso","Digite descrição")
            return
        eq = {
            'descricao': desc,
            'quantidade': self.instancia.vars['quantidade'].get(),
            'unidade': self.instancia.vars['unidade'].get(),
            'estoque': self.instancia.vars['estoque'].get()
        }
        self.instancia.equipamentos.append(eq)
        self.instancia.lista_equipamentos.insert(tk.END,
            f"{eq['quantidade']} {eq['unidade']} - {eq['descricao']} ({eq['estoque']})"
        )
        # limpa campos
        self.instancia.vars['descricao'].set("")
        self.instancia.vars['quantidade'].set(1)
        self.instancia.vars['unidade'].set("UND")
        self.instancia.vars['estoque'].set("SEDE")