import tkinter as tk
from src.utils.generate import get_next_control

class ClearTerm:
    def __init__(self, instancia):
        self.instancia = instancia
        
    def execute(self):
        for k, v in self.instancia.vars.items():
            if isinstance(v, tk.IntVar):
                v.set(1)
            elif isinstance(v, tk.BooleanVar):
                v.set(True) 
            else:
                v.set("")
        self.instancia.vars['controle'].set(get_next_control())
        self.instancia.equipamentos.clear()
        self.instancia.lista_equipamentos.delete(0, tk.END)
        self.instancia.on_toggle_return()  
