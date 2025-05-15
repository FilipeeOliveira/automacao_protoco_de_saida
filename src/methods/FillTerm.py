from tkinter import messagebox
from src.utils.generate import fill_term

class FillTerm:
    def __init__(self, instancia):
        self.instancia = instancia
        
    def execute(self):
        if not self.instancia.vars['tem_retorno'].get():
            self.instancia.vars['data_retorno'].set("")
        if not self.instancia.equipamentos:
            messagebox.showwarning("Aviso","Adicione ao menos um equipamento")
            return
        data = {k: v.get() for k,v in self.instancia.vars.items()}
        try:
            path = fill_term(data, self.instancia.equipamentos, self.instancia.output_dir)
            messagebox.showinfo("Sucesso", f"Termo salvo em:\n{path}")
            self.instancia.limpar_campos()
        except Exception as e:
            messagebox.showerror("Erro", str(e))