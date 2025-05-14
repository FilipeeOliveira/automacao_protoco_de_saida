from tkinter import messagebox

class OpenFolder:
    def __init__(self, instancia):
        self.instancia = instancia
        
    def execute(self):
        import os, sys, subprocess
        p = os.path.abspath(self.instancia.output_dir)
        try:
            if sys.platform.startswith('darwin'):
                subprocess.call(['open', p])
            elif sys.platform.startswith('linux'):
                subprocess.call(['xdg-open', p])
            else:
                os.startfile(p)
        except Exception as e:
            messagebox.showerror("Erro", str(e))