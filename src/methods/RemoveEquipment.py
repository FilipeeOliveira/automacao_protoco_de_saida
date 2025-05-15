class RemoveEquipment:
    def __init__(self, instancia):
        self.instancia = instancia
        
    def execute(self):
        sel = self.instancia.lista_equipamentos.curselection()
        if not sel: return
        idx = sel[0]
        self.instancia.equipamentos.pop(idx)
        self.instancia.lista_equipamentos.delete(idx)