class OnToggleReturn:
    def __init__(self, instancia):
        self.instancia = instancia
        
    def execute(self):
        """ Habilita/desabilita (e limpa) o campo de data_retorno na UI """
        entry = getattr(self.instancia, 'data_retorno_widget', None)
        if not self.instancia.vars['tem_retorno'].get():
            # desliga e limpa
            self.instancia.vars['data_retorno'].set("")
            entry.config(state='disabled')
        else:
            entry.config(state='normal')