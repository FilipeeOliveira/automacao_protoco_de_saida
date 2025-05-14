import tkinter as tk
from src.utils.generate import get_next_control
from src.methods.Widgets import Widgets
from src.methods.AddEquipment import AddEquipment
from src.methods.RemoveEquipment import RemoveEquipment
from src.methods.FillTerm import FillTerm
from src.methods.ClearTerm import ClearTerm
from src.methods.OpenFolder import OpenFolder
from src.methods.OnToggleReturn import OnToggleReturn
class TermoEntregaApp:
    def __init__(self, root, output_dir="termos_salvos"):
        self.root = root
        self.output_dir = output_dir
        
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
        
        Widgets(self).execute()
        
        
    def on_toggle_return(self):
        OnToggleReturn(self).execute()
        
    def adicionar_equipamento(self):
        AddEquipment(self).execute()

    def remover_equipamento(self):
        RemoveEquipment(self).execute()

    def preencher_termo(self):
        FillTerm(self).execute()

    def limpar_campos(self):
        ClearTerm(self).execute()

    def open_folder(self):
        OpenFolder(self).execute()
