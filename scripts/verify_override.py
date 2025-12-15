import sys, os
script_dir = os.path.dirname(os.path.abspath(__file__))
proj_root = os.path.abspath(os.path.join(script_dir, '..'))
sys.path.insert(0, proj_root)
from DB import completar_informacoes, consolidar_dados
from tkinter import ttk, Tk, Canvas

root = Tk()
root.withdraw()

tree = ttk.Treeview(root)
tree_resumo = ttk.Treeview(root)
canvas = Canvas(root, width=1, height=1)

# Use vehicle code 4 as example and force manual
completar_informacoes(tree, 4, tree_resumo, canvas, None, usar_manual=True)
print('completar_informacoes completed')

# Run consolidation to generate Volume_por_rota.xlsx
consolidar_dados()
print('consolidar_dados completed')

# Inspect template file
import pandas as pd
if __name__ == '__main__':
    t = pd.read_excel('Template.xlsx')
    print('Unique VEICULO in Template.xlsx:', t['VEICULO'].dropna().unique())
    v = pd.read_excel('Volume_por_rota.xlsx')
    print('Unique VEÍCULO in Volume_por_rota.xlsx:', v['VEÍCULO'].dropna().unique())