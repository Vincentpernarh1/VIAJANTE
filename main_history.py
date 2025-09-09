from tkinter import *
from tkinter import ttk
from tkinter import Canvas
from PIL import Image, ImageTk
# Assuming DB.py contains the functions as used in your original code
from DB import completar_informacoes, consolidar_dados, Processar_Demandas
import pandas as pd
import re
import os

def normalizar_codigos(campo):
    if pd.isna(campo):
        return []
    return re.split(r'\s*/\s*', str(campo).strip())

def input_demanda(cod_destino):
    veiculos_dict = {
        'BIG SIDER': 6,
        'BITREM': 7,
        'CARRETA': 4,
        'CARRETA LINE HAUL': 14,
        'CARRETA REBAIXADA': 9,
        'CTNR 20': 15,
        'CTNR 40': 16,
        'FIORINO': 11,
        'RODOTREM': 8,
        'TRUCK 3M': 3,
        'TRUCK 3M ALONGADO': 18,
        'TRUCK 3M PLUS': 13,
        'TRUCK ALONGADO': 17,
        'TRUCK VIAGEM': 2,
        'TRUCK VIAGEM PLUS': 12,
        'VAN': 10,
        'VANDERLEA': 5,
        'VEÍCULO 3/4': 1
    }

    db_fluxos = pd.read_excel('BD_Viajante.xlsx', sheet_name='FLUXOS')
    df = Processar_Demandas(cod_destino)

    # Adiciona a coluna de código de veículo
    cod_veiculos = []
    tipos_saturacao = []
    cod_fornecedor = []

    for _, row in df.iterrows():
        cod_forn = str(row["COD FORNECEDOR"])
        cod_dest = str(row["COD DESTINO"])
        codigo = None
        tipo = None
        cod_ims = None

        for _, linha_fluxo in db_fluxos.iterrows():
            cods_sap = normalizar_codigos(linha_fluxo["COD FORNECEDOR"])
            cods_dest = normalizar_codigos(linha_fluxo["COD DESTINO"])

            if cod_forn in cods_sap and cod_dest in cods_dest:
                nome_veiculo = linha_fluxo["VEICULO PRINCIPAL"]
                codigo = veiculos_dict.get(nome_veiculo, None)
                tipo = linha_fluxo.get("TIPO SATURACAO", None)
                cod_ims = linha_fluxo.get("COD IMS", None)
                break

        cod_veiculos.append(codigo)
        tipos_saturacao.append(tipo)
        cod_fornecedor.append(cod_ims)

    df["VEICULO"] = cod_veiculos
    df["TIPO SATURACAO"] = tipos_saturacao
    df["COD IMS"] = cod_fornecedor

    df = df[['COD FORNECEDOR','COD IMS', 'COD DESTINO', 'DESENHO', 'QTDE', 'VEICULO', 'TIPO SATURACAO']]

    # Salvar em Excel
    df.to_excel("Template.xlsx", index=False)
    print("Template.xlsx")


# Dicionário com {descrição: código}
veiculos_dict = {
    'BIG SIDER': 6,
    'BITREM': 7,
    'CARRETA': 4,
    'CARRETA LINE HAUL': 14,
    'CARRETA REBAIXADA': 9,
    'CTNR 20': 15,
    'CTNR 40': 16,
    'FIORINO': 11,
    'RODOTREM': 8,
    'TRUCK 3M': 3,
    'TRUCK 3M ALONGADO': 18,
    'TRUCK 3M PLUS': 13,
    'TRUCK ALONGADO': 17,
    'TRUCK VIAGEM': 2,
    'TRUCK VIAGEM PLUS': 12,
    'VAN': 10,
    'VANDERLEA': 5,
    'VEÍCULO 3/4': 1
}

janela = Tk()
# Carrega a imagem da carreta
try:
    img = Image.open("carreta.png").resize((140, 100))
    caminhao_img = ImageTk.PhotoImage(img)
except Exception as e:
    print(f"Erro ao carregar imagem da carreta: {e}")
    caminhao_img = None
janela.title("VIAJANTE")
janela.geometry("1400x700")
janela.state('zoomed')

# Frame principal
frame_principal = Frame(janela)
frame_principal.pack(fill=BOTH, expand=True)

# -------------------
# ***** NEW CODE START *****
# 1. Create the loading label but keep it hidden.
# This label will be shown during processing.
loading_label = Label(frame_principal, 
                      text="Processando... Por favor, aguarde.",
                      font=("Arial", 18, "bold"), 
                      bg="white", 
                      fg="#007acc",
                      relief="solid", 
                      borderwidth=2)
# We will use .place() to show it and .place_forget() to hide it.
# ***** NEW CODE END *****
# -------------------

# Parte superior (linha única com seleção, canvas e resumo lado a lado)
frame_top = Frame(frame_principal)
frame_top.pack(fill=X, padx=10, pady=10)

# ----- Coluna 1: Seleção de veículos -----
frame_selecao = Frame(frame_top)
frame_selecao.grid(row=0, column=0, sticky='nw', padx=10)

Label(frame_selecao, text="Selecione o tipo de veículo:").grid(row=0, column=0, columnspan=3, pady=(0, 5))

veiculo_var = StringVar(value='')

frame_veiculos = Frame(frame_selecao)
frame_veiculos.grid(row=1, column=0, columnspan=3, sticky='w')

colunas = 3
for i, (nome, cod) in enumerate(sorted(veiculos_dict.items())):
    rb = Radiobutton(frame_veiculos, text=nome, variable=veiculo_var, value=str(cod))
    rb.grid(row=i // colunas, column=i % colunas, sticky='w', padx=5, pady=2)

label_veiculo = Label(frame_selecao, text="")
label_veiculo.grid(row=2, column=0, columnspan=3, pady=5)

# Variável para armazenar o modo de atualização
modo_manual = BooleanVar(value=False)

# Checkbutton para alternar o modo
check_manual = Checkbutton(
    frame_selecao,
    text="Usar veículo escolhido para todos",
    variable=modo_manual
)
check_manual.grid(row=3, column=0, columnspan=3, sticky='w')

Button(frame_selecao, text="Atualizar", command=lambda: atualizar()).grid(row=4, column=1, pady=5)

# ----- Coluna 2: Canvas de caminhões -----
frame_caminhoes = Frame(frame_top)
frame_caminhoes.grid(row=0, column=1, padx=20)

canvas_caminhoes = Canvas(frame_caminhoes, width=400, height=300)
canvas_caminhoes.pack()

# ----- Coluna 3: Tabela resumo -----
frame_resumo = Frame(frame_top)
frame_resumo.grid(row=0, column=2, sticky='ne', padx=10)

tree_resumo = ttk.Treeview(frame_resumo, columns=("Info", "Valor"), show="headings", height=10)
tree_resumo.heading("Info", text="Info")
tree_resumo.heading("Valor", text="Valor")
tree_resumo.column("Info", width=120, anchor='center')
tree_resumo.column("Valor", width=100, anchor='center')
tree_resumo.pack()

# Preencher exemplo
tree_resumo.insert("", END, values=("Ocupação Total", ""))
tree_resumo.insert("", END, values=("Qtd Veículos", ""))
tree_resumo.insert("", END, values=("Volume Total", ""))
tree_resumo.insert("", END, values=("Peso Total", ""))
#tree_resumo.insert("", END, values=("Peso Máximo", ""))
tree_resumo.insert("", END, values=("Embalagens", ""))

# -------------------
# Parte inferior: Tabela principal
frame_bottom = Frame(frame_principal)
frame_bottom.pack(fill=BOTH, expand=True, padx=10, pady=(0, 10))

scroll_y = Scrollbar(frame_bottom, orient=VERTICAL)
scroll_y.pack(side=RIGHT, fill=Y)

tree = ttk.Treeview(frame_bottom, yscrollcommand=scroll_y.set)
tree.pack(fill=BOTH, expand=True)
scroll_y.config(command=tree.yview)

# Estilo do cabeçalho da tabela
style = ttk.Style()
style.configure("Treeview.Heading", background="#007acc", foreground="blue", font=("Arial", 10, "bold"))


# -------------------
# ***** MODIFIED FUNCTION START *****
# 2. Modify the 'atualizar' function to show and hide the loading label.
def atualizar():
    # Show the loading label in the center of the main frame
    loading_label.place(relx=0.5, rely=0.5, anchor='center')
    loading_label.lift()  # Ensure the label is on top of other widgets
    janela.update_idletasks()  # Force the UI to refresh and show the label immediately

    try:
        # --- This is your original code block ---
        cod = veiculo_var.get()
        label_veiculo.config(text=f"Código selecionado: {cod}")
        if cod:
            input_demanda(1080)
            completar_informacoes(tree, int(cod), tree_resumo, canvas_caminhoes, caminhao_img, usar_manual=modo_manual.get())
            consolidar_dados()
        # ----------------------------------------
    except Exception as e:
        # It's good practice to handle potential errors
        print(f"Ocorreu um erro durante a atualização: {e}")
    finally:
        # This block will run whether an error occurred or not,
        # ensuring the loading label is always hidden afterward.
        loading_label.place_forget()
# ***** MODIFIED FUNCTION END *****
# -------------------

janela.mainloop()













