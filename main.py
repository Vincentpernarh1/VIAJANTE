from tkinter import *
from tkinter import ttk
from tkinter import Canvas
from PIL import Image, ImageTk
# Assuming DB.py contains the functions as used in your original code
from DB import completar_informacoes, consolidar_dados, Processar_Demandas
import pandas as pd
import re
import os
import sys
import threading

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    if hasattr(sys, '_MEIPASS'):
        # When running from the .exe
        return os.path.join(sys._MEIPASS, relative_path)
    else:
        # When running from source
        return os.path.join(os.path.abspath("."), relative_path)

import warnings # <-- 1. Import the library

# 2. Add these lines to ignore the specific warnings from the Excel reader
warnings.filterwarnings(
    "ignore",
    category=UserWarning,
    message=".*OLE2.*"
)
warnings.filterwarnings(
    "ignore",
    category=UserWarning,
    message=".*CODEPAGE.*"
)
warnings.filterwarnings(
    "ignore",
    message="^WARNING .*" # Hides the file size warnings which don't have a category
)

caminho_base = os.getcwd()
# --- START: Global variables for filtering ---
# Stores the complete, unfiltered data from the Treeview
original_tree_data = []
# A dictionary to hold the filter Combobox widgets
filter_widgets = {}
# --- END: Global variables ---


# ------------------- carregar veículos dinâmicos -------------------
def load_veiculos(caminho_base):
    """
    Tries to load VEÍCULOS.xlsx (or VEICULOS.xlsx) from caminho_base/BD and returns a dict:
    {descricao_stripped: codigo_int_or_str}
    If the file or expected columns are not found, returns None.
    Only returns the display mapping (original descriptions -> code).
    """
    possible_files = [
        os.path.join(caminho_base, "BD", "VEÍCULOS.xlsx"),
        os.path.join(caminho_base, "BD", "VEICULOS.xlsx"),
        os.path.join(caminho_base, "BD", "Veiculos.xlsx"),
        os.path.join(caminho_base, "BD", "VEICULOS.xls")
    ]
    for fpath in possible_files:
        if os.path.exists(fpath):
            try:
                df_veh = pd.read_excel(fpath, sheet_name=0, dtype=str)  # read as str to be safe
                # normalize column names (case-insensitive)
                cols = {c.strip().upper(): c for c in df_veh.columns}
                # find code column (prefer "COD VEICULO" or similar)
                code_col = None
                desc_col = None
                for key_upper, orig in cols.items():
                    if "COD" in key_upper and "VEIC" in key_upper:
                        code_col = orig
                    if "DESCR" in key_upper or "DESC" in key_upper:
                        desc_col = orig
                # fallback: use first column as code and second (or next) as desc
                if code_col is None and len(df_veh.columns) >= 1:
                    code_col = df_veh.columns[0]
                if desc_col is None and len(df_veh.columns) >= 2:
                    # try second column
                    desc_col = df_veh.columns[1]
                if code_col is None or desc_col is None:
                    # can't map properly from this file
                    continue
                veic_map = {}
                for _, r in df_veh.iterrows():
                    desc = str(r.get(desc_col, "")).strip()
                    code_raw = r.get(code_col, "")
                    # try to convert code to int if possible, else keep as string
                    try:
                        code = int(float(str(code_raw).strip()))
                    except Exception:
                        code = str(code_raw).strip()
                    if desc:
                        # store only the original description as display key
                        if desc not in veic_map:
                            veic_map[desc] = code
                if veic_map:
                    print(f"[INFO] Loaded {len(veic_map)} vehicles from {fpath}")
                    return veic_map
            except Exception as e:
                print(f"[WARN] Could not read vehicles file {fpath}: {e}")
    # If we get here, no good file found
    return None

# Keep your original static mapping as a fallback so behavior remains unchanged if file is missing.
_FALLBACK_VEICULOS_DISPLAY = {
    'BIG SIDER': 6, 'BITREM': 7, 'CARRETA': 4, 'CARRETA LINE HAUL': 14,
    'CARRETA REBAIXADA': 9, 'CTNR 20': 15, 'CTNR 40': 16, 'FIORINO': 11,
    'RODOTREM': 8, 'TRUCK 3M': 3, 'TRUCK 3M ALONGADO': 18, 'TRUCK 3M PLUS': 13,
    'TRUCK ALONGADO': 17, 'TRUCK VIAGEM': 2, 'TRUCK VIAGEM PLUS': 12, 'VAN': 10,
    'VANDERLEA': 5, 'VEÍCULO 3/4': 1, 'TRUCK SIDER': 2
}

# load display dict (used for labels) and build lookup dict (case-insensitive) used for mapping
veiculos_display = load_veiculos(caminho_base) or _FALLBACK_VEICULOS_DISPLAY
# build lookup dict (includes uppercase keys for robustness)
veiculos_lookup = {}
for k, v in veiculos_display.items():
    veiculos_lookup[k] = v
    veiculos_lookup[k.upper()] = v
# ------------------------------------------------------------------


def get_vehicle_code(nome_veiculo):
    """
    Robust lookup: try exact name, stripped, upper-case, and finally fallback to None.
    Uses veiculos_lookup (built from the display mapping).
    """
    if nome_veiculo is None:
        return None
    s = str(nome_veiculo).strip()
    if s in veiculos_lookup:
        return veiculos_lookup[s]
    su = s.upper()
    if su in veiculos_lookup:
        return veiculos_lookup[su]
    return None


def normalizar_codigos(campo):
    if pd.isna(campo):
        return []
    return re.split(r'\s*/\s*', str(campo).strip())


def input_demanda(cod_destinos):
    """
    cod_destinos: list of codes entered by the user, e.g. [1080, 1046]
    Returns a DataFrame with all matched rows, saving full COD DESTINO values.
    """
    fluxos = os.path.join(caminho_base, "BD", "FLUXO.xlsx")
    db_fluxos = pd.read_excel(fluxos, sheet_name='FLUXOS')
    
    all_rows = []  # collect all rows here

    # ensure cod_destinos is a list of strings
    cod_destinos = [str(c).strip() for c in cod_destinos]

    for cod_dest in cod_destinos:
        df = Processar_Demandas(cod_dest)
        
        for _, row in df.iterrows():
            cod_forn = str(row["COD FORNECEDOR"]).strip()
            codigo = None
            tipo = None
            cod_ims = None
            cod_dest_full = cod_dest  # default

            for _, linha_fluxo in db_fluxos.iterrows():
                cods_sap = normalizar_codigos(linha_fluxo["COD FORNECEDOR"])
                cods_dest_raw = str(linha_fluxo["COD DESTINO"]).strip()
                cods_dest_split = re.split(r'\s*/\s*', cods_dest_raw)

                if cod_forn in cods_sap and cod_dest in cods_dest_split:
                    nome_veiculo = linha_fluxo["VEICULO PRINCIPAL"]
                    codigo = get_vehicle_code(nome_veiculo)
                    tipo = linha_fluxo.get("TIPO SATURACAO", None)
                    cod_ims = linha_fluxo.get("COD IMS", None)
                    cod_dest_full = cods_dest_raw
                    break

            # append full row data
            all_rows.append({
                "COD FORNECEDOR": cod_forn,
                "COD IMS": cod_ims,
                "COD DESTINO": cod_dest_full,
                "DESENHO": row["DESENHO"],
                "QTDE": row["QTDE"],
                "VEICULO": codigo,
                "TIPO SATURACAO": tipo
            })

    df_final = pd.DataFrame(all_rows)
    df_final.to_excel("Template.xlsx", index=False)
    return df_final  # optionally return for further processing


def apply_filters(event=None):
    """
    Filters the Treeview using "contains" logic for typed text.
    Also handles dropdown selections.
    """
    if event and event.widget.get() == "-- All --":
        event.widget.set('')

    tree.delete(*tree.get_children())

    filters = {col: widget.get() for col, widget in filter_widgets.items()}
    
    column_ids = tree["columns"]

    for row_values in original_tree_data:
        match = True
        row_dict = dict(zip(column_ids, row_values))

        for col_id, filter_value in filters.items():
            if filter_value:
                cell_value = str(row_dict.get(col_id, "")).lower()
                text_to_find = filter_value.lower()
                if text_to_find not in cell_value:
                    match = False
                    break
        
        if match:
            tree.insert("", END, values=row_values)


# Use veiculos_display for UI labels (no duplicate uppercase keys)
veiculos_dict = veiculos_display.copy()


# --------------------- GUI (mantive seu design e cores originais) ---------------------
janela = Tk()
try:
    img = Image.open(resource_path("carreta.png")).resize((140, 100))
    caminhao_img = ImageTk.PhotoImage(img)
except Exception as e:
    print(f"Erro ao carregar imagem da carreta: {e}")
    caminhao_img = None
janela.title("VIAJANTE")
janela.geometry("1400x700")
janela.state('zoomed')

frame_principal = Frame(janela)
frame_principal.pack(fill=BOTH, expand=True)

loading_label = Label(frame_principal, text="Processando... Por favor, aguarde.",
                      font=("Arial", 18, "bold"), bg="white", fg="#007acc",
                      relief="solid", borderwidth=2)

frame_top = Frame(frame_principal)
frame_top.pack(fill=X, padx=10, pady=5)

frame_top.grid_columnconfigure(0, weight=0)
frame_top.grid_columnconfigure(1, weight=1)
frame_top.grid_columnconfigure(2, weight=0)

frame_selecao = Frame(frame_top)
frame_selecao.grid(row=0, column=0, sticky='nw', padx=10)

Label(frame_selecao, text="Selecione o tipo de veículo:", font=("Arial", 10)).grid(row=0, column=0, columnspan=3, pady=(0, 5), sticky='w')

veiculo_var = StringVar(value='')
frame_veiculos = Frame(frame_selecao)
frame_veiculos.grid(row=1, column=0, columnspan=3, sticky='w')

style = ttk.Style()
style.theme_use('clam')
style.configure("Modern.TButton", font=("Arial", 11, "bold"), background="#007acc",
                foreground="white", padding=(10, 5), borderwidth=0, relief="flat")
style.map("Modern.TButton", background=[('active', '#005f9e'), ('!disabled', '#007acc')])
style.configure("Vehicle.Toolbutton", padding=5, font=("Arial", 9), width=17,
                anchor="center", relief="raised")
style.map("Vehicle.Toolbutton", background=[('active', '#e0e0e0'), ('selected', '#007acc')],
          foreground=[('selected', 'white')])

colunas = 3
# build radio buttons from veiculos_dict (which is the display dict)
for i, (nome, cod) in enumerate(sorted(veiculos_dict.items())):
    rb = ttk.Radiobutton(frame_veiculos, text=nome, variable=veiculo_var,
                         value=str(cod), style="Vehicle.Toolbutton")
    rb.grid(row=i // colunas, column=i % colunas, sticky='w', padx=2, pady=2)

label_veiculo = Label(frame_selecao, text="")
label_veiculo.grid(row=2, column=1, columnspan=3, pady=2)
modo_manual = BooleanVar(value=False)
check_manual = Checkbutton(frame_selecao, text="Usar veículo escolhido para todos", variable=modo_manual)
check_manual.grid(row=2, column=0, columnspan=3, sticky='w', pady=(5,5))

btn_atualizar = ttk.Button(frame_selecao, text="Atualizar Dados",
                           command=lambda: atualizar(), style="Modern.TButton")

Label(frame_selecao, text="Cód. Destino:", font=("Arial", 10)).grid(row=2, column=5, pady=(10, 5), padx=(10, 0), sticky='e')
cod_destino_var = StringVar(value='1080')

def validate_numeric(P):
    # Allow digits, commas, and optional spaces
    return all(c.isdigit() or c in [',', ' '] for c in P)


vcmd = (janela.register(validate_numeric), '%P')
entry_cod_destino = Entry(frame_selecao, textvariable=cod_destino_var, width=10, validate="key", validatecommand=vcmd)
entry_cod_destino.grid(row=2, column=6, pady=(10, 5), sticky='w')
btn_atualizar.grid(row=2, column=7, pady=(10, 5), padx=5, sticky="ew")

frame_caminhoes = Frame(frame_top)
frame_caminhoes.grid(row=0, column=0, padx=(550, 0), sticky='nw')
canvas_caminhoes = Canvas(frame_caminhoes, width=360, height=240)
canvas_caminhoes.pack()

frame_resumo = Frame(frame_top)
frame_resumo.grid(row=0, column=2, sticky='ne', padx=20)
tree_resumo = ttk.Treeview(frame_resumo, columns=("Info", "Valor"), show="headings", height=5)
tree_resumo.heading("Info", text="Info")
tree_resumo.heading("Valor", text="Valor")
tree_resumo.column("Info", width=120, anchor='center')
tree_resumo.column("Valor", width=100, anchor='center')
tree_resumo.pack()
for item in ["Ocupação Total", "Qtd Veículos", "Volume Total", "Peso Total", "Embalagens"]:
    tree_resumo.insert("", END, values=(item, ""))

frame_bottom = Frame(frame_principal)
frame_bottom.pack(fill=BOTH, expand=True, padx=10, pady=(0, 10))

frame_filters = Frame(frame_bottom)
frame_filters.pack(fill=X, pady=(5, 2))

scroll_y = Scrollbar(frame_bottom, orient=VERTICAL)
scroll_y.pack(side=RIGHT, fill=Y)
tree = ttk.Treeview(frame_bottom, yscrollcommand=scroll_y.set)
tree.pack(fill=BOTH, expand=True)
scroll_y.config(command=tree.yview)

style.configure("Treeview.Heading", background="#007acc", foreground="white",
                font=("Arial", 10, "bold"), relief="flat")
style.map("Treeview.Heading", background=[('active', '#005f9e')])




def atualizar():
    # --- Start spinner ---
    start_loading()

    def processar():
        try:
            cod_destino_str = cod_destino_var.get()
            try:
                cod_destino_values = [int(code.strip()) for code in cod_destino_str.split(',') if code.strip().isdigit()]
                if not cod_destino_values:
                    cod_destino_values = [1080]  # fallback if empty
            except ValueError:
                cod_destino_values = [1080]

            cod = veiculo_var.get()
            label_veiculo.config(text=f"Código selecionado: {cod}")

            if cod:
                # split input codes by comma
                cod_destino_values = [c.strip() for c in cod_destino_var.get().split(',') if c.strip()]
                df_final = input_demanda(cod_destino_values)  # all codes processed together

                completar_informacoes(
                    tree, int(cod), tree_resumo, canvas_caminhoes, caminhao_img, usar_manual=modo_manual.get()
                )

                global original_tree_data
                original_tree_data = [tree.item(child)['values'] for child in tree.get_children()]

                columns_to_filter = ['COD FORNECEDOR', 'FORNECEDOR', 'DESENHO']
                all_table_columns = list(tree["columns"])

                if not filter_widgets:
                    for widget in frame_filters.winfo_children():
                        widget.destroy()

                    for col_id in columns_to_filter:
                        if col_id in all_table_columns:
                            col_frame = Frame(frame_filters)
                            col_frame.pack(side=LEFT, padx=2, fill=X, expand=True)
                            Label(col_frame, text=col_id, font=("Arial", 8)).pack(anchor='w')
                            combo = ttk.Combobox(col_frame, font=("Arial", 9))
                            combo.pack(fill=X)
                            combo.bind('<KeyRelease>', apply_filters)
                            combo.bind('<<ComboboxSelected>>', apply_filters)
                            filter_widgets[col_id] = combo

                for col_id, combo in filter_widgets.items():
                    col_index = all_table_columns.index(col_id)
                    unique_values = sorted(
                        list(set(str(row[col_index]) for row in original_tree_data if str(row[col_index]).strip()))
                    )
                    combo['values'] = ["-- All --"] + unique_values
                    combo.set('')

                consolidar_dados()

            # --- Stop spinner and show success ---
            loading_label.spinning = False
            janela.after(0, lambda: finalizar_status("Concluído com sucesso!", "#2e8b57"))

        except Exception as e:
            print(f"Ocorreu um erro durante a atualização: {e}")
            loading_label.spinning = False
            janela.after(0, lambda: finalizar_status(f"Erro: {e}", "red"))

    threading.Thread(target=processar, daemon=True).start()



def start_loading():
    spinner_chars = ['|', '/', '--', '\\']
    loading_label.place(relx=0.5, rely=0.5, anchor='center')
    loading_label.lift()
    janela.update_idletasks()

    def spin():
        i = 0
        while getattr(loading_label, "spinning", False):
            loading_label.config(text=f"Processando... {spinner_chars[i % len(spinner_chars)]}")
            i += 1
            janela.update_idletasks()
            threading.Event().wait(0.1)  # short delay for animation

    loading_label.spinning = True
    threading.Thread(target=spin, daemon=True).start()

  
def finalizar_status(msg, color):
    """Atualiza o texto e esconde após 2 segundos"""
    loading_label.config(text=msg, fg=color)
    janela.after(2000, loading_label.place_forget)


janela.mainloop()

