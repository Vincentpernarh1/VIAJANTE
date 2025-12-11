from tkinter import *
from tkinter import ttk
from tkinter import Canvas
from tkinter import messagebox
from PIL import Image, ImageTk
# Assuming DB.py contains the functions as used in your original code
from DB import completar_informacoes, consolidar_dados, Processar_Demandas, limpar_erros, obter_erros, adicionar_erro
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
    return [c.strip() for c in re.split(r'\s*,\s*', str(campo).strip()) if c.strip()]


def input_demanda(cod_destinos, use_all_codes=False, sheet_name=None):
    """
    cod_destinos: list of codes entered by the user, e.g. [1080, 1046]
    use_all_codes: if True, process all demand rows and map COD DESTINO from FLUXO
    sheet_name: optional sheet name to read from demand files (Geral, Sábado, Domingo)
    Returns a DataFrame with all matched rows, saving full COD DESTINO values.
    """
    fluxos = os.path.join(caminho_base, "BD", "FLUXO.xlsx")
    db_fluxos = pd.read_excel(fluxos, sheet_name='FLUXOS')
    
    all_rows = []  # collect all rows here

    if use_all_codes:
        # Process all demand rows without filtering by COD DESTINO
        df = Processar_Demandas(None, sheet_name=sheet_name)
        
        for _, row in df.iterrows():
            cod_forn = str(row["COD FORNECEDOR"]).strip() if pd.notna(row.get("COD FORNECEDOR")) else None
            cod_ims_from_file = str(row.get("COD IMS", "")).strip() if pd.notna(row.get("COD IMS")) else None
            
            codigo = None
            tipo = None
            cod_ims = None
            cod_dest_full = None  # to be set from fluxo
            matched_cod_forn = cod_forn  # usar o original se não encontrar match

            for _, linha_fluxo in db_fluxos.iterrows():
                fornecedor_str = str(linha_fluxo["COD FORNECEDOR"]).strip()
                cods_dest_raw = str(linha_fluxo["COD DESTINO"]).strip()
                
                # Pega o COD IMS do fluxo (pode estar na coluna COD IMS)
                fluxo_cod_ims = str(linha_fluxo.get("COD IMS", "")).strip() if pd.notna(linha_fluxo.get("COD IMS")) else None

                # Match por COD FORNECEDOR ou por COD IMS
                match_fornecedor = cod_forn and cod_forn in fornecedor_str
                match_ims = cod_ims_from_file and fluxo_cod_ims and cod_ims_from_file == fluxo_cod_ims
                
                if match_fornecedor or match_ims:
                    nome_veiculo = linha_fluxo["VEICULO PRINCIPAL"]
                    codigo = get_vehicle_code(nome_veiculo)
                    tipo = linha_fluxo.get("TIPO SATURACAO", None)
                    cod_ims = linha_fluxo.get("COD IMS", None)
                    cod_dest_full = cods_dest_raw
                    
                    # Se foi match por IMS, pega o COD FORNECEDOR do fluxo
                    if match_ims and not cod_forn:
                        matched_cod_forn = fornecedor_str
                    
                    break

            # Only append if matched
            if cod_dest_full is not None:
                all_rows.append({
                    "COD FORNECEDOR": matched_cod_forn,
                    "COD IMS": cod_ims or cod_ims_from_file,
                    "COD DESTINO": cod_dest_full,
                    "DESENHO": row["DESENHO"],
                    "QTDE": row["QTDE"],
                    "VEICULO": codigo,
                    "TIPO SATURACAO": tipo
                })
    else:
        # ensure cod_destinos is a list of strings
        cod_destinos = [str(c).strip() for c in cod_destinos]

        for cod_dest in cod_destinos:
            # Pass sheet_name to Processar_Demandas for saturação file processing
            df = Processar_Demandas(cod_dest, sheet_name=sheet_name)
            
            for _, row in df.iterrows():
                cod_forn = str(row["COD FORNECEDOR"]).strip() if pd.notna(row.get("COD FORNECEDOR")) else None
                cod_ims_from_file = str(row.get("COD IMS", "")).strip() if pd.notna(row.get("COD IMS")) else None
                
                codigo = None
                tipo = None
                cod_ims = None
                cod_dest_full = cod_dest  # default
                matched_cod_forn = cod_forn  # usar o original se não encontrar match

                for _, linha_fluxo in db_fluxos.iterrows():
                    fornecedor_str = str(linha_fluxo["COD FORNECEDOR"]).strip()
                    cods_dest_raw = str(linha_fluxo["COD DESTINO"]).strip()
                    
                    # Pega o COD IMS do fluxo (pode estar na coluna COD IMS)
                    fluxo_cod_ims = str(linha_fluxo.get("COD IMS", "")).strip() if pd.notna(linha_fluxo.get("COD IMS")) else None

                    # Match por COD FORNECEDOR ou por COD IMS
                    match_fornecedor = cod_forn and cod_forn in fornecedor_str
                    match_ims = cod_ims_from_file and fluxo_cod_ims and cod_ims_from_file == fluxo_cod_ims
                    
                    # Exact match for COD DESTINO (no splitting, match the whole string)
                    match_cod_dest = str(cod_dest) == cods_dest_raw
                    
                   
                    
                    if (match_fornecedor or match_ims) and match_cod_dest:
                        nome_veiculo = linha_fluxo["VEICULO PRINCIPAL"]
                        codigo = get_vehicle_code(nome_veiculo)
                        tipo = linha_fluxo.get("TIPO SATURACAO", None)
                        cod_ims = linha_fluxo.get("COD IMS", None)
                        cod_dest_full = cods_dest_raw
                        print(f"match_fornecedor={fornecedor_str}, cod_dest={cod_dest}, cods_dest_raw={cods_dest_raw}, match_cod_dest={match_cod_dest}")
                        
                        # Se foi match por IMS, pega o COD FORNECEDOR do fluxo
                        if match_ims and not cod_forn:
                            matched_cod_forn = fornecedor_str
                        
                        break

                # append full row data
                all_rows.append({
                    "COD FORNECEDOR": matched_cod_forn,
                    "COD IMS": cod_ims or cod_ims_from_file,
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
janela.config(bg="#002855")

frame_principal = Frame(janela, bg="#002855")
frame_principal.pack(fill=BOTH, expand=True, pady=(0, 0))

frame_top = Frame(frame_principal, bg="#002855")
frame_top.pack(fill=X, padx=10, pady=5)

# Configure grid columns for proper spacing
frame_top.grid_columnconfigure(0, weight=0, minsize=500)
frame_top.grid_columnconfigure(1, weight=1, minsize=360)
frame_top.grid_columnconfigure(2, weight=0)

frame_selecao = Frame(frame_top, bg="#002855")
frame_selecao.grid(row=0, column=0, sticky='nw', padx=10)

Label(frame_selecao, text="Selecione o tipo de veículo:", font=("Arial", 10, "bold"), bg="#002855", fg="#FFCC00").grid(row=0, column=0, columnspan=3, pady=(0, 3), sticky='w')

veiculo_var = StringVar(value='')
frame_veiculos = Frame(frame_selecao, bg="#002855")
frame_veiculos.grid(row=1, column=0, columnspan=3, sticky='w')

style = ttk.Style()
style.theme_use('clam')
style.configure("Modern.TButton", font=("Arial", 11, "bold"), background="#002855",
                foreground="white", padding=(10, 5), borderwidth=0, relief="flat")
style.map("Modern.TButton", background=[('active', '#004080'), ('!disabled', '#002855')])
style.configure("Highlight.TButton", font=("Arial", 9, "bold"), background="#FFCC00",
                foreground="#002855", padding=(12, 6), borderwidth=2, relief="raised")
style.map("Highlight.TButton", background=[('active', '#FFD633'), ('!disabled', '#FFCC00')])
style.configure("Vehicle.Toolbutton", padding=5, font=("Arial", 9), width=17,
                anchor="center", relief="raised", background="#FFCC00")
style.map("Vehicle.Toolbutton", background=[('active', '#FFD633'), ('selected', '#002855')],
          foreground=[('selected', 'white')])

colunas = 3
# build radio buttons from veiculos_dict (which is the display dict)
for i, (nome, cod) in enumerate(sorted(veiculos_dict.items())):
    rb = ttk.Radiobutton(frame_veiculos, text=nome, variable=veiculo_var,
                         value=str(cod), style="Vehicle.Toolbutton")
    rb.grid(row=i // colunas, column=i % colunas, sticky='w', padx=2, pady=1)

label_veiculo = Label(frame_selecao, text="", bg="#002855", fg="#FFCC00", font=("Arial", 9, "bold"))
label_veiculo.grid(row=2, column=1, columnspan=3, pady=1)

# Custom checkbox function to show dark checkmark
def create_custom_checkbox(parent, text, variable, row, col):
    frame = Frame(parent, bg="#002855")
    frame.grid(row=row, column=col, columnspan=3, sticky='w', pady=(3,2))
    
    # Canvas for custom checkbox
    canvas = Canvas(frame, width=16, height=16, bg="#002855", highlightthickness=0)
    canvas.pack(side=LEFT, padx=(0, 5))
    
    # Draw checkbox background
    canvas.create_rectangle(2, 2, 14, 14, fill="#FFCC00", outline="#FFCC00", tags="box")
    
    # Label for text
    label = Label(frame, text=text, bg="#002855", fg="white", font=("Arial", 9))
    label.pack(side=LEFT)
    
    def toggle():
        variable.set(not variable.get())
        update_display()
    
    def update_display():
        canvas.delete("check")
        if variable.get():
            # Draw checkmark in dark blue
            canvas.create_line(4, 8, 7, 11, fill="#002855", width=2, tags="check")
            canvas.create_line(7, 11, 12, 4, fill="#002855", width=2, tags="check")
    
    canvas.bind("<Button-1>", lambda e: toggle())
    label.bind("<Button-1>", lambda e: toggle())
    
    update_display()
    return frame

modo_manual = BooleanVar(value=False)
check_manual = create_custom_checkbox(frame_selecao, "Usar veículo escolhido para todos", modo_manual, 2, 0)

use_all_cod_destino = BooleanVar(value=False)
check_all_cod = create_custom_checkbox(frame_selecao, "Usar todos os COD DESTINO dos arquivos", use_all_cod_destino, 3, 0)

# Create a frame for Cód. Destino and button on same row as second checkbox
frame_cod_destino = Frame(frame_selecao, bg="#002855")
frame_cod_destino.grid(row=3, column=3, columnspan=5, sticky='w', padx=(20, 0), pady=(0,3))

Label(frame_cod_destino, text="Cód. Destino:", font=("Arial", 10, "bold"), bg="#002855", fg="#FFCC00").pack(side=LEFT)
cod_destino_var = StringVar(value='1080')

def validate_numeric(P):
    # Allow digits, commas, and optional spaces
    return all(c.isdigit() or c in [',', ' ','/',''] for c in P)


vcmd = (janela.register(validate_numeric), '%P')
entry_cod_destino = Entry(frame_cod_destino, textvariable=cod_destino_var, width=12, validate="key", validatecommand=vcmd)
entry_cod_destino.pack(side=LEFT, padx=5)

btn_atualizar = ttk.Button(frame_cod_destino, text="Atualizar Dados",
                           command=lambda: atualizar(), style="Highlight.TButton")
btn_atualizar.pack(side=LEFT, padx=5)

# Flechinha dropdown for sheet selection
Label(frame_cod_destino, text="Flechinha:", font=("Arial", 10, "bold"), bg="#002855", fg="#FFCC00").pack(side=LEFT, padx=(30, 5))
flechinha_var = StringVar(value='')

# Configure style for Flechinha combobox with yellow background
style.configure('Flechinha.TCombobox', fieldbackground='#FFCC00', background='#FFCC00', foreground='#002855')
style.map('Flechinha.TCombobox', 
          fieldbackground=[('readonly', '#FFCC00')],
          selectbackground=[('readonly', '#FFCC00')],
          selectforeground=[('readonly', '#002855')],
          foreground=[('readonly', '#002855')])

flechinha_combo = ttk.Combobox(frame_cod_destino, textvariable=flechinha_var, width=12, state='readonly', 
                                font=("Arial", 9), style='Flechinha.TCombobox')
flechinha_combo['values'] = ['', 'Geral', 'Sábado', 'Domingo']
flechinha_combo.pack(side=LEFT, padx=5)

# Configure the dropdown list colors
janela.option_add('*TCombobox*Listbox.background', '#FFCC00')
janela.option_add('*TCombobox*Listbox.foreground', '#002855')
janela.option_add('*TCombobox*Listbox.selectBackground', '#FFD633')
janela.option_add('*TCombobox*Listbox.selectForeground', '#002855')

frame_caminhoes = Frame(frame_top,  bg="#002855")
frame_caminhoes.grid(row=0, column=0, sticky='ne', padx=(480, 0))
canvas_caminhoes = Canvas(frame_caminhoes, width=450, height=250,  bg="#002855", highlightthickness=0)
canvas_caminhoes.pack()

frame_resumo = Frame(frame_top, bg="#002855")
frame_resumo.grid(row=0, column=1, sticky='nw', padx=(0, 10))

tree_resumo = ttk.Treeview(frame_resumo, columns=("Info", "Valor"), show="headings", height=6)
tree_resumo.heading("Info", text="Info")
tree_resumo.heading("Valor", text="Valor")
tree_resumo.column("Info", width=140, anchor='center')
tree_resumo.column("Valor", width=120, anchor='center')

# Configure row height and font size for summary table
style.configure("Treeview", rowheight=25, font=("Arial", 10))

tree_resumo.pack()
for item in ["Ocupação Total", "Qtd Veículos", "Volume Total", "Peso Total", "Embalagens"]:
    tree_resumo.insert("", END, values=(item, ""))

frame_bottom = Frame(frame_principal, bg="white")
frame_bottom.pack(fill=BOTH, expand=True, padx=10, pady=(0, 0))

# Loading label - position in data area
loading_label = Label(frame_bottom, text="Processando... Por favor, aguarde.",
                      font=("Arial", 14, "bold"), bg="#002855", fg="#FFCC00",
                      relief="solid", borderwidth=2, padx=15, pady=8)

frame_filters = Frame(frame_bottom, bg="#f0f0f0")
frame_filters.pack(fill=X, pady=(5, 2))

# Create a frame for the treeview with scrollbars
tree_frame = Frame(frame_bottom, bg="white")
tree_frame.pack(fill=BOTH, expand=True)

scroll_y = Scrollbar(tree_frame, orient=VERTICAL)
scroll_y.pack(side=RIGHT, fill=Y)

scroll_x = Scrollbar(tree_frame, orient=HORIZONTAL)
scroll_x.pack(side=BOTTOM, fill=X)

tree = ttk.Treeview(tree_frame, yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
tree.pack(fill=BOTH, expand=True)

scroll_y.config(command=tree.yview)
scroll_x.config(command=tree.xview)

style.configure("Treeview.Heading", background="#002855", foreground="#FFCC00",
                font=("Arial", 8, "bold"), relief="flat")
style.map("Treeview.Heading", background=[('active', '#004080')])


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
                # Limpa erros anteriores antes de processar
                limpar_erros()
                
                # split input codes by comma
                cod_destino_values = [c.strip() for c in cod_destino_var.get().split(',') if c.strip()]
                # Use all COD DESTINO if checkbox is checked
                use_all = use_all_cod_destino.get()
                # Get selected sheet name from Flechinha dropdown (only if not empty)
                selected_sheet = flechinha_var.get() if flechinha_var.get() else None
                df_final = input_demanda(cod_destino_values, use_all_codes=use_all, sheet_name=selected_sheet)  # all codes processed together

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
                
                # Mostra erros/avisos se houver
                erros = obter_erros()
                if erros:
                    # Separa erros e avisos
                    erros_criticos = [e for e in erros if '[ERRO]' in e]
                    avisos = [e for e in erros if '[AVISO]' in e]
                    
                    mensagem = ""
                    if erros_criticos:
                        mensagem += "ERROS ENCONTRADOS:\n" + "\n".join(erros_criticos) + "\n\n"
                    if avisos:
                        mensagem += "AVISOS:\n" + "\n".join(avisos)
                    
                    # Mostra popup com os erros
                    if erros_criticos:
                        messagebox.showwarning("Atenção - Problemas Detectados", mensagem)
                    else:
                        messagebox.showinfo("Avisos de Processamento", mensagem)

            # --- Stop spinner and show success ---
            loading_label.spinning = False
            janela.after(0, lambda: finalizar_status("Concluído com sucesso!", "#2e8b57"))

        except Exception as e:
            adicionar_erro(str(e), "AVISO")
            # Mostra erros/avisos se houver
            erros = obter_erros()
            if erros:
                # Separa erros e avisos
                erros_criticos = [e for e in erros if '[ERRO]' in e]
                avisos = [e for e in erros if '[AVISO]' in e]
                
                mensagem = ""
                if erros_criticos:
                    mensagem += "ERROS ENCONTRADOS:\n" + "\n".join(erros_criticos) + "\n\n"
                if avisos:
                    mensagem += "AVISOS:\n" + "\n".join(avisos)
                
                # Mostra popup com os erros
                if erros_criticos:
                    messagebox.showwarning("Atenção - Problemas Detectados", mensagem)
                else:
                    messagebox.showinfo("Avisos de Processamento", mensagem)
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
    # Check if Flechinha was selected
    flechinha_selected = flechinha_var.get() != ''
    
    if "sucesso" in msg.lower():
        if flechinha_selected:
            loading_label.config(text=msg, fg="#002855", bg="#FFCC00", relief="solid", borderwidth=2)
        else:
            loading_label.config(text=msg, fg="#FFCC00", bg="#2e8b57", relief="solid", borderwidth=2)
    else:
        loading_label.config(text=msg, fg="#FFCC00", bg="#002855", relief="solid", borderwidth=2)
    janela.after(2000, loading_label.place_forget)


footer_frame = Frame(janela, bg="#002855", height=18)
footer_frame.pack(side=BOTTOM, fill=X)
footer_frame.pack_propagate(False)

footer_left = Label(footer_frame, text="DHL → STELLANTIS", 
                    font=("Arial", 7, "bold"), bg="#002855", fg="#FFCC00", 
                    anchor="w", padx=8, pady=0)
footer_left.pack(side=LEFT, fill=Y)

footer_right = Label(footer_frame, text="Developer: Vincent Pernarh", 
                     font=("Arial", 7), bg="#002855", fg="#FFCC00", 
                     anchor="e", padx=8, pady=0)
footer_right.pack(side=RIGHT, fill=Y)

janela.mainloop()

