from tkinter import *
from tkinter import ttk
from tkinter import Canvas
from tkinter import messagebox
from PIL import Image, ImageTk
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


def show_temporary_message(master, title, mensagem, kind="info", timeout=10000):
    try:
        top = Toplevel(master)
        top.title(title)
        top.transient(master)
        top.attributes("-topmost", True)
        top.resizable(False, False)

        # Compact frame to mimic the native messagebox layout
        frm = Frame(top, padx=12, pady=8)
        frm.pack(fill=BOTH, expand=True)

        # Use a simple Label with wraplength to resemble messagebox text
        lbl = Label(frm, text=mensagem, justify=LEFT, anchor='w', wraplength=420)
        lbl.pack(fill=BOTH, expand=True)

        # Button frame to center the OK button similar to messagebox
        btn_frm = Frame(frm)
        btn_frm.pack(fill=X, pady=(8, 0))
        btn = Button(btn_frm, text="OK", width=10, command=top.destroy)
        btn.pack(side=RIGHT)

        # center relative to master if possible
        try:
            top.update_idletasks()
            mw = master.winfo_width()
            mh = master.winfo_height()
            mx = master.winfo_rootx()
            my = master.winfo_rooty()
            w = top.winfo_width()
            h = top.winfo_height()
            x = mx + (mw // 2) - (w // 2)
            y = my + (mh // 2) - (h // 2)
            top.geometry(f"+{x}+{y}")
        except Exception:
            pass

        top.after(timeout, top.destroy)
    except Exception:
        # fallback to blocking messagebox
        if kind == "warning":
            messagebox.showwarning(title, mensagem)
        else:
            messagebox.showinfo(title, mensagem)



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


def input_demanda(cod_destinos, use_all_codes=False, sheet_name=None, use_manual=False, manual_veiculo=None):
    """
    cod_destinos: list of codes entered by the user, e.g. [1080, 1046]
    use_all_codes: if True, process all demand rows and map COD DESTINO from FLUXO
    sheet_name: optional sheet name to read from demand files (Geral, Sábado, Domingo)
    Returns a DataFrame with all matched rows, saving full COD DESTINO values.
    """
    fluxos = os.path.join(caminho_base, "BD", "FLUXO.xlsx")
    db_fluxos = pd.read_excel(fluxos, sheet_name='FLUXOS')
    
    # DEBUG: Track specific PN
    DEBUG_PN = 520820720
    
    # Load PN_Conta_trabalho for CT filtering during template creation
    pn_ct_lookup = set()  # Will store (FORNECEDOR/COD_IMS, DESENHO) pairs
    try:
        pn_ct_path = os.path.join(caminho_base, "BD", "PN_Conta_trabalho.xlsx")
        if os.path.exists(pn_ct_path):
            db_pn_ct = pd.read_excel(pn_ct_path)
            # Find columns
            col_map = {}
            for col in db_pn_ct.columns:
                col_upper = str(col).upper().strip()
                if 'FORNECEDOR' in col_upper:
                    col_map['FORNECEDOR'] = col
                elif 'DESENHO' in col_upper:
                    col_map['DESENHO'] = col
                elif 'DESTINO' in col_upper:
                    col_map['DESTINO'] = col
            
            if all(k in col_map for k in ['FORNECEDOR', 'DESENHO']):
                for _, row in db_pn_ct.iterrows():
                    try:
                        forn = str(int(float(row[col_map['FORNECEDOR']]))) if pd.notna(row[col_map['FORNECEDOR']]) else ''
                    except (ValueError, TypeError):
                        forn = str(row[col_map['FORNECEDOR']]).strip()
                    
                    try:
                        desenho = str(int(float(row[col_map['DESENHO']]))) if pd.notna(row[col_map['DESENHO']]) else ''
                    except (ValueError, TypeError):
                        desenho = str(row[col_map['DESENHO']]).strip()
                    
                    if forn and desenho:
                        pn_ct_lookup.add((forn, desenho))  # (FORNECEDOR/COD_IMS, DESENHO)

    except Exception as e:
        print(f"[WARNING] Could not load PN_Conta_trabalho: {e}")
    
    # Helper function to check if a PN should be included based on MOT and CT lookup
    def should_include_pn(desenho, destino, mot, cod_ims):
        """
        Returns True if this PN should be included in the template for the given MOT.
        - If MOT=CT: include only if (cod_ims, desenho) IS in pn_ct_lookup
        - If MOT=FTL/LTL: include only if (cod_ims, desenho) NOT in pn_ct_lookup
        - If no COD IMS or no valid MOT: include by default (skip CT filtering)
        - Handles compound COD IMS like "24149/36190" by checking each part
        """
        if len(pn_ct_lookup) == 0:
            return True  # No CT data, include everything
        
        # Normalize MOT
        mot_upper = str(mot).strip().upper() if pd.notna(mot) else ''
        
        # Skip CT filtering if MOT is not CT or FTL/LTL
        if mot_upper not in ['CT', 'FTL', 'LTL']:
            return True  # Include by default
        
        # Normalize COD IMS for lookup
        try:
            cod_ims_str = str(int(float(cod_ims))) if pd.notna(cod_ims) else ''
        except (ValueError, TypeError):
            cod_ims_str = str(cod_ims).strip() if cod_ims else ''
        
        # Skip CT filtering if no COD IMS
        if not cod_ims_str:
            return True  # Include by default
        
        try:
            des_str = str(int(float(desenho))) if pd.notna(desenho) else ''
        except (ValueError, TypeError):
            des_str = str(desenho).strip()
        
        # Handle compound COD IMS (e.g., "24149/36190")
        # Check if ANY part of the compound IMS matches in the CT lookup
        cod_ims_parts = [p.strip() for p in cod_ims_str.split('/')]
        
        is_in_ct = False
        for ims_part in cod_ims_parts:
            key = (ims_part, des_str)  # (COD_IMS_PART, DESENHO)
            if key in pn_ct_lookup:
                is_in_ct = True
                break
        
        if mot_upper == 'CT':
            return is_in_ct  # CT: include only if in CT lookup
        else:  # FTL or LTL
            return not is_in_ct  # FTL/LTL: include only if NOT in CT lookup
    
    all_rows = []  # collect all rows here

    if use_all_codes:
        # Process all demand rows without filtering by COD DESTINO
        df = Processar_Demandas(None, sheet_name=sheet_name)
        
        print(f"\n[DEBUG] Processing {len(df)} demand rows...")
        debug_found = df[df['DESENHO'] == DEBUG_PN]
        if len(debug_found) > 0:
            print(f"[DEBUG] ✓ Found PN {DEBUG_PN} in demands ({len(debug_found)} times)")
            print(debug_found[['DESENHO', 'COD FORNECEDOR', 'COD IMS', 'QTDE']].to_string(index=False))
        else:
            print(f"[DEBUG] ✗ PN {DEBUG_PN} NOT in demand files")
        
        for _, row in df.iterrows():
            cod_forn = str(row["COD FORNECEDOR"]).strip() if pd.notna(row.get("COD FORNECEDOR")) else None
            cod_ims_from_file = str(row.get("COD IMS", "")).strip() if pd.notna(row.get("COD IMS")) else None
            is_flechinha = int(row.get("IS_FLECHINHA", 0))
           
            # DEBUG: Track specific PN
            if row["DESENHO"] == DEBUG_PN:
                print(f"\n[DEBUG] Processing PN {DEBUG_PN}:")
                print(f"  COD FORNECEDOR: {cod_forn}")
                print(f"  COD IMS from file: {cod_ims_from_file}")
                print(f"  QTDE: {row['QTDE']}")
                print(f"  IS_FLECHINHA: {is_flechinha}")

            matched_cod_forn = cod_forn  # usar o original se não encontrar match
            matched_any = False

            # For each fluxo row, if it matches, append a separate output row
            for _, linha_fluxo in db_fluxos.iterrows():
                fornecedor_str = str(linha_fluxo["COD FORNECEDOR"]).strip()
                cods_dest_raw = str(linha_fluxo["COD DESTINO"]).strip()

                # Pega o COD IMS do fluxo (pode estar na coluna COD IMS)
                fluxo_cod_ims = str(linha_fluxo.get("COD IMS", "")).strip() if pd.notna(linha_fluxo.get("COD IMS")) else None
                fluxo_has_ims = fluxo_cod_ims and fluxo_cod_ims != '' and fluxo_cod_ims != '0'
                
                # Match por COD FORNECEDOR ou por COD IMS
                match_fornecedor = cod_forn and cod_forn in fornecedor_str
                match_ims = cod_ims_from_file and fluxo_cod_ims and cod_ims_from_file in fluxo_cod_ims
                
                # DEBUG: Track matching for specific PN
                if row["DESENHO"] == DEBUG_PN and (match_fornecedor or match_ims):
                    print(f"  [MATCH] Fluxo: Forn={fornecedor_str}, Dest={cods_dest_raw}, IMS={fluxo_cod_ims}")
                    print(f"    match_fornecedor={match_fornecedor}, match_ims={match_ims}")
                
                # For REGULAR demands (FLECHINHA=0): Skip if FLUXO has COD IMS AND destination is 1046
                if is_flechinha == 0 and fluxo_has_ims and cods_dest_raw == '1046':
                    if row["DESENHO"] == DEBUG_PN:
                        print(f"  [SKIP] Skipping FLUXO with IMS={fluxo_cod_ims} and dest=1046 (reserved for flechinha)")
                    continue  # Skip this FLUXO row (reserved for flechinha to 1046)
               
                if match_fornecedor or match_ims:
                    matched_any = True
                    nome_veiculo = linha_fluxo["VEICULO PRINCIPAL"]
                    codigo = get_vehicle_code(nome_veiculo)
                    tipo = linha_fluxo.get("TIPO SATURACAO", None)
                    mot = linha_fluxo.get("MOT", None)
                    fluxo_cod_ims_val = linha_fluxo.get("COD IMS", None)
                    cod_dest_full = cods_dest_raw
                    
                    # Se foi match por IMS e arquivo não trazia fornecedor, usa o fornecedor do fluxo
                    matched_fornecedor_to_use = fornecedor_str if (match_ims and not cod_forn) else matched_cod_forn
                    
                    # Filter: only include if PN belongs to this MOT (CT or FTL)
                    cod_ims_for_check = fluxo_cod_ims_val or cod_ims_from_file
                    should_include = should_include_pn(row["DESENHO"], cod_dest_full, mot, cod_ims_for_check)
                    
                    # DEBUG: Track filtering decision
                    if row["DESENHO"] == DEBUG_PN:
                        print(f"  MOT={mot}, COD_IMS={cod_ims_for_check}, DESTINO={cod_dest_full}")
                        print(f"  should_include_pn returned: {should_include}")
                    
                    if should_include:
                        all_rows.append({
                            "COD FORNECEDOR": matched_fornecedor_to_use,
                            "COD IMS": fluxo_cod_ims_val or cod_ims_from_file,
                            "COD DESTINO": cod_dest_full,
                            "DESENHO": row["DESENHO"],
                            "QTDE": row["QTDE"],
                            "VEICULO": codigo,
                            "TIPO SATURACAO": tipo,
                            "MOT": mot,
                            "FLECHINHA": is_flechinha
                        })
    
                                   
                        
            # If no fluxo match was found for this demand row, still append a row indicating missing fornecedor
            if not matched_any:
                if row["DESENHO"] == DEBUG_PN:
                    print(f"  [NO MATCH] No FLUXO matched for PN {DEBUG_PN}")
                    print(f"    Will create row with no destination/vehicle")
                
                all_rows.append({
                    "COD FORNECEDOR": matched_cod_forn,
                    "COD IMS": cod_ims_from_file,
                    "COD DESTINO": None,
                    "DESENHO": row["DESENHO"],
                    "QTDE": row["QTDE"],
                    "VEICULO": None,
                    "TIPO SATURACAO": None,
                    "MOT": None,
                    "FLECHINHA": is_flechinha
                })
    else:
        # ensure cod_destinos is a list of strings
       
        cod_destinos = [str(c).strip() for c in cod_destinos]
        cod_destinos = list(set(cod_destinos))  # remove duplicates
        
        # Pass sheet_name to Processar_Demandas for saturação file processing
        df = Processar_Demandas(cod_destinos[0], sheet_name=sheet_name)

        for cod_dest in cod_destinos:
            
            for _, row in df.iterrows():
                cod_forn = str(row["COD FORNECEDOR"]).strip() if pd.notna(row.get("COD FORNECEDOR")) else None
                cod_ims_from_file = str(row.get("COD IMS", "")).strip() if pd.notna(row.get("COD IMS")) else None
                is_flechinha = int(row.get("IS_FLECHINHA", 0))
                
                # Check if this is Flechinho data (COD FORNECEDOR = COD IMS)
                is_flechinho = cod_forn and cod_ims_from_file and cod_forn == cod_ims_from_file
                
                codigo = None
                tipo = None
                mot = None
                cod_ims = None
                cod_dest_full = cod_dest  # default
                matched_cod_forn = cod_forn  # usar o original se não encontrar match
                matched = False  # flag to track if a match was found

                for _, linha_fluxo in db_fluxos.iterrows():
                    fornecedor_str = str(linha_fluxo["COD FORNECEDOR"]).strip()
                    cods_dest_raw = str(linha_fluxo["COD DESTINO"]).strip()
                    
                    # Pega o COD IMS do fluxo (pode estar na coluna COD IMS)
                    fluxo_cod_ims = str(linha_fluxo.get("COD IMS", "")).strip() if pd.notna(linha_fluxo.get("COD IMS")) else None
                    fluxo_has_ims = fluxo_cod_ims and fluxo_cod_ims != '' and fluxo_cod_ims != '0'

                    # For REGULAR demands (FLECHINHA=0): Skip if FLUXO has COD IMS AND destination is 1046
                    if is_flechinha == 0 and fluxo_has_ims and cods_dest_raw == '1046':
                        continue  # Skip this FLUXO row (reserved for flechinha to 1046)

                    # Match logic: EXACT for Flechinho, CONTAINS for others
                    if is_flechinho:
                        # Flechinho: Use EXACT match to avoid matching "1094/1097" when we want "1097"
                        match_fornecedor = cod_forn and cod_forn == fornecedor_str
                    else:
                        # Non-Flechinho: Use CONTAINS match (original logic)
                        match_fornecedor = cod_forn and cod_forn in fornecedor_str
                    
                    match_ims = cod_ims_from_file and fluxo_cod_ims and cod_ims_from_file == fluxo_cod_ims
                    
                    # Exact match for COD DESTINO (no splitting, match the whole string)
                    match_cod_dest = str(cod_dest) == cods_dest_raw
                    
                   
                    
                    if (match_fornecedor or match_ims) and match_cod_dest:
                        nome_veiculo = linha_fluxo["VEICULO PRINCIPAL"]
                        codigo = get_vehicle_code(nome_veiculo)
                        tipo = linha_fluxo.get("TIPO SATURACAO", None)
                        mot = linha_fluxo.get("MOT", None)
                        cod_ims = linha_fluxo.get("COD IMS", None)
                        cod_dest_full = cods_dest_raw
                        # print(f"match_fornecedor={fornecedor_str}, cod_dest={cod_dest}, cods_dest_raw={cods_dest_raw}, match_cod_dest={match_cod_dest}")
                        
                        # Se foi match por IMS, pega o COD FORNECEDOR do fluxo
                        if match_ims and not cod_forn:
                            matched_cod_forn = fornecedor_str
                        
                        matched = True  # set flag to True on match
                        break

                # Only append if matched and PN belongs to this MOT
                cod_ims_for_check = cod_ims or cod_ims_from_file
                if matched and should_include_pn(row["DESENHO"], cod_dest_full, mot, cod_ims_for_check):
                    all_rows.append({
                        "COD FORNECEDOR": matched_cod_forn,
                        "COD IMS": cod_ims or cod_ims_from_file,
                        "COD DESTINO": cod_dest_full,
                        "DESENHO": row["DESENHO"],
                        "QTDE": row["QTDE"],
                        "VEICULO": codigo,
                        "TIPO SATURACAO": tipo,
                        "MOT": mot,
                        "FLECHINHA": is_flechinha
                    })

   
    df_final = pd.DataFrame(all_rows).drop_duplicates().reset_index(drop=True)
    
    # DEBUG: Final check
    if DEBUG_PN in df_final['DESENHO'].values:
        print(f"\n[DEBUG] ✓ PN {DEBUG_PN} MADE IT to final template ({len(df_final[df_final['DESENHO'] == DEBUG_PN])} rows)")
        print(df_final[df_final['DESENHO'] == DEBUG_PN][['DESENHO', 'COD FORNECEDOR', 'COD IMS', 'COD DESTINO', 'MOT']].to_string(index=False))
    else:
        print(f"\n[DEBUG] ✗ PN {DEBUG_PN} NOT in final template")
        print(f"  Total rows in template: {len(df_final)}")
   
    # If user chose to force a manual vehicle, override the VEICULO column
    if use_manual and manual_veiculo is not None:
        try:
            df_final['VEICULO'] = int(manual_veiculo)
        except Exception:
            df_final['VEICULO'] = manual_veiculo

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

            # Prepare manual vehicle code to pass into input_demanda and consolidar_dados
            try:
                manual_code = int(cod) if cod not in [None, ''] else None
            except Exception:
                manual_code = cod if cod not in [None, ''] else None

            if cod:
                # Limpa erros anteriores antes de processar
                limpar_erros()
                
                # split input codes by comma
                cod_destino_values = [c.strip() for c in cod_destino_var.get().split(',') if c.strip()]
                # Use all COD DESTINO if checkbox is checked
                use_all = use_all_cod_destino.get()
                # Get selected sheet name from Flechinha dropdown (only if not empty)
                selected_sheet = flechinha_var.get() if flechinha_var.get() else None
                df_final = input_demanda(cod_destino_values, use_all_codes=use_all, sheet_name=selected_sheet, use_manual=modo_manual.get(), manual_veiculo=manual_code)  # all codes processed together

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

                # Prepare manual vehicle code to pass into consolidar_dados
                try:
                    manual_code = int(cod) if cod not in [None, ''] else None
                except Exception:
                    manual_code = cod if cod not in [None, ''] else None

                consolidar_dados(use_manual=modo_manual.get(), manual_veiculo=manual_code)
                
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
                    
                    # Mostra popup com os erros (auto-closing)
                    if erros_criticos:
                        janela.after(0, lambda: show_temporary_message(janela, "Atenção - Problemas Detectados", mensagem, kind="warning", timeout=10000))
                    else:
                        janela.after(0, lambda: show_temporary_message(janela, "Avisos de Processamento", mensagem, kind="info", timeout=10000))

            # --- Stop spinner and show success ---
            loading_label.spinning = False
            janela.after(0, lambda: finalizar_status("Concluído com sucesso!", "#2e8b57"))

        except Exception as e:
            error_msg = str(e)  # Capture error message before list comprehensions shadow 'e'
            adicionar_erro(error_msg, "AVISO")
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
                
                # Mostra popup com os erros (auto-closing)
                if erros_criticos:
                    janela.after(0, lambda: show_temporary_message(janela, "Atenção - Problemas Detectados", mensagem, kind="warning", timeout=10000))
                else:
                    janela.after(0, lambda: show_temporary_message(janela, "Avisos de Processamento", mensagem, kind="info", timeout=10000))
            loading_label.spinning = False
            janela.after(0, lambda msg=error_msg: finalizar_status(f"Erro: {msg}", "red"))

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

# ------------------- Database Update Check -------------------
# Check and update database files from SharePoint if needed
# This runs after GUI is created so we can show progress in the loading_label

def update_progress_callback(message):
    """Callback to update the loading label with progress messages"""
    loading_label.config(text=message, bg="#002855", fg="#FFCC00")
    loading_label.place(relx=0.5, rely=0.5, anchor='center')
    loading_label.lift()
    janela.update_idletasks()

def check_database_updates():
    """Check and update database files in a thread"""
    # Use resource_path to handle both dev and PyInstaller paths
    update_db_path = resource_path('Update DataBase')
    if update_db_path not in sys.path:
        sys.path.insert(0, update_db_path)
    try:
        from Update_Manager import check_and_update_files
        
        # Show initial message
        update_progress_callback("Verificando atualizações do banco de dados...")
        
        # Check files and update if older than 5 days
        update_result = check_and_update_files(
            max_age_days=5, 
            silent=False,
            progress_callback=update_progress_callback
        )
        
        if update_result.get("updated"):
            janela.after(0, lambda: finalizar_status("✓ Banco de dados atualizado!", "#2e8b57"))
        else:
            janela.after(0, lambda: finalizar_status("✓ Banco de dados atualizado!", "#2e8b57"))
            
    except Exception as e:
        print(f"⚠️ Aviso: Não foi possível verificar atualizações: {e}")
        janela.after(0, lambda: loading_label.place_forget())

# Start update check in background thread
threading.Thread(target=check_database_updates, daemon=True).start()
# ---------------------------------------------------------------

janela.mainloop()

