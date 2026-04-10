import pandas as pd
import os

print("=" * 80)
print("TRACKING PN 520820720")
print("=" * 80)

target_pn = 520820720

# Check if it exists in demand files
print("\n1. Checking demand files...")
demandas_path = "Demandas"
found_in_demands = False

for file in os.listdir(demandas_path):
    if file.endswith(('.txt', '.csv', '.xls', '.xlsx')):
        filepath = os.path.join(demandas_path, file)
        try:
            if file.endswith(('.xls', '.xlsx')):
                df = pd.read_excel(filepath)
                if 'DESENHO' in df.columns:
                    if target_pn in df['DESENHO'].values:
                        print(f"   ✓ Found in {file}")
                        found_in_demands = True
                        print(df[df['DESENHO'] == target_pn][['DESENHO', 'COD ORIGEM', 'ENTREGA SOLICITADA', 'COD DESTINO']].to_string(index=False))
            else:
                # Text file - search for the PN
                with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                    if str(target_pn) in content:
                        print(f"   ✓ Found in {file} (text file)")
                        found_in_demands = True
        except Exception as e:
            pass

if not found_in_demands:
    print("   ✗ NOT found in any demand file")

# Check FLUXO database
print("\n2. Checking FLUXO database...")
df_fluxo = pd.read_excel('BD/FLUXO.xlsx')
print(f"   Total FLUXO rows: {len(df_fluxo)}")

# Check Template.xlsx
print("\n3. Checking Template.xlsx...")
try:
    template = pd.read_excel('Template.xlsx')
    if target_pn in template['DESENHO'].values:
        print(f"   ✓ PN {target_pn} IS in Template.xlsx")
        print(template[template['DESENHO'] == target_pn].to_string(index=False))
    else:
        print(f"   ✗ PN {target_pn} NOT in Template.xlsx")
except Exception as e:
    print(f"   Error reading Template.xlsx: {e}")

print("\n" + "=" * 80)
print("To debug template generation, add this to main.py input_demanda():")
print("=" * 80)
print("""
DEBUG_PN = 520820720

# After processing each demand row:
if row["DESENHO"] == DEBUG_PN:
    print(f"[DEBUG] Found PN {DEBUG_PN}:")
    print(f"  COD FORNECEDOR: {cod_forn}")
    print(f"  COD IMS: {cod_ims_from_file}")
    print(f"  COD DESTINO: {cod_dest}")
    print(f"  QTDE: {row['QTDE']}")
    print(f"  Matched any fluxo: {matched}")
    if matched:
        print(f"  MOT: {mot}")
        print(f"  Should include: {should_include_pn(...)}")
""")
