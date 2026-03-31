import pandas as pd

# Check if supplier 800014209 has specific capacity data for MDR 4315
df = pd.read_excel('BD/BD_CADASTRO_MDR.xlsx')
df['MDR'] = df['MDR'].astype(str).str.upper()

supplier = 800014209
mdr = '4315'
truck3m = '8 x 2,4 x 3'

print('\n' + '='*120)
print(f'Checking BD_CADASTRO_MDR for Supplier {supplier} + MDR {mdr}')
print('='*120)

# Check if there are supplier-specific rows
supplier_rows = df[(df['MDR'] == mdr) & (df['CÓD. FORNECEDOR'] == supplier)]
print(f'\n1. Rows with COD FORNECEDOR={supplier} AND MDR={mdr}: {len(supplier_rows)} rows')

if len(supplier_rows) > 0:
    caps = supplier_rows[truck3m].dropna().unique()
    print(f'   TRUCK 3M capacities for this supplier: {caps}')
    print(f'   ✅ Supplier HAS specific capacity data')
else:
    print(f'   ❌ NO supplier-specific rows found')
    print(f'\n2. Falling back to ALL rows with MDR={mdr} (any supplier):')
    all_mdr_rows = df[df['MDR'] == mdr]
    print(f'   Total rows: {len(all_mdr_rows)}')
    
    caps_all = all_mdr_rows[truck3m].dropna()
    print(f'   Non-NaN capacities: {len(caps_all)} values')
    print(f'   Unique capacities: {sorted(caps_all.unique())}')
    print(f'   FIRST non-NaN capacity (current code picks this): {caps_all.values[0] if len(caps_all) > 0 else "None"}')
    print(f'   Most common capacity: {caps_all.mode().values[0] if len(caps_all.mode()) > 0 else "None"}')

print('\n' + '='*120)
print('CURRENT CODE BUG:')
print('='*120)
print('Line 903 in DB.py:')
print('  filtro = db_MDR["MDR"] == mdr')
print('  capacidade_series = db_MDR.loc[filtro, coluna].dropna()')
print('  return capacidade_series.values[0]  # ❌ Returns FIRST, not supplier-specific!')
print('')
print('FIX:')
print('  filtro = (db_MDR["MDR"] == mdr) & (db_MDR["CÓD. FORNECEDOR"] == fornecedor)')
print('  # OR use CHAVE EMBALAGENS if available')
print('='*120)

# Check what CHAVE EMBALAGENS looks like
print('\nCHAVE EMBALAGENS format check:')
sample_chaves = df[df['MDR'] == mdr]['CHAVE EMBALAGENS'].dropna().head(5)
for chave in sample_chaves:
    print(f'  • {chave}')
