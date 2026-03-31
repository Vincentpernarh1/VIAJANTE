import pandas as pd

# Check MDR 4315 in database
df = pd.read_excel('BD/BD_CADASTRO_MDR.xlsx')
df['MDR'] = df['MDR'].astype(str)
df['MDR'] = df['MDR'].str.upper()  # Convert to uppercase to match code

print('\n' + '='*120)
print('Searching for MDR 4315 in BD_CADASTRO_MDR (case-insensitive, exact match on MDR column)')
print('='*120)

# Exact match
rows_exact = df[df['MDR'] == '4315']
print(f'\nExact match for MDR="4315": {len(rows_exact)} rows found')

if len(rows_exact) > 0:
    cols = ['MDR', 'DESCRIÇÃO', 'CAIXA PLÁSTICA', 'CAIXAS POR PALLET', '8 x 2,4 x 3', '14 x 2,4 x 2,78']
    for col in cols:
        if col not in df.columns:
            print(f'Column {col} not found!')
    
    available_cols = [c for c in cols if c in df.columns]
    print('\nData for MDR 4315:')
    print(rows_exact[available_cols].to_string(index=False))
else:
    print('❌ NO ROWS FOUND for exact MDR="4315"')
    print('\nSearching with partial match (contains "4315"):')
    rows_contain = df[df['MDR'].str.contains('4315', na=False)]
    print(f'  Rows containing "4315": {len(rows_contain)}')
    
    if len(rows_contain) > 0:
        for idx, row in rows_contain.head(5).iterrows():
            print(f'    • MDR: "{row["MDR"]}"')

print('\n' + '='*120)
print('CRITICAL QUESTION:')
print('='*120)
print('Your debug output shows MDR 4315 has CAPACIDADE=14.0')
print('But BD_CADASTRO_MDR shows MDR 4315 has NO capacity data.')
print('')
print('Where is the capacity=14 coming from?')
print('  1. Is MDR stored differently? (e.g., with leading zeros, spaces?)')
print('  2. Is the capacity coming from a DIFFERENT column?')
print('  3. Is it using "veic_anterior" (smaller vehicle) capacity?')
print('='*120)

# Check if  there are similar MDR codes
print('\nMDRs that START with "43":'
)
similar = df[df['MDR'].str.startswith('43', na=False)]
print(f'Found {len(similar)} MDRs starting with "43"')
for mdr in similar['MDR'].unique()[:20]:
    print(f'  • {mdr}')
