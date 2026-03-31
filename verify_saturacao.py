import pandas as pd
import numpy as np
from math import ceil

# Check what's in VIAJANTE.xlsx for supplier 800014209
viajante = pd.read_excel('VIAJANTE - Copy.xlsx')
supplier = viajante[viajante['COD FORNECEDOR'] == 800014209]

print('\n' + '='*100)
print('VIAJANTE.XLSX - Supplier 800014209 SAT VOLUME Analysis')
print('='*100)

print(f'\nTotal rows: {len(supplier)}')
total_sat = supplier["SAT VOLUME (%)"].sum()
print(f'Total SAT VOLUME (%): {total_sat:.2f}%')
print(f'Expected CARGAS: {ceil(total_sat / 100)}')

# Group by MDR to see breakdown
mdr_breakdown = supplier.groupby('MDR').agg({
    'QTD EMBALAGENS': 'sum',
    'SAT VOLUME (%)': 'sum'
}).sort_values('SAT VOLUME (%)', ascending=False)

print('\n' + '='*100)
print('SAT VOLUME (%) by MDR in VIAJANTE.xlsx:')
print('='*100)
for mdr, row in mdr_breakdown.iterrows():
    print(f"  • MDR {mdr}: QTD_EMB={row['QTD EMBALAGENS']:.0f}, SAT VOL={row['SAT VOLUME (%)']:.2f}%")

print('\n' + '='*100)
print('Comparing to Debug Output:')
print('='*100)
print('Debug showed:')
print('  • MDR 4315: 714.29% (from 8 rows)')
print('  • MDR E516: 222.22% (from 1 rows)')
print('  • MDR ZHSH: 187.50% (from 1 rows)')
print('  • MDR 3001: 111.91% (from 10 rows)')
print()
print('If these match, calculation is CORRECT.')
print('If they differ, there is a bug in the Excel writing or aggregation.')
print('='*100)
