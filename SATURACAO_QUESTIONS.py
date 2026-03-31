"""
Question for User: What saturação value does the VBA macro return?

Current Python calculation: 1542.93% (16 trucks)

To help identify the issue, please answer:

1. What saturação does the VBA macro calculate for supplier 800014209? 
   (Example: "The macro shows 800%, not 1542%")

2. Is the VBA code you shared for the SAME supplier?
   - The code shows: COD FORNECEDOR = "21544"
   - We're debugging: COD FORNECEDOR = "800014209"
   - Are these the same route?

3. What is the EXPECTED saturação for this supplier/route?

4. Does the VBA macro use:
   - Volume-based calculation? (M³ / Truck M³)
   - Pallet-based calculation? (Pallets / Truck capacity)
   - Something else?

CURRENT CALCULATION METHOD (Python):
=====================================
For each MDR:
  SATURAÇÃO = (Total Boxes / Boxes per Pallet / Truck Capacity) × Efficiency

For supplier 800014209:
  - Total: 788 boxes across 12 MDRs
  - Problem: MDR 4315 has 100 boxes but capacity = 14
    → This alone needs 7.14 trucks (714% saturation)
  - Result: 15.43 trucks total = 1542%

DIAGNOSIS:
==========
The calculation is MATHEMATICALLY CORRECT based on:
  ✓ QTD EMBALAGENS values (verified)
  ✓ CAPACIDADE values from BD_CADASTRO_MDR
  ✓ CXS_POR_PALLET values from BD_CADASTRO_MDR
  ✓ EFICIÊNCIA values from BD_PERDA_COMPRIMENTO

The high value (1542%) is caused by:
  ⚠️ MDR 4315: capacity=14, needs 100 boxes = 714% alone
  ⚠️ Several MDRs are non-palletizable (CXS_POR_PALLET=1)
  ⚠️ Small truck capacity values for TRUCK 3M

NEXT STEPS:
===========
1. Tell me what saturação value you EXPECT
2. Share the VBA calculation formula/logic
3. Confirm if MDR 4315 capacity should really be 14
4. Check if MDRs should be palletizable but aren't
"""

print(__doc__)
