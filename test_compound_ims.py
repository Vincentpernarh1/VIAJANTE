import pandas as pd

print("=" * 80)
print("VERIFYING COMPOUND COD IMS FIX")
print("=" * 80)

# Load PN_Conta_trabalho
df_ct = pd.read_excel('BD/PN_Conta_trabalho.xlsx')
pn_ct_lookup = set()

for _, row in df_ct.iterrows():
    try:
        forn = str(int(row['Fornecedor']))
        desenho = str(int(row['Desenho']))
        pn_ct_lookup.add((forn, desenho))
    except:
        pass

print(f"\n1. Loaded {len(pn_ct_lookup)} (FORNECEDOR, DESENHO) pairs")

# Test compound COD IMS
test_compound = "24149/36190"
test_desenho = "520820720"

print(f"\n2. Testing compound COD IMS: {test_compound} with PN {test_desenho}")

# OLD LOGIC (would fail):
key_old = (test_compound, test_desenho)
found_old = key_old in pn_ct_lookup
print(f"   OLD: Checking exact match ({test_compound}, {test_desenho}): {found_old}")

# NEW LOGIC (should work):
parts = test_compound.split('/')
found_new = False
matched_part = None

for part in parts:
    part = part.strip()
    key_new = (part, test_desenho)
    if key_new in pn_ct_lookup:
        found_new = True
        matched_part = part
        print(f"   NEW: ✓ Found match with part ({part}, {test_desenho})")
        break

if not found_new:
    print(f"   NEW: ✗ No parts of {test_compound} matched")
    # Check what's actually in the CT file for these IMS codes
    for part in parts:
        part = part.strip()
        matches = [k for k in pn_ct_lookup if k[0] == part]
        print(f"   Available PNs for COD IMS {part}: {len(matches)} entries")
        if len(matches) > 0:
            print(f"   Sample: {list(matches)[:5]}")

print("\n" + "=" * 80)
print("RESULT:")
print("=" * 80)
if found_new:
    print(f"✓ FIX WORKS: Compound IMS {test_compound} will match via part {matched_part}")
    print(f"  PN 520820720 with MOT=CT will be INCLUDED in CT flow")
else:
    print(f"✗ FIX WON'T WORK: Neither part of {test_compound} is in PN_Conta_trabalho")
    print(f"  PN 520820720 might need to be added to the CT file")
