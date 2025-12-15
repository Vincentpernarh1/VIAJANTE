import os
script_dir = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(script_dir, '..', 'DB.py'), 'r', encoding='utf-8') as f:
    s = f.readlines()
for t in [76,92,159,480,886,912,986,1057]:
    print('---',t,'--',s[t-1].rstrip())
for e in [104,168,228,914,921,988,1059]:
    print('***',e,'--',s[e-1].rstrip())
