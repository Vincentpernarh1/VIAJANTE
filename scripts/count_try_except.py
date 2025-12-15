import os
script_dir = os.path.dirname(os.path.abspath(__file__))
db_path = os.path.abspath(os.path.join(script_dir, '..', 'DB.py'))
s = open(db_path,'r',encoding='utf-8').read().splitlines()
tries = [i+1 for i,l in enumerate(s) if l.strip().startswith('try:')]
excs = [i+1 for i,l in enumerate(s) if l.strip().startswith('except')]
print('tries:', len(tries), tries[:10])
print('excepts:', len(excs), excs[:10])
# show surrounding context for last try before error line 895
for t in tries:
    if t > 800 and t < 900:
        print('try at', t)
for e in excs:
    if e > 800 and e < 920:
        print('except at', e)
print('\nContext around 860-905:')
for i in range(858,906):
    print(i, s[i-1])
