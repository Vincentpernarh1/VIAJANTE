import ast
s = open('../DB.py','r',encoding='utf-8').read()
for i in range(100, len(s.splitlines())+1, 50):
    try:
        ast.parse('\n'.join(s.splitlines()[:i]))
    except SyntaxError as e:
        print('Fail at line:', i, 'SyntaxError at', e.lineno, e.msg)
        print('\n'.join(s.splitlines()[max(0,e.lineno-5):e.lineno+2]))
        break
else:
    print('No errors found in incremental parse')
