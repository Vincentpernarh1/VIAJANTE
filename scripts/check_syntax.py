import ast, sys
s = open('DB.py','r',encoding='utf-8').read()
try:
    ast.parse(s)
    print('AST parsed OK')
except SyntaxError as e:
    print('SyntaxError', e.msg, 'at', e.lineno, e.offset)
    print('\n'.join(s.splitlines()[e.lineno-5:e.lineno+2]))
