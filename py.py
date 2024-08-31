from pyeda.inter import *

a, b = map(exprvar, 'ab')
# not(a and b) or (not(b) and not(b))
expr = a & ~b | (~b & ~b)
# expr = a & ~(a & b) | (~b & ~b) & a & a
simplified_expr = expr.simplify()
print(simplified_expr)  # Outputs: ~b
