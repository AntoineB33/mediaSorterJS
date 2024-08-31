from sympy import symbols, sympify, simplify

# Define your logical expression as a string
expression = "a | (c & b)"  # This uses SymPy's symbols directly

# Convert the string into a SymPy expression
sympy_expr = sympify(expression)

# Simplify the expression
simplified_expr = simplify(sympy_expr)

# Output the simplified expression
print(simplified_expr)
