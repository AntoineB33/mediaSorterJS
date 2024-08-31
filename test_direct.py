from sympy import symbols, And, Or, Not, simplify

# Define the symbolic variables
word1, word2, word3 = symbols('word1 word2 word3')

# Define the logical expression
expression = And(word1, Or(And(word2, Not(word3)), word2))

# Simplify the expression
simplified_expression = simplify(expression)

# Print the simplified expression
print("Simplified Expression:", simplified_expression)

# Extract the terms directly linked to the result
terms_linked = simplified_expression.free_symbols
print("Terms directly linked to the result:", terms_linked)
