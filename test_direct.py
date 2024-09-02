import sympy as sp

def find_false_terms(expression_str):
    # Convert the string expression into a sympy logical expression
    expression = sp.sympify(expression_str)
    
    # Find the variables in the expression
    variables = expression.free_symbols
    
    # Prepare a dictionary to store whether setting each variable to False makes the expression False
    false_terms = []
    
    for var in variables:
        # Substitute the variable with False in the expression
        test_expr = expression.subs(var, False)
        
        # Check if the whole expression becomes False
        if test_expr == False:
            false_terms.append(str(var))
    
    return false_terms

# Example usage:
expression_str = "~(A | C & ~A) & (B | ~~C & B)"
false_terms = find_false_terms(expression_str)
print("Terms that if false make the expression false:", false_terms)
