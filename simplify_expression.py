from sympy import symbols, simplify_logic, And, Or, Not, sympify, simplify
import sys
import json

def expr_to_list(expr):
    """
    Convert sympy logical expression to a nested list representation.
    """
    if isinstance(expr, And):
        return [-expr.args + 1] + [expr_to_list(arg) for arg in expr.args]
    elif isinstance(expr, Or):
        return [0] + [expr_to_list(arg) for arg in expr.args]
    elif isinstance(expr, Not):
        return [0, 2, expr_to_list(expr.args[0])]
    else:
        return str(expr)

def main(args):
    expression = args[0].replace("&&", "&").replace("||", "|").replace("!", "~")

    # Convert the string into a SymPy expression
    sympy_expr = sympify(expression)
    # sympy_expr = sympify("B | ~C & ~B")

    # Simplify the expression
    simplified_expr = simplify(sympy_expr)
    
    # result = expr_to_list(simplified_expr)

    # Output the simplified expression
    print(json.dumps(str(simplified_expr)))

if __name__ == "__main__":
    import json
    import sys
    main(sys.argv[1:])