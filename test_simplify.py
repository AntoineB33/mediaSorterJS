def parse_expression(expression):
    # Define the precedence of operators
    precedence = {'!': 3, '&': 2, '|': 1}
    
    def is_operator(c):
        return c in precedence

    def apply_operator(operators, values):
        operator = operators.pop()
        if operator == '!':
            value = values.pop()
            values.append(not value)
        else:
            right = values.pop()
            left = values.pop()
            if operator == '&':
                values.append(left and right)
            elif operator == '|':
                values.append(left or right)

    def to_postfix(expression):
        output = []
        operators = []
        i = 0
        while i < len(expression):
            c = expression[i]
            if c.isalnum():
                output.append(c)
            elif c == '(':
                operators.append(c)
            elif c == ')':
                while operators and operators[-1] != '(':
                    output.append(operators.pop())
                operators.pop()  # Remove the '('
            elif is_operator(c):
                while (operators and operators[-1] != '(' and
                       precedence[operators[-1]] >= precedence[c]):
                    output.append(operators.pop())
                operators.append(c)
            i += 1
        
        while operators:
            output.append(operators.pop())
        
        return ''.join(output)

    postfix = to_postfix(expression)
    return postfix

def evaluate_postfix(postfix, values):
    stack = []
    for c in postfix:
        if c.isalnum():
            stack.append(values[c])
        elif c == '!':
            stack.append(not stack.pop())
        elif c == '&':
            b = stack.pop()
            a = stack.pop()
            stack.append(a and b)
        elif c == '|':
            b = stack.pop()
            a = stack.pop()
            stack.append(a or b)
    return stack[0]

def simplify_expression(expression, values):
    postfix = parse_expression(expression)
    return evaluate_postfix(postfix, values)


# Example usage
expression1 = "A & !B | A"
expression2 = "!(A | !A) & (B | !!C)"

values1 = {'A': True, 'B': False, 'C': True}  # Example values for variables
values2 = {'A': False, 'B': True, 'C': False}

simplified_expr1 = simplify_expression(expression1, values1)
simplified_expr2 = simplify_expression(expression2, values2)

print(simplified_expr1)  # Output: True (example output)
print(simplified_expr2)  # Output: False (example output)
