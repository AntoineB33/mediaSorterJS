import ast

class Node:
    def __init__(self, value):
        self.value = value
        self.left = None
        self.right = None

def parse_expression(expr):
    """Parses a logical expression and creates a binary tree."""
    # Replace '&&' with 'and' and '||' with 'or' to match Python syntax
    expr = expr.replace('&&', 'and').replace('||', 'or')
    tree = ast.parse(expr, mode='eval')
    return build_tree(tree.body)

def build_tree(node):
    """Recursively builds a tree from the AST node."""
    if isinstance(node, ast.BoolOp):  # Handles and/or
        op = '&&' if isinstance(node.op, ast.And) else '||'
        root = Node(op)
        root.left = build_tree(node.values[0])
        root.right = build_tree(node.values[1])
        return root
    elif isinstance(node, ast.Compare):  # Handles comparisons like A == 1
        left = node.left.id if isinstance(node.left, ast.Name) else node.left
        right = node.comparators[0].n if isinstance(node.comparators[0], ast.Constant) else node.comparators[0]
        return Node(f"{left} == {right}")
    elif isinstance(node, ast.Name):
        return Node(node.id)
    elif isinstance(node, ast.Constant):
        return Node(str(node.value))
    else:
        raise ValueError("Unsupported AST node.")

def find_necessary_subformulas(node):
    """Traverses the tree to find necessary subformulas."""
    if node.value == '&&':
        left = find_necessary_subformulas(node.left)
        right = find_necessary_subformulas(node.right)
        return left + right
    elif node.value == '||':
        left = find_necessary_subformulas(node.left)
        right = find_necessary_subformulas(node.right)
        common = set(left).intersection(set(right))
        return list(common)
    else:
        return [node.value]

# Example usage:
formula = "(A == 1) && (B == 2 || (C == 3) && (D == 4))"
formula = "(A == 1) && ((C == 3) && (B == 2) || (C == 3) && (D == 4))"
formula = "(A == 1) && (B == 2) || (C == 3) && (D == 4)"
formula = "(A == 1) && (B == 2 || (C == 3) && (D == 4))"
formula = "(A == 1) && (C == 3) || (C == 3) && (D == 4)"
tree = parse_expression(formula)
necessary_subformulas = find_necessary_subformulas(tree)

print("Necessary Subformulas:", necessary_subformulas)
