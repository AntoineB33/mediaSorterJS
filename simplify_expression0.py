# my_script.py

import sys
import json

def my_python_function(arg1, arg2):
    return f"Received: {arg1} and {arg2}"

if __name__ == "__main__":
    # Assume arguments are passed as JSON via stdin
    input_json = sys.stdin.read()
    args = json.loads(input_json)
    
    result = my_python_function(args['arg1'], args['arg2'])
    
    # Output result as JSON
    print(json.dumps({"result": result}))
