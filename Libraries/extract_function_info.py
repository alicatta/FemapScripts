import ast

# Load the content of both files
with open("FemapAPI_KL.py", "r") as f1, open("FemapAPI.py", "r") as f2:
    femapapi_kl_content = f1.read()
    femapapi_content = f2.read()

# Function to extract functions from the content
def extract_functions(content):
    tree = ast.parse(content)
    return {func.name: func for func in ast.walk(tree) if isinstance(func, ast.FunctionDef)}

# Extract functions from both files
femapapi_kl_functions = extract_functions(femapapi_kl_content)
femapapi_functions = extract_functions(femapapi_content)

# Return the number of functions found in each file for a start
len(femapapi_kl_functions), len(femapapi_functions)

# Compare the functions from both files
common_functions = {}
differing_functions = {}

for func_name, func_kl in femapapi_kl_functions.items():
    if func_name in femapapi_functions:
        common_functions[func_name] = True
        if ast.dump(func_kl) != ast.dump(femapapi_functions[func_name]):
            differing_functions[func_name] = True

# Return the number of common functions and the number of functions that differ in implementation
len(common_functions), len(differing_functions)