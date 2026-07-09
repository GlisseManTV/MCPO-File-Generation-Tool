import ast
import json
import sys
import os


def imported_from_utils(server_path: str):
    with open(server_path, "r", encoding="utf-8") as f:
        src = f.read()
    tree = ast.parse(src, server_path)
    required = set()
    has_resolve = False

    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and node.module == "utils":
            for alias in node.names:
                # We care about symbol names expected to be available from utils
                required.add(alias.asname or alias.name)

    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and node.module == "utils.pptx_treatment":
            for alias in node.names:
                if alias.name == "_resolve_donor_simple":
                    has_resolve = True

    return sorted(required), has_resolve


def exported_from_utils_init(utils_init_path: str):
    with open(utils_init_path, "r", encoding="utf-8") as f:
        src = f.read()
    tree = ast.parse(src, utils_init_path)
    exported = set()

    for node in tree.body:
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name) and target.id == "__all__":
                    lst = node.value
                    if isinstance(lst, (ast.List, ast.Tuple)):
                        for elt in lst.elts:
                            if isinstance(elt, ast.Str):
                                exported.add(elt.s)
                            elif isinstance(elt, ast.Constant) and isinstance(elt.value, str):
                                exported.add(elt.value)

    return sorted(exported)


if __name__ == "__main__":
    server_path = os.path.join("tools", "server.py")
    utils_init_path = os.path.join("utils", "__init__.py")

    required, has_resolve = imported_from_utils(server_path)
    exported = exported_from_utils_init(utils_init_path)

    missing = [name for name in required if name not in exported]

    result = {
        "required_from_server": required,
        "exported_in_utils_all": exported,
        "missing_in_utils_all": missing,
        "has__resolve_donor_simple": has_resolve,
    }
    print(json.dumps(result, indent=2, ensure_ascii=False))
    sys.exit(0 if not missing and has_resolve else 1)
