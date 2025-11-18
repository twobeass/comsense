"""comsense MVP - COM Type Library Extractor

Usage:
    python extract_com.py <ProgID> <output_file>
    
Example:
    python extract_com.py "Visio.Application" "examples/visio-api.json"
"""

import sys
import json
from pathlib import Path
import win32com.client
import pywintypes

def extract_class(com_class) -> dict:
    """
    Extract properties and methods from a COM class.
    
    Args:
        com_class: Generated COM class from win32com
        
    Returns:
        Dictionary with properties and methods
    """
    class_info = {
        "properties": {},
        "methods": {}
    }
    
    # Extract properties
    if hasattr(com_class, "_prop_map_get_"):
        for prop_name, prop_info in com_class._prop_map_get_.items():
            class_info["properties"][prop_name] = {
                "type": "variant",  # Simplified - don't parse VT types yet
                "readonly": prop_name not in getattr(com_class, "_prop_map_put_", {})
            }
    
    # Extract methods
    if hasattr(com_class, "_method_map_"):
        for method_name, method_info in com_class._method_map_.items():
            # method_info is typically (dispid, flags, arg_types, ...)
            class_info["methods"][method_name] = {
                "parameters": []  # Simplified - don't extract params yet
            }
    
    return class_info

def extract_com_library(prog_id: str) -> dict:
    """
    Extract COM type library to dictionary.
    
    Args:
        prog_id: COM ProgID (e.g., "Visio.Application")
        
    Returns:
        Dictionary with structure:
        {
            "metadata": {...},
            "classes": {...}
        }
    """
    print(f"Extracting {prog_id}...")
    
    # Generate COM wrapper
    try:
        obj = win32com.client.gencache.EnsureDispatch(prog_id)
    except pywintypes.com_error as e:
        raise RuntimeError(f"Failed to load {prog_id}: {e}")
    
    # Get generated module
    module = obj.__class__.__module__
    if not module or module == "win32com.gen_py":
        raise RuntimeError(f"No type library wrapper generated for {prog_id}")
    
    # Import the generated module
    import importlib
    gen_module = importlib.import_module(module)
    
    # Extract classes
    classes = {}
    
    # Iterate through module attributes to find COM classes
    for name in dir(gen_module):
        attr = getattr(gen_module, name)
        
        # Check if it's a COM class (has _prop_map_get_ or _method_map_)
        if hasattr(attr, "_prop_map_get_") or hasattr(attr, "_method_map_"):
            classes[name] = extract_class(attr)
    
    print(f"  Found {len(classes)} classes")
    
    return {
        "metadata": {
            "prog_id": prog_id,
            "version": "0.1.0-mvp",
            "generator": "comsense-mvp"
        },
        "classes": classes
    }

def main():
    if len(sys.argv) != 3:
        print("Usage: python extract_com.py <ProgID> <output_file>")
        sys.exit(1)
    
    prog_id = sys.argv[1]
    output_file = sys.argv[2]
    
    try:
        data = extract_com_library(prog_id)
        
        # Save to JSON
        Path(output_file).parent.mkdir(parents=True, exist_ok=True)
        Path(output_file).write_text(json.dumps(data, indent=2))
        
        print(f"✓ Saved to {output_file}")
        
    except Exception as e:
        print(f"✗ Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
