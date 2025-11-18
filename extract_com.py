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
    
    # TODO: Implement extraction logic
    
    return {
        "metadata": {
            "prog_id": prog_id,
            "version": "0.1.0-mvp"
        },
        "classes": {}
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
