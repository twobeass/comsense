# comsense: User To-Do Guide

## Phase 1: COM Extraction Script — Checklist & Next Steps

### 1. Prerequisites

- Run on **Windows** (COM extraction only works on Windows)
- Ensure **Python 3.10+** is installed  
  - Check: `python --version`
- Install **pywin32** dependency  
  - Run: `pip install pywin32`
  - Test: `python -c "import win32com.client; print('OK')"`

---

### 2. List Available COM Libraries

**To find the correct names for extraction, run:**

```bash
python list_com_progids.py
```
- This prints available ProgIDs and CLSIDs ("ProgID" column is what you enter for extraction)
- Use the ProgID exactly as printed for `extract_com.py` (e.g., `Visio.Application`, `Excel.Application`, etc.)

---

### 3. Run Extraction Script

#### Step 1: Extract Scripting.FileSystemObject (baseline test)
```bash
python extract_com.py "Scripting.FileSystemObject" "examples/scripting-fso.json"
```
- Confirm:  
  - The script prints “✓ Saved to examples/scripting-fso.json”
  - The `examples/` folder and JSON file exist

#### Step 2: Extract Visio COM API (if Microsoft Visio is installed)
```bash
python extract_com.py "Visio.Application" "examples/visio-api.json"
```
- Confirm:  
  - The script succeeds, or prints a clear message if Visio is not installed

#### Step 3: Check JSON Output
- Open the generated JSON files in `examples/`
- Confirm both of the following:
  - JSON is valid and opens in an editor
  - Structure matches the sample:
    ```json
    {
      "metadata": { "prog_id": "...", "version": "...", "generator": "comsense-mvp" },
      "classes": {
        "FileSystemObject": {
          "properties": { ... },
          "methods": { ... }
        }
      }
    }
    ```
  - At least one class with at least one property or method is present

---

### 4. Troubleshooting

- **pywin32 errors:**  
  - Reinstall: `pip install --upgrade pywin32`
  - Make sure you’re running as a standard Windows user (not admin required)
- **COM library not found:**  
  - Try fallback: `"Scripting.FileSystemObject"` always works on Windows
- **No classes/properties/methods:**  
  - The COM library might not expose type info; try another ProgID or check the script logic

---

### 5. Commit Example Outputs

- Add and commit working JSON examples:
```bash
git add examples/scripting-fso.json
git add examples/visio-api.json  # Only if extraction succeeded
git commit -m "feat(phase1): add extracted JSON examples\n\nTask: 1.4 from PROJECT_PLAN_MVP.md"
```

---

### 6. Validate Phase 1 Deliverables

- [ ] `extract_com.py` exists and runs without syntax errors
- [ ] At least one JSON file in `examples/` contains valid data
- [ ] JSON contains both `metadata` and `classes` keys
- [ ] At least one class includes properties or methods
- [ ] Basic error handling for missing COM libraries is working

---

### 7. Report Results & Move to Phase 2

- Document the following (in an issue, commit message, or separate file):  
  - Which COM libraries you successfully extracted
  - Encountered issues or notes
  - File sizes of the generated JSONs
  - Are all MVP success criteria for Phase 1 met?
- **Do not begin Phase 2 until above is confirmed!**

---

### 8. Ready? Proceed to VSCode Extension (Phase 2)

- When all above boxes are ticked, move to the extension work as described in `PROJECT_PLAN_MVP.md` Phase 2.

---

**Tip:** Each time you make progress, commit often with clear messages as specified in the project plan.

---

This guide ensures MVP Phase 1 is proven on your system before continuing.
