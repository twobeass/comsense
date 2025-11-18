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

### 3. Extract a COM Library

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

### 4. Create/Run the VSCode Extension (Phase 2)

**Phase 2 can be built/demoed using mock data (no extraction required):**

#### Step 1: Install Node.js (v18+) and npm

```bash
node --version
npm --version
```

#### Step 2: Initialize and Build the Extension

```bash
cd vscode-extension
npm install
npm run compile
```

#### Step 3: Launch in VSCode Development Mode

- Open the `vscode-extension` folder in VSCode.
- Press `F5` or select "Run Extension". (Opens Extension Development Host.)

#### Step 4: Test Completion Using Mock API

- In the development environment, create a new file with `.vba` or `.bas` extension.
- Type:

  ```vba
  Dim x As DemoClass
  x.
  ```
- You should see completions for `SomeProp`, `DoThing`, etc., from the mock API (`mock-api.json`).
- Modifying `vscode-extension/data/apis/mock-api.json` replaces the autocompletion set instantly.

#### Step 5: Swap in Real Data

- As soon as real extracted JSONs (from `extract_com.py`) are available, copy them into `vscode-extension/data/apis/` and reload the extension. Completions will then match real COM APIs.

---

### 5. Troubleshooting

- **pywin32 errors:**  
  - Reinstall: `pip install --upgrade pywin32`
  - Make sure you’re running as a standard Windows user (not admin required)
- **COM library not found:**  
  - Try fallback: `"Scripting.FileSystemObject"` always works on Windows
- **No classes/properties/methods:**  
  - The COM library might not expose type info; try another ProgID or check the script logic
- **VSCode extension issues:**
  - Make sure TypeScript sources are compiled (`npm run compile`)
  - Ensure the test file extension is `.vba`, `.bas`, `.cls`, or `.frm`

---

### 6. Commit Example Outputs

- Add and commit working JSON examples:
```bash
git add examples/scripting-fso.json
git add examples/visio-api.json  # Only if extraction succeeded
git commit -m "feat(phase1): add extracted JSON examples\n\nTask: 1.4 from PROJECT_PLAN_MVP.md"
```

---

### 7. Validate Phase 1 Deliverables

- [ ] `extract_com.py` exists and runs without syntax errors
- [ ] At least one JSON file in `examples/` contains valid data
- [ ] JSON contains both `metadata` and `classes` keys
- [ ] At least one class includes properties or methods
- [ ] Basic error handling for missing COM libraries is working

---

### 8. Report Results & Move to Further Phases

- Document the following (in an issue, commit message, or separate file):  
  - Which COM libraries you successfully extracted
  - Encountered issues or notes
  - File sizes of the generated JSONs
  - Are all MVP success criteria for Phase 1 met?
- **Do not begin Phase 2 (production extension) until above is confirmed!** (But you can demo Phase 2 with the mock now.)

---

**Tip:** Each time you make progress, commit often with clear messages as specified in the project plan.

---

This guide ensures both extraction and the editor extension flow are validated with or without real data.
