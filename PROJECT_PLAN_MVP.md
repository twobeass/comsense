# comsense – MVP Implementation Plan

**Status:** Planning (MVP Fast-Track)  
**Primary audience:** Autonomous coding agents  
**Language:** English only  
**Version:** 0.1.0-mvp  
**Last updated:** 2025-11-18  
**Goal:** Validate the concept with a working VSCode extension in 10 days

---

## Philosophy: Validate Before Polish

This MVP plan prioritizes **speed to validation** over completeness. Build the minimum needed to test if COM extraction → VSCode IntelliSense is viable, then iterate based on real usage.

**Core Principle:** Working demo in 2 weeks beats perfect tool in 2 months.

---

## 1. Project Overview

### 1.1 MVP Scope

**What we're building:**

1. **Simple Python script** (`extract_com.py`) that extracts ONE COM library to JSON
2. **Basic VSCode extension** that provides IntelliSense from that JSON
3. **Proof that the approach works** with real VBA code

**What we're NOT building (yet):**

- Full CLI with multiple commands
- Registry scanner
- JSON schema validation
- Test suite
- Documentation beyond README
- Community repository infrastructure

**Success criteria:** Can you type `Application.` in VSCode and see real Visio methods?

### 1.2 Timeline

- **Phase 1:** Extraction script (3 days)
- **Phase 2:** VSCode extension (5 days)
- **Phase 3:** Validation (2 days)
- **Total:** 10 days to working demo

---

## 2. Repository Structure (MVP)

```
comsense/
├── extract_com.py              # Single-file extraction script (~150 lines)
├── examples/
│   ├── visio-api.json          # Extracted Visio API (Phase 1 output)
│   └── scripting-fso.json      # Extracted FileSystemObject (Phase 1 output)
├── vscode-extension/           # Minimal extension (Phase 2)
│   ├── package.json
│   ├── tsconfig.json
│   ├── src/
│   │   └── extension.ts        # Main extension logic (~100 lines)
│   └── data/
│       └── apis/               # Copied from examples/
│           ├── visio-api.json
│           └── scripting-fso.json
├── README.md                   # Updated with MVP instructions
├── PROJECT_PLAN_MVP.md         # This file
└── PROJECT_PLAN.md             # Full plan (for later)
```

**Note:** No `comsense/` package directory, no `tests/`, no `schemas/` yet.

---

## 3. Phase 1: Extraction Script (Days 1-3)

### 3.1 Goal

Create `extract_com.py` that extracts a COM library to a usable JSON format.

### 3.2 Requirements

**Platforms:**
- Windows only (COM dependency)
- Python 3.10+

**Dependencies:**
- `pywin32>=305` only
- No additional packages

### 3.3 Implementation Tasks

#### Task 1.1: Create extraction script skeleton

**File:** `extract_com.py`

```python
"""
comsense MVP - COM Type Library Extractor

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
```

**Definition of Done:**
- Script runs without errors
- Takes ProgID and output path as arguments
- Creates JSON file at specified location

#### Task 1.2: Implement basic type extraction

**Add to `extract_com_library()` function:**

```python
def extract_com_library(prog_id: str) -> dict:
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
```

**Definition of Done:**
- Successfully loads COM library with `EnsureDispatch`
- Finds generated wrapper module
- Extracts class names from module
- Handles errors gracefully

#### Task 1.3: Extract class members (properties and methods)

**Add helper function:**

```python
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
```

**Definition of Done:**
- Extracts property names and readonly status
- Extracts method names
- Returns structured dictionary
- Simplified types (no VT mapping yet)

#### Task 1.4: Test extraction with real COM libraries

**Create `examples/` directory and extract:**

```bash
# Extract Scripting.FileSystemObject (always available on Windows)
python extract_com.py "Scripting.FileSystemObject" "examples/scripting-fso.json"

# Extract Visio.Application (if installed)
python extract_com.py "Visio.Application" "examples/visio-api.json"
```

**Manual validation:**

1. Open `examples/scripting-fso.json`
2. Check structure:
   ```json
   {
     "metadata": {
       "prog_id": "Scripting.FileSystemObject",
       ...
     },
     "classes": {
       "FileSystemObject": {
         "properties": {
           "Drives": {...}
         },
         "methods": {
           "GetFolder": {...},
           "CreateTextFile": {...}
         }
       }
     }
   }
   ```
3. Verify classes and methods look reasonable

**Definition of Done:**
- Successfully extracts `Scripting.FileSystemObject`
- Successfully extracts `Visio.Application` (if available)
- JSON is valid and readable
- Contains classes with properties and methods

### 3.4 Phase 1 Deliverables

- [ ] `extract_com.py` script (~150 lines)
- [ ] `examples/scripting-fso.json`
- [ ] `examples/visio-api.json` (if Visio installed)
- [ ] Basic error handling for missing COM libraries

### 3.5 Phase 1 Time Estimate

- Task 1.1: 2 hours
- Task 1.2: 4 hours
- Task 1.3: 4 hours
- Task 1.4: 2 hours
- **Total:** ~12 hours (1.5 days with debugging)

---

## 4. Phase 2: VSCode Extension (Days 4-8)

### 4.1 Goal

Create a minimal VSCode extension that provides IntelliSense from extracted JSON.

### 4.2 Requirements

**Tools:**
- Node.js 18+
- VSCode 1.80+
- `npm` or `yarn`

**Extension capabilities:**
- Autocomplete class names after typing
- Autocomplete properties/methods after typing `.`
- Basic hover documentation

### 4.3 Implementation Tasks

#### Task 2.1: Initialize VSCode extension

**Commands:**

```bash
# Create extension directory
mkdir vscode-extension
cd vscode-extension

# Initialize npm project
npm init -y

# Install dependencies
npm install --save-dev @types/vscode @types/node typescript
npm install --save-dev @vscode/vsce
```

**Create `package.json`:**

```json
{
  "name": "comsense",
  "displayName": "COM IntelliSense (MVP)",
  "description": "IntelliSense for COM libraries (Visio, Excel, etc.) in VBA",
  "version": "0.1.0",
  "engines": {
    "vscode": "^1.80.0"
  },
  "categories": ["Programming Languages"],
  "activationEvents": ["onLanguage:vb"],
  "main": "./out/extension.js",
  "contributes": {
    "languages": [{
      "id": "vb",
      "aliases": ["Visual Basic", "VBA"],
      "extensions": [".vba", ".bas", ".cls", ".frm"]
    }]
  },
  "scripts": {
    "compile": "tsc -p ./",
    "watch": "tsc -watch -p ./"
  },
  "devDependencies": {
    "@types/node": "^18.0.0",
    "@types/vscode": "^1.80.0",
    "typescript": "^5.0.0"
  }
}
```

**Create `tsconfig.json`:**

```json
{
  "compilerOptions": {
    "module": "commonjs",
    "target": "ES2020",
    "outDir": "out",
    "lib": ["ES2020"],
    "sourceMap": true,
    "rootDir": "src",
    "strict": true
  },
  "exclude": ["node_modules", ".vscode-test"]
}
```

**Definition of Done:**
- `npm install` runs without errors
- Extension skeleton created
- TypeScript configured

#### Task 2.2: Create data loading infrastructure

**Create `vscode-extension/data/apis/` directory:**

```bash
mkdir -p data/apis

# Copy extracted JSON files
cp ../examples/visio-api.json data/apis/
cp ../examples/scripting-fso.json data/apis/
```

**Create `src/types.ts`:**

```typescript
/**
 * Type definitions for extracted COM API data
 */

export interface ComApi {
    metadata: {
        prog_id: string;
        version: string;
        generator: string;
    };
    classes: {
        [className: string]: ComClass;
    };
}

export interface ComClass {
    properties: {
        [propName: string]: {
            type: string;
            readonly: boolean;
        };
    };
    methods: {
        [methodName: string]: {
            parameters: any[];  // Simplified for MVP
        };
    };
}
```

**Create `src/apiLoader.ts`:**

```typescript
import * as fs from 'fs';
import * as path from 'path';
import { ComApi } from './types';

/**
 * Load all COM API JSON files from data/apis directory
 */
export class ApiLoader {
    private apis: Map<string, ComApi> = new Map();

    constructor(private dataDir: string) {}

    load(): void {
        const apisDir = path.join(this.dataDir, 'apis');
        
        if (!fs.existsSync(apisDir)) {
            console.warn('APIs directory not found:', apisDir);
            return;
        }

        const files = fs.readdirSync(apisDir).filter(f => f.endsWith('.json'));
        
        for (const file of files) {
            try {
                const filePath = path.join(apisDir, file);
                const content = fs.readFileSync(filePath, 'utf-8');
                const api: ComApi = JSON.parse(content);
                
                this.apis.set(api.metadata.prog_id, api);
                console.log(`Loaded API: ${api.metadata.prog_id}`);
            } catch (error) {
                console.error(`Failed to load ${file}:`, error);
            }
        }
    }

    getAll(): ComApi[] {
        return Array.from(this.apis.values());
    }

    getByProgId(progId: string): ComApi | undefined {
        return this.apis.get(progId);
    }
}
```

**Definition of Done:**
- Type definitions created
- API loader can read JSON files from `data/apis/`
- Error handling for missing/invalid files

#### Task 2.3: Implement completion provider

**Create `src/completionProvider.ts`:**

```typescript
import * as vscode from 'vscode';
import { ComApi, ComClass } from './types';

/**
 * Provides IntelliSense completions for COM APIs
 */
export class ComCompletionProvider implements vscode.CompletionItemProvider {
    constructor(private apis: ComApi[]) {}

    provideCompletionItems(
        document: vscode.TextDocument,
        position: vscode.Position
    ): vscode.CompletionItem[] {
        const lineText = document.lineAt(position).text;
        const textBeforeCursor = lineText.substring(0, position.character);

        // Check if we're after a dot (e.g., "Application.")
        const dotMatch = textBeforeCursor.match(/([\w]+)\.$/)
        if (dotMatch) {
            const objectName = dotMatch[1];
            return this.getClassMembers(objectName);
        }

        // Otherwise, suggest class names
        return this.getAllClasses();
    }

    private getAllClasses(): vscode.CompletionItem[] {
        const items: vscode.CompletionItem[] = [];

        for (const api of this.apis) {
            for (const className of Object.keys(api.classes)) {
                const item = new vscode.CompletionItem(
                    className,
                    vscode.CompletionItemKind.Class
                );
                item.detail = api.metadata.prog_id;
                items.push(item);
            }
        }

        return items;
    }

    private getClassMembers(className: string): vscode.CompletionItem[] {
        const items: vscode.CompletionItem[] = [];

        for (const api of this.apis) {
            const comClass = api.classes[className];
            if (!comClass) continue;

            // Add properties
            for (const [propName, propInfo] of Object.entries(comClass.properties)) {
                const item = new vscode.CompletionItem(
                    propName,
                    vscode.CompletionItemKind.Property
                );
                item.detail = propInfo.type;
                if (propInfo.readonly) {
                    item.detail += ' (readonly)';
                }
                items.push(item);
            }

            // Add methods
            for (const methodName of Object.keys(comClass.methods)) {
                const item = new vscode.CompletionItem(
                    methodName,
                    vscode.CompletionItemKind.Method
                );
                item.insertText = new vscode.SnippetString(`${methodName}($0)`);
                items.push(item);
            }
        }

        return items;
    }
}
```

**Definition of Done:**
- Completion provider returns class names
- Completion provider returns properties/methods after `.`
- Items have correct icons (class, property, method)
- Methods insert with parentheses

#### Task 2.4: Create extension entry point

**Create `src/extension.ts`:**

```typescript
import * as vscode from 'vscode';
import * as path from 'path';
import { ApiLoader } from './apiLoader';
import { ComCompletionProvider } from './completionProvider';

export function activate(context: vscode.ExtensionContext) {
    console.log('comsense extension activated');

    // Load COM APIs
    const dataDir = path.join(context.extensionPath, 'data');
    const loader = new ApiLoader(dataDir);
    loader.load();

    const apis = loader.getAll();
    console.log(`Loaded ${apis.length} COM APIs`);

    // Register completion provider for VBA/VB files
    const provider = new ComCompletionProvider(apis);
    
    const disposable = vscode.languages.registerCompletionItemProvider(
        { language: 'vb', scheme: 'file' },
        provider,
        '.'  // Trigger on dot
    );

    context.subscriptions.push(disposable);
}

export function deactivate() {}
```

**Definition of Done:**
- Extension activates on VBA files
- APIs loaded at activation
- Completion provider registered
- Triggers on `.` character

#### Task 2.5: Test extension locally

**Compile extension:**

```bash
cd vscode-extension
npm run compile
```

**Test in VSCode:**

1. Open `vscode-extension` folder in VSCode
2. Press `F5` to launch Extension Development Host
3. Create a test file `test.vba`:
   ```vba
   Sub TestCompletion()
       Dim app As Application
       app.  ' <-- Trigger completion here
   End Sub
   ```
4. Verify completion suggestions appear

**Definition of Done:**
- Extension compiles without errors
- Extension runs in development mode
- Typing `Application.` shows completions
- Completions show Visio methods/properties

### 4.4 Phase 2 Deliverables

- [ ] VSCode extension project structure
- [ ] Completion provider implementation
- [ ] Working extension in development mode
- [ ] Test file demonstrating completions

### 4.5 Phase 2 Time Estimate

- Task 2.1: 2 hours
- Task 2.2: 3 hours
- Task 2.3: 6 hours
- Task 2.4: 2 hours
- Task 2.5: 3 hours
- **Total:** ~16 hours (2 days with debugging)

---

## 5. Phase 3: Validation & Documentation (Days 9-10)

### 5.1 Goal

Verify the MVP works with real VBA code and document findings.

### 5.2 Validation Tasks

#### Task 3.1: Test with real Visio VBA code

**Create test file `test-visio.vba`:**

```vba
Sub TestVisioCompletion()
    Dim app As Visio.Application
    Set app = CreateObject("Visio.Application")
    
    ' Test completion after dot
    app.ActiveDocument.
    
    Dim doc As Visio.Document
    Set doc = app.Documents.Add("")
    
    doc.Pages.
    
    Dim page As Visio.Page
    Set page = doc.Pages(1)
    
    page.DrawRectangle.
    
End Sub
```

**Validation checklist:**

- [ ] `app.` shows Application methods (ActiveDocument, Documents, Quit, etc.)
- [ ] `doc.` shows Document methods (Pages, Save, Close, etc.)
- [ ] `page.` shows Page methods (DrawRectangle, Drop, etc.)
- [ ] Completions appear within 500ms
- [ ] No errors in extension console

#### Task 3.2: Document what works and what doesn't

**Create `VALIDATION.md`:**

```markdown
# MVP Validation Report

**Date:** 2025-11-XX  
**Test environment:** Windows 11, VSCode 1.XX, Visio 2019/365

## What Works ✓

- Class name completion: Application, Document, Page, Shape
- Property completion after dot: app.ActiveDocument, doc.Pages
- Method completion after dot: page.DrawRectangle, shape.Delete
- Completion performance: <500ms
- Multiple COM libraries loaded simultaneously

## What Doesn't Work ✗

- Parameter hints for methods (not implemented)
- Hover documentation (not implemented)
- Type inference (e.g., `Set doc = app.Documents.Add` doesn't infer doc type)
- Property/method descriptions (extraction script doesn't capture them)
- Enum values (not extracted)

## Missing Features (Known Limitations)

- No "Go to Definition" support
- No signature help
- No syntax highlighting improvements
- No error checking/linting
- Limited to manually extracted COM libraries

## Performance

- Extension activation time: ~100ms
- API loading time: ~50ms (2 APIs)
- Completion response time: ~200ms
- Memory usage: ~15MB

## Usability Observations

[Note findings about user experience, friction points, confusing behavior]

## Next Steps

Based on this validation:

1. **High priority:** [Feature/fix that would have most impact]
2. **Medium priority:** [Nice-to-have improvements]
3. **Low priority:** [Polish items]

## Decision: Continue or Pivot?

[✓ Continue with production version] / [✗ Needs major rework] / [? Unclear - need more testing]

Rationale: [Explain based on validation results]
```

#### Task 3.3: Update README with MVP instructions

**Update `README.md`:**

```markdown
# comsense

> COM Type Library extraction for VSCode IntelliSense

**Status:** MVP - Proof of Concept ✨

## What is this?

A tool to extract COM type libraries (like Visio, Excel, Word) into JSON format and use them for IntelliSense in VSCode when writing VBA code.

**This is an MVP:** It works, but it's minimal. See [VALIDATION.md](VALIDATION.md) for what works and what doesn't.

## Quick Start

### 1. Extract a COM Library

```bash
# Requires: Windows, Python 3.10+, pywin32
pip install pywin32

# Extract Visio API
python extract_com.py "Visio.Application" "examples/visio-api.json"

# Extract any COM library
python extract_com.py "Excel.Application" "examples/excel-api.json"
```

### 2. Install VSCode Extension (Local)

```bash
cd vscode-extension
npm install
npm run compile

# Install extension locally
code --install-extension .
```

### 3. Use IntelliSense

Create a `.vba` file and start typing:

```vba
Dim app As Application
app. ' <-- IntelliSense appears here!
```

## What Works

✓ Extracts classes, properties, and methods from COM libraries  
✓ Provides completion suggestions in VSCode  
✓ Works with multiple COM libraries simultaneously  
✓ Fast (<500ms completion response)  

## What Doesn't Work (Yet)

✗ Parameter hints  
✗ Hover documentation  
✗ Type inference  
✗ Enum values  
✗ Go to definition  

See [VALIDATION.md](VALIDATION.md) for full details.

## Project Structure

```
comsense/
├── extract_com.py          # Extraction script
├── examples/               # Extracted JSON files
├── vscode-extension/       # VSCode extension
└── PROJECT_PLAN_MVP.md     # This implementation plan
```

## Next Steps

If the MVP proves useful:

1. Build full CLI tool (see [PROJECT_PLAN.md](PROJECT_PLAN.md))
2. Add JSON schema validation
3. Create community repository for shared extractions
4. Publish extension to marketplace

## License

MIT
```

### 5.3 Phase 3 Deliverables

- [ ] `VALIDATION.md` with test results
- [ ] Updated `README.md` with MVP instructions
- [ ] Decision on whether to proceed to production
- [ ] List of priority improvements

### 5.4 Phase 3 Time Estimate

- Task 3.1: 3 hours
- Task 3.2: 2 hours
- Task 3.3: 1 hour
- **Total:** ~6 hours (1 day)

---

## 6. Coding Agent Execution Guidelines

### 6.1 Execution Order

**CRITICAL:** Execute phases sequentially. Do not proceed to Phase 2 until Phase 1 deliverables are complete and tested.

**Phase 1 → Phase 2 gate:**
- `extract_com.py` must successfully extract at least one COM library
- JSON output must be valid and contain classes with properties/methods

**Phase 2 → Phase 3 gate:**
- Extension must compile without errors
- Extension must activate in VSCode development mode
- At least one completion suggestion must appear in test file

### 6.2 Task Execution Protocol

For each task:

1. **Read task description** completely before starting
2. **Create files** as specified in task
3. **Test immediately** after implementation
4. **Fix errors** before proceeding to next task
5. **Commit** when task Definition of Done is met

**Commit message format:**

```
feat(phase-N): <task description>

<what was implemented>
<any issues encountered>

Task: X.Y from PROJECT_PLAN_MVP.md
```

### 6.3 Error Handling

**If a task fails:**

1. **Check prerequisites:** Are all dependencies installed?
2. **Check platform:** Are you on Windows (for Phase 1)?
3. **Check file paths:** Are all paths correct?
4. **Log the error:** Include full stack trace in commit message
5. **Document workaround:** If you can't fix it, document the issue

**Unblocking strategies:**

- **Missing COM library:** Use `Scripting.FileSystemObject` instead of Visio
- **Extraction fails:** Start with minimal extraction (just class names)
- **Extension won't compile:** Check Node.js version (18+ required)
- **Completions don't appear:** Verify file extension is `.vba` or `.bas`

### 6.4 Testing Requirements

**Phase 1 testing:**

```bash
# Must pass
python extract_com.py "Scripting.FileSystemObject" "test-output.json"
cat test-output.json | grep '"classes"'  # Should find classes object

# Nice to have
python extract_com.py "Visio.Application" "test-visio.json"
```

**Phase 2 testing:**

```bash
# Must pass
cd vscode-extension
npm run compile  # No errors

# Manual test in VSCode
# 1. Press F5
# 2. Create test.vba
# 3. Type "Application." and verify completions appear
```

### 6.5 Success Criteria

**MVP is complete when:**

- [ ] `extract_com.py` exists and extracts at least one COM library
- [ ] `examples/` contains at least one valid JSON file
- [ ] VSCode extension compiles without errors
- [ ] Extension provides completions in development mode
- [ ] `VALIDATION.md` documents what works and what doesn't
- [ ] `README.md` has clear instructions for using the MVP

**Then you can declare:** "MVP validated - ready for production planning"

### 6.6 What NOT to Do

**Do NOT:**

- ❌ Add test frameworks (pytest, jest) - manual testing only
- ❌ Create `comsense/` package structure - single file only
- ❌ Implement JSON schema validation - not needed for MVP
- ❌ Build a CLI with multiple commands - single script only
- ❌ Optimize performance - make it work first
- ❌ Add comprehensive error handling - basic try/catch is enough
- ❌ Write extensive documentation - README + VALIDATION.md only
- ❌ Implement all VT type mappings - use "variant" for everything

**Remember:** This is a proof of concept. Speed over perfection.

---

## 7. Transition to Production

### 7.1 When to Transition

**Transition to full implementation ([PROJECT_PLAN.md](PROJECT_PLAN.md)) when:**

1. MVP validation shows the approach works
2. You've used the extension yourself for 1+ week
3. You have specific pain points to address ("I wish it had X")
4. At least 2-3 people express interest in using it

**Don't transition if:**

- Completions are too slow or unreliable
- JSON format doesn't capture important information
- VSCode IntelliSense approach doesn't work well for VBA
- You discover a better alternative solution

### 7.2 Transition Checklist

**Before starting production implementation:**

- [ ] Document all MVP learnings in `VALIDATION.md`
- [ ] List specific improvements needed (prioritized)
- [ ] Review [PROJECT_PLAN.md](PROJECT_PLAN.md) and adjust based on MVP findings
- [ ] Decide on JSON schema format (informed by what VSCode actually needs)
- [ ] Set up GitHub Issues for production features
- [ ] Create `develop` branch for production work
- [ ] Archive MVP code in `mvp-archive/` branch

### 7.3 Production Plan Adjustments

Based on MVP results, adjust [PROJECT_PLAN.md](PROJECT_PLAN.md):

**If extraction was easy:** Keep Milestone 3 as-is  
**If extraction was hard:** Add more error handling to Milestone 3  

**If VSCode integration was easy:** Simplify future extension plans  
**If VSCode integration was hard:** Add more research/prototyping phases  

**If JSON format worked well:** Lock down schema early (Milestone 4)  
**If JSON format needs changes:** Iterate on schema before committing  

---

## 8. Timeline & Milestones

### 8.1 Realistic Timeline

**Full-time (8 hours/day):**
- Phase 1: Days 1-2 (12 hours)
- Phase 2: Days 3-4 (16 hours)
- Phase 3: Day 5 (6 hours)
- **Total:** 5 working days

**Part-time (2-3 hours/day):**
- Phase 1: Days 1-4
- Phase 2: Days 5-9
- Phase 3: Days 10-11
- **Total:** 11 calendar days

**Agent execution (continuous):**
- Phase 1: 4-6 hours
- Phase 2: 8-12 hours
- Phase 3: 2-3 hours
- **Total:** 14-21 hours

### 8.2 Milestone Markers

**🎯 Milestone 1:** Extraction script works (Day 2)
**🎯 Milestone 2:** Extension shows completions (Day 5)
**🎯 Milestone 3:** MVP validated (Day 10)

---

## 9. Risk Management

### 9.1 Known Risks

| Risk | Likelihood | Impact | Mitigation |
|------|-----------|--------|------------|
| COM extraction fails on some libraries | High | Medium | Use Scripting.FileSystemObject as fallback |
| VSCode completions don't appear | Medium | High | Test with simple JSON first |
| Generated JSON is too large | Low | Medium | Start with small COM libraries |
| VBA file extension not recognized | Low | Low | Add multiple extensions to package.json |
| Extraction is too slow | Medium | Low | Acceptable for MVP |

### 9.2 Showstopper Scenarios

**If this happens, STOP and reassess:**

1. **Cannot extract any COM library** - pywin32 not working
   - Solution: Check Python/pywin32 installation
   
2. **VSCode completions never appear** - fundamental issue
   - Solution: Review VSCode API documentation, check language ID
   
3. **Extracted JSON doesn't contain useful information** - wrong approach
   - Solution: Research COM reflection APIs more

---

## 10. Success Metrics

### 10.1 MVP Success Criteria

**Technical:**
- [ ] Extract at least 1 COM library successfully
- [ ] VSCode extension activates without errors
- [ ] Completions appear within 1 second
- [ ] At least 50% of expected properties/methods are extracted

**Usability:**
- [ ] You would use this yourself when writing VBA
- [ ] Completions are accurate (few false positives)
- [ ] Setup takes less than 30 minutes

**Validation:**
- [ ] Tested with real VBA code (not just hello world)
- [ ] Works with at least 2 different COM libraries
- [ ] Documented decision to proceed or pivot

### 10.2 What "Success" Looks Like

**After 10 days, you should have:**

1. A Python script that extracts COM APIs to JSON
2. A VSCode extension that uses those JSONs for IntelliSense
3. Evidence that the approach works with real code
4. A clear decision: build production version or try different approach

**You do NOT need:**

- Perfect extraction (60-70% coverage is fine)
- Beautiful code (hacky MVP is fine)
- Tests (manual validation is fine)
- Documentation (README + validation report is fine)

---

## 11. Appendix

### 11.1 Useful Commands

**Python:**
```bash
# Install dependencies
pip install pywin32

# Run extraction
python extract_com.py "<ProgID>" "output.json"

# List available COM libraries (manual)
python -c "import win32com.client; print(dir(win32com.client.gencache))"
```

**VSCode Extension:**
```bash
# Setup
cd vscode-extension
npm install

# Compile
npm run compile

# Watch mode (auto-recompile)
npm run watch

# Test in VSCode
# Press F5 in VSCode with extension folder open
```

### 11.2 Common ProgIDs

**Always available:**
- `Scripting.FileSystemObject`
- `WScript.Shell`

**Microsoft Office (if installed):**
- `Visio.Application`
- `Excel.Application`
- `Word.Application`
- `PowerPoint.Application`
- `Outlook.Application`

**Other common:**
- `InternetExplorer.Application`
- `Shell.Application`

### 11.3 Troubleshooting

**"Module not found" when importing win32com:**
```bash
pip install --upgrade pywin32
python -c "import win32com; print('OK')"
```

**"com_error: (-2147221005, 'Invalid class string', None, None)":**
- ProgID doesn't exist or isn't registered
- Try `Scripting.FileSystemObject` instead

**VSCode extension doesn't activate:**
- Check file extension is `.vba` or `.bas`
- Check Output panel → comsense for errors
- Verify `activationEvents` in package.json

**Completions don't appear:**
- Type slowly and wait 1-2 seconds
- Check the dot is recognized as trigger character
- Verify JSON files are in `data/apis/` directory
- Check extension console for errors

---

## 12. Changelog

**2025-11-18:** Initial MVP plan created
- Defined 3-phase approach (10 days)
- Focus on speed-to-validation
- Simplified from full production plan

---

**End of MVP Plan** - Let's build something that works! 🚀
