```markdown
# comsense – Project Implementation Plan

**Status:** Planning  
**Primary audience:** Autonomous coding agents and human contributors  
**Language:** English only  
**Version:** 1.0.0  
**Last updated:** 2025-11-18

---

## 1. Project Overview

### 1.1 Vision

`comsense` is a **COM type library extraction tool** that:

- Discovers COM type libraries installed on a Windows system.  
- Lets the user choose a library (e.g. `Visio.Application`, `Excel.Application`).  
- Extracts its public API (classes, properties, methods, enums) into a **standardized JSON format**.  
- Produces JSON files that can be consumed by tools such as **VS Code extensions** for IntelliSense / code completion.  
- Enables a **community-driven JSON library**: users can contribute extracted JSON files to a shared repository.

### 1.2 Scope (v0.x)

The initial project focus (this repository `comsense`) covers:

- A **Python library** that:
  - Scans Windows registry for COM ProgIDs and type libraries.
  - Extracts type info for a selected COM library using `pywin32` / `win32com`.
  - Normalizes the extracted information into a stable JSON schema.

- A **CLI interface** that exposes:
  - `comsense list` – list available COM ProgIDs/type libraries.
  - `comsense info <prog_id>` – show detailed info about a ProgID.
  - `comsense extract <prog_id> [...]` – extract API to JSON file.
  - `comsense validate <file>` – validate JSON files against the schema.

Out of scope for this repo (but to be supported by its output):

- VS Code extension implementation (should live in a separate repo).
- Hosting the shared JSON library (should be a separate data repo, e.g. `comsense-data`).

---

## 2. High-Level Architecture

### 2.1 Components

1. **Core library (`comsense` package)**  
   - `registry` module: discovers COM ProgIDs / CLSIDs from the Windows registry.  
   - `extractor` module: uses `win32com.client` to introspect type libraries.  
   - `model` module: internal representation of COM types (classes, interfaces, methods, properties, enums).  
   - `serializer` module: converts the internal model to/from JSON per schema.  
   - `validator` module: validates JSON files using a JSON schema.  

2. **CLI (`comsense.cli`)**
   - Built on `click` or `argparse`.
   - Provides commands for listing, inspecting, extracting, validating.

3. **Schemas (`schemas/`)**
   - JSON Schema describing the standardized output format.
   - Used by both `validator` and external tools.

4. **Tests (`tests/`)**
   - Unit tests for registry scanning, extraction, serialization, validation, CLI.

### 2.2 Target Platform & Dependencies

- **OS:** Windows only (COM / registry dependent).  
- **Runtime:** Python 3.10+ (exact min version: `>=3.10,<4.0`).  
- **Main dependencies:**
  - `pywin32>=305` (`win32com.client`) – COM interaction.
  - `click>=8.0` – CLI framework.
  - `jsonschema>=4.17` – JSON Schema validation.
  - `pytest>=7.0` – test framework.
  - `pytest-mock>=3.10` – mocking for tests.

### 2.3 Development Dependencies

- `black>=23.0` – code formatting.
- `mypy>=1.0` – static type checking.
- `pytest-cov>=4.0` – test coverage reporting.

---

## 3. Repository Structure

The repository should follow this structure:

```
comsense/
├── comsense/
│   ├── __init__.py
│   ├── cli.py
│   ├── registry.py
│   ├── extractor.py
│   ├── model.py
│   ├── serializer.py
│   ├── validator.py
│   └── logging_config.py
├── schemas/
│   └── typelib-schema.v1.json
├── tests/
│   ├── __init__.py
│   ├── fixtures/
│   │   ├── valid_simple.json
│   │   ├── valid_complex.json
│   │   └── invalid_missing_metadata.json
│   ├── test_registry.py
│   ├── test_extractor.py
│   ├── test_serializer.py
│   ├── test_validator.py
│   └── test_cli.py
├── examples/
│   ├── visio-application.json
│   ├── excel-application.json
│   └── README.md
├── docs/
│   ├── ARCHITECTURE.md
│   ├── CHECKLIST.md
│   ├── ERROR_HANDLING.md
│   └── CONTRIBUTING.md
├── pyproject.toml
├── README.md
├── PROJECT_PLAN.md (this file)
├── LICENSE
└── .gitignore
```

---

## 4. JSON Schema Design (Output Format)

### 4.1 Goals

- Stable, tool-friendly format for code completion engines.  
- Language-agnostic representation of COM type info.  
- Human-readable but structured.
- Versioned schema with explicit migration paths.

### 4.2 Draft Schema (Conceptual)

Each extracted file corresponds to a **COM ProgID** (e.g. `Visio.Application`):

```
{
  "metadata": {
    "schema_version": "1.0.0",
    "generator": "comsense",
    "generator_version": "0.1.0",
    "prog_id": "Visio.Application",
    "clsid": "{00021A98-0000-0000-C000-000000000046}",
    "timestamp_utc": "2025-01-01T12:00:00Z",
    "warnings": []
  },
  "classes": {
    "Application": {
      "kind": "class",
      "documentation": "Optional description, may be empty.",
      "members": {
        "properties": {
          "ActiveDocument": {
            "type": "Document",
            "readonly": true,
            "documentation": "Returns the active document."
          },
          "Version": {
            "type": "string",
            "readonly": true,
            "documentation": "Returns the application version."
          }
        },
        "methods": {
          "Quit": {
            "return_type": "void",
            "parameters": [],
            "documentation": "Closes all documents and exits the application."
          },
          "OpenEx": {
            "return_type": "Document",
            "parameters": [
              {
                "name": "FileName",
                "type": "string",
                "optional": false,
                "by_ref": false
              },
              {
                "name": "Flags",
                "type": "int",
                "optional": true,
                "default": 0
              }
            ],
            "documentation": "Opens a document with options."
          }
        }
      }
    }
  },
  "enums": {
    "VisOpenSaveArgs": {
      "documentation": "Visio open/save flags.",
      "values": {
        "visOpenHidden": 64,
        "visOpenRO": 2,
        "visOpenDocked": 4
      }
    }
  }
}
```

### 4.3 Schema Versioning Strategy

- `typelib-schema.v1.json` is **immutable** once published to production.
- Breaking changes require `v2.json` with a migration guide in `docs/MIGRATIONS.md`.
- The `schema_version` field in output JSON must match the schema used.
- The validator must support validation against specific schema versions.

### 4.4 Type Mapping Rules

**VT_VARIANT types map to logical types:**

- `VT_BSTR` → `"string"`
- `VT_I4`, `VT_I2` → `"int"`
- `VT_R8`, `VT_R4` → `"float"`
- `VT_BOOL` → `"boolean"`
- `VT_DISPATCH` → `"object"` (with class name if available)
- `VT_VARIANT` → `"any"`
- Unmapped types → `"unknown"` + add warning to `metadata.warnings`

---

## 5. Documentation Strategy

### 5.1 Documentation-as-Code Philosophy

Following industry best practices for autonomous agent development, **all documentation lives in the repository** and evolves with the code.

### 5.2 Core Documentation Files

**Must be created and maintained:**

1. **README.md** (root)
   - Quick start guide
   - Installation instructions
   - Basic usage examples
   - Link to detailed docs

2. **docs/ARCHITECTURE.md**
   - System design overview
   - Module responsibilities
   - Data flow diagrams (text-based)
   - Technology choices and rationale

3. **docs/CHECKLIST.md**
   - Milestone completion tracking
   - Feature implementation status
   - Known issues and workarounds
   - Next steps

4. **docs/ERROR_HANDLING.md**
   - Common HRESULT error codes
   - Exception handling patterns
   - Debugging procedures
   - Platform-specific issues

5. **docs/CONTRIBUTING.md**
   - Code style guidelines
   - Testing requirements
   - PR process
   - Agent execution guidelines

6. **PROJECT_PLAN.md** (this file)
   - Master implementation plan
   - Updated as milestones complete

### 5.3 Documentation Update Protocol

**CRITICAL:** Documentation must stay synchronized with code at all times.

**Every milestone completion requires:**

1. Update `docs/CHECKLIST.md` to mark completed tasks.
2. Update `docs/ARCHITECTURE.md` if structure changed.
3. Add new error codes to `docs/ERROR_HANDLING.md` if discovered.
4. Update `README.md` examples if CLI changed.
5. Add changelog entry to `PROJECT_PLAN.md` section 10.

**Every commit that changes public APIs must:**

- Update relevant docstrings.
- Update CLI help text if applicable.
- Update example code in `examples/README.md`.

**Weekly documentation review checklist:**

- [ ] All modules have up-to-date docstrings
- [ ] CLI help text matches implementation
- [ ] Example JSON files validate against current schema
- [ ] CHECKLIST.md reflects current milestone status
- [ ] No TODOs in code without corresponding CHECKLIST items

### 5.4 Test-Case-Driven Documentation

Tests serve as **executable documentation**. Every test should:

- Have a descriptive name: `test_extract_handles_missing_type_library_gracefully`
- Include docstring explaining scenario and expected behavior
- Use explicit assertions with failure messages
- Reference relevant documentation sections in comments

---

## 6. Error Handling & Edge Cases

### 6.1 Expected COM Exceptions

**The extractor must handle these HRESULT codes:**

| HRESULT | Constant | Meaning | Handler Action |
|---------|----------|---------|----------------|
| `0x80040154` | `REGDB_E_CLASSNOTREG` | ProgID not registered | Raise `COMNotRegisteredError` with ProgID |
| `0x80029C4A` | `TYPE_E_CANTLOADLIBRARY` | Type library DLL missing | Raise `TypeLibraryLoadError` with path |
| `0x8002801D` | `TYPE_E_LIBNOTREGISTERED` | Type library not registered | Raise `TypeLibraryNotRegisteredError` |
| `0x80020009` | `DISP_E_EXCEPTION` | Exception in COM method | Log full traceback, raise `COMExtractionError` |
| `0x80070005` | `E_ACCESSDENIED` | Permission denied | Raise `PermissionError` with elevation hint |

### 6.2 Edge Case Decision Tree

**If `win32com.client.gencache.EnsureDispatch()` returns object but extraction yields empty **

1. Check if `obj.__class__.__module__` exists → if not, log warning "No generated wrapper found"
2. Try accessing `win32com.client.constants` for enum extraction
3. Produce valid JSON with empty `classes` and `enums` objects
4. Add warning to `metadata.warnings`: `"No type information extracted - check if type library supports ITypeLib"`

**If type mapping is ambiguous:**

- Default `VT_VARIANT` to `"any"`
- Log at DEBUG level: `"Unmapped type VT_<code> for <member> - using 'any'"`
- Add to `metadata.warnings`: `"Type mapping incomplete for <N> members"`

**If multiple versions of same ProgID exist (e.g. Excel.Application.16, Excel.Application.15):**

- `comsense list` shows all versions
- `comsense extract` requires exact ProgID
- Agent prefers latest version when ambiguous

### 6.3 Platform-Specific Constraints

**Windows Registry Access:**

- Requires read access to `HKEY_LOCAL_MACHINE\SOFTWARE\Classes`
- No special privileges needed for reading
- If registry key missing, return empty list (not an error)

**32-bit vs 64-bit COM:**

- Python bitness must match COM server bitness
- Add `--prefer-32bit` / `--prefer-64bit` flags to CLI (future)
- Document in README.md under "Known Limitations"

---

## 7. Testing Strategy

### 7.1 Test Categories

**Unit tests (must run on any platform with mocks):**

- Mark with `@pytest.mark.unit`
- Mock all `win32com` calls
- Focus on logic: parsing, serialization, validation

**Integration tests (require real Windows + COM):**

- Mark with `@pytest.mark.windows_only`
- Skip automatically on non-Windows: `@pytest.mark.skipif(sys.platform != "win32", reason="Windows only")`
- Test with real COM libraries if available
- Gracefully skip if test ProgIDs not installed

**Example:**

```
import sys
import pytest

@pytest.mark.unit
def test_serializer_converts_model_to_json():
    """Unit test - runs anywhere with mocks"""
    model = MockTypeLibModel()
    json_dict = to_json_dict(model)
    assert json_dict["metadata"]["schema_version"] == "1.0.0"

@pytest.mark.windows_only
@pytest.mark.skipif(sys.platform != "win32", reason="Requires Windows")
def test_extract_visio_application():
    """Integration test - requires real Visio installation"""
    try:
        result = TypeLibraryExtractor().extract("Visio.Application")
        assert "Application" in result.classes
    except COMNotRegisteredError:
        pytest.skip("Visio not installed")
```

### 7.2 Test Data Management

- **Mock ** `tests/fixtures/mock_com_objects.py`
- **Valid JSON:** `tests/fixtures/valid_*.json`
- **Invalid JSON:** `tests/fixtures/invalid_*.json`
- **Real extractions:** `examples/*.json` (committed, validated)

**Do NOT commit:**

- Test extractions that change frequently
- Large (>100KB) JSON files to `tests/fixtures/`

### 7.3 CI/CD Configuration

**GitHub Actions workflow (`.github/workflows/test.yml`):**

```
name: Test Suite

on: [push, pull_request]

jobs:
  unit-tests:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      - uses: actions/setup-python@v4
        with:
          python-version: '3.10'
      - run: pip install -e ".[dev]"
      - run: pytest -m unit --cov=comsense

  integration-tests:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v3
      - uses: actions/setup-python@v4
        with:
          python-version: '3.10'
      - run: pip install -e ".[dev]"
      - run: pytest -m windows_only
```

---

## 8. Git Workflow & Branch Strategy

### 8.1 Branch Model

- **`main`**: Stable releases only, protected branch
- **`develop`**: Integration branch for completed milestones
- **`feature/milestone-N`**: One feature branch per milestone
- **`bugfix/<issue-number>`**: Hotfixes and bug corrections

### 8.2 Commit Message Format

```
<type>(<scope>): <short summary>

<optional detailed description>

Refs: #<issue-number>
```

**Types:** `feat`, `fix`, `docs`, `test`, `refactor`, `chore`

**Examples:**

```
feat(extractor): add support for VT_ARRAY parameter types

Handles safe array types by recursing into element types.

Refs: #12

***

docs(checklist): mark Milestone 3 as complete

All tasks in M3 finished, tests passing.
```

### 8.3 Pull Request Requirements

- All tests passing (unit + applicable integration)
- Documentation updated (CHECKLIST.md, relevant docs/)
- Code formatted with `black`
- Type hints validated with `mypy`
- At least one example updated if CLI changed

---

## 9. Milestones & Tasks (for Coding Agent)

Below, each milestone lists:

- **Goal**
- **Tasks** (atomic, implementable)
- **Files to touch**
- **Definition of Done (DoD)**
- **Documentation updates required**

---

### Milestone 1: Repository Bootstrap

**Goal:** Make the repo installable as a bare-bones Python package with a placeholder CLI.

**Tasks:**

1. **Initialize Python package structure**
   - Create `comsense/__init__.py` with basic metadata (`__version__ = "0.1.0"`).
   - Create empty modules: `cli.py`, `registry.py`, `extractor.py`, `model.py`, `serializer.py`, `validator.py`, `logging_config.py`.

2. **Configure packaging**
   - Add `pyproject.toml` with:
     - Project meta name, version, description, authors, license
     - Dependencies: `pywin32>=305`, `click>=8.0`, `jsonschema>=4.17`
     - Dev dependencies: `pytest>=7.0`, `pytest-mock>=3.10`, `black>=23.0`, `mypy>=1.0`, `pytest-cov>=4.0`
     - Python version constraint: `requires-python = ">=3.10,<4.0"`
     - Console script entry point: `comsense = "comsense.cli:main"`

3. **Add basic files**
   - `README.md` with project description and placeholder usage.
   - `LICENSE` (MIT recommended).
   - `.gitignore` (Python defaults: `__pycache__/`, `*.pyc`, `.pytest_cache/`, `dist/`, `*.egg-info/`).

4. **Implement placeholder CLI**
   - In `cli.py`, implement a `main()` function using `click`.
   - Commands: `list`, `info`, `extract`, `validate` (all print "Not implemented yet").
   - Ensure `python -m comsense --version` prints version.

5. **Create initial documentation structure**
   - Create `docs/` directory
   - Create `docs/ARCHITECTURE.md` with placeholder sections
   - Create `docs/CHECKLIST.md` with Milestone 1 tasks
   - Create `docs/ERROR_HANDLING.md` with placeholder
   - Create `docs/CONTRIBUTING.md` with placeholder

**Files to create/modify:**

- `comsense/__init__.py`
- `comsense/cli.py`
- `comsense/logging_config.py` (placeholder)
- `pyproject.toml`
- `README.md`
- `LICENSE`
- `.gitignore`
- `docs/ARCHITECTURE.md`
- `docs/CHECKLIST.md`
- `docs/ERROR_HANDLING.md`
- `docs/CONTRIBUTING.md`
- `PROJECT_PLAN.md` (add this file)

**Definition of Done:**

- [x] `pip install -e .` works without errors
- [x] Running `comsense --version` prints `0.1.0`
- [x] Running `comsense --help` shows all four commands
- [x] No import errors when running `python -c "import comsense"`
- [x] Git repository initialized with `.gitignore`
- [x] All documentation files created with basic structure

**Documentation updates required:**

- Update `docs/CHECKLIST.md`: Mark Milestone 1 tasks as complete
- Update `docs/ARCHITECTURE.md`: Document package structure
- Update `README.md`: Add installation section

---

### Milestone 2: Registry Scanner

**Goal:** Implement discovery of COM ProgIDs and associated type library information.

**Tasks:**

1. **Implement `COMTypeLibrary` dataclass in `registry.py`**
   - Fields: `prog_id: str`, `clsid: str`, `description: str`, `version: str`, `path: Optional[str]`
   - Add type hints and docstrings

2. **Implement `RegistryScanner` class in `registry.py`**
   - Method: `list_type_libraries(filter_pattern: Optional[str] = None) -> List[COMTypeLibrary]`
   - Scan under `HKEY_LOCAL_MACHINE\SOFTWARE\Classes` for ProgIDs
   - Recognize ProgID patterns (e.g. `*.Application`, `*.Document`)
   - Resolve `CLSID` subkey via `ProgID\CLSID` registry key
   - Resolve server path from `CLSID\{clsid}\InprocServer32` or `LocalServer32`
   - Implement filtering using `fnmatch` for shell-style wildcards
   - Add proper exception handling for missing registry keys
   - Log progress at DEBUG level: "Scanning registry key: <key>"

3. **Configure logging in `logging_config.py`**
   - Function: `configure_logging(debug: bool = False)`
   - Format: `[%(levelname)s] %(module)s: %(message)s`
   - Levels:
     - `INFO` (default): High-level progress messages
     - `DEBUG` (with `--debug` flag): Detailed registry scanning, each key visited
     - `ERROR`: Exceptions and failures

4. **Expose scanner via CLI**
   - In `cli.py`, implement `comsense list`:
     - Options: `--filter/-f` (optional pattern), `--debug` (enable debug logging)
     - Output: Table with columns `ProgID`, `Description`, `Version`
     - Use `click.echo()` for output, `tabulate` or manual formatting
     - Handle `PermissionError` gracefully with helpful message

5. **Create unit tests in `tests/test_registry.py`**
   - Mock `winreg.OpenKey`, `winreg.QueryValueEx` using `pytest-mock`
   - Test cases:
     - `test_list_all_type_libraries()`: Returns non-empty list
     - `test_filter_by_pattern()`: Filters work correctly
     - `test_missing_clsid_key()`: Handles missing CLSID gracefully
     - `test_permission_denied()`: Raises appropriate exception
   - Mark tests with `@pytest.mark.unit` (mocked, runs anywhere)

**Files to create/modify:**

- `comsense/registry.py`
- `comsense/logging_config.py`
- `comsense/cli.py`
- `tests/test_registry.py`

**Definition of Done:**

- [x] `comsense list` runs on Windows without errors
- [x] `comsense list --filter "*.Application"` shows only matching ProgIDs
- [x] `comsense list --debug` shows detailed registry scanning logs
- [x] All unit tests pass: `pytest -m unit tests/test_registry.py`
- [x] Type hints validated with `mypy comsense/registry.py`
- [x] Code formatted with `black`

**Documentation updates required:**

- Update `docs/CHECKLIST.md`: Mark Milestone 2 complete
- Update `docs/ARCHITECTURE.md`: Document `RegistryScanner` class and registry structure
- Update `README.md`: Add example of `comsense list` command
- Update `docs/ERROR_HANDLING.md`: Add registry access permission issues

---

### Milestone 3: Type Library Extraction Core

**Goal:** Extract COM type information for a selected ProgID using `win32com`.

**Tasks:**

1. **Implement internal model in `model.py`**
   - Dataclasses with type hints:
     - `TypeLibMetadata`: `schema_version`, `generator`, `generator_version`, `prog_id`, `clsid`, `timestamp_utc`, `warnings: List[str]`
     - `ComParameter`: `name`, `type`, `optional`, `default`, `by_ref`, `documentation`
     - `ComMethod`: `name`, `return_type`, `parameters: List[ComParameter]`, `documentation`
     - `ComProperty`: `name`, `type`, `readonly`, `documentation`
     - `ComClass`: `name`, `kind`, `documentation`, `properties: Dict[str, ComProperty]`, `methods: Dict[str, ComMethod]`
     - `ComEnum`: `name`, `documentation`, `values: Dict[str, int]`
     - `TypeLibrary`: `meta TypeLibMetadata`, `classes: Dict[str, ComClass]`, `enums: Dict[str, ComEnum]`
   - Add comprehensive docstrings to each class

2. **Implement `TypeLibraryExtractor` in `extractor.py`**
   - Public method: `extract(prog_id: str) -> TypeLibrary`
   - Steps:
     - Call `win32com.client.gencache.EnsureDispatch(prog_id)` to generate wrapper
     - Resolve module: `obj.__class__.__module__`
     - Import generated module dynamically
     - Iterate over module attributes to find COM classes
     - Identify COM classes via `_prop_map_get_`, `_prop_map_put_`, `_method_map_` attributes
     - Map `pywin32` structures to internal model:
       - Properties: iterate `_prop_map_get_` and `_prop_map_put_`
       - Methods: iterate `_method_map_`
       - Enums: check `win32com.client.constants`
     - Apply type mapping rules (see section 4.4)
     - Populate `metadata.warnings` for unmapped types
   - Logging:
     - INFO: "Extracting ProgID: <prog_id>"
     - INFO: "Found <N> classes, <M> enums"
     - DEBUG: "Extracting class: <class_name>"
     - DEBUG: "Found method: <method_name> with <N> parameters"

3. **Implement error handling**
   - Catch `pywintypes.com_error` and map HRESULT to custom exceptions:
     - `0x80040154` → `COMNotRegisteredError(prog_id)`
     - `0x80029C4A` → `TypeLibraryLoadError(prog_id, path)`
     - `0x8002801D` → `TypeLibraryNotRegisteredError(prog_id)`
     - Others → `COMExtractionError(prog_id, hresult, message)`
   - Define custom exceptions in `extractor.py`:
     ```
     class COMExtractionError(Exception): pass
     class COMNotRegisteredError(COMExtractionError): pass
     class TypeLibraryLoadError(COMExtractionError): pass
     class TypeLibraryNotRegisteredError(COMExtractionError): pass
     ```

4. **Create unit tests in `tests/test_extractor.py`**
   - Mock `win32com.client.gencache.EnsureDispatch` using `pytest-mock`
   - Create mock COM class structures in `tests/fixtures/mock_com_objects.py`:
     - `MockCOMClass` with `_prop_map_get_`, `_method_map_`
   - Test cases:
     - `test_extract_simple_class()`: Extracts properties and methods
     - `test_extract_handles_missing_wrapper()`: Raises appropriate error
     - `test_extract_unmapped_types_adds_warning()`: Populates `metadata.warnings`
     - `test_extract_enum_from_constants()`: Extracts enums correctly
   - Mark all as `@pytest.mark.unit` (fully mocked)

5. **Add integration test (optional, Windows-only)**
   - In `tests/test_extractor.py`:
     ```
     @pytest.mark.windows_only
     @pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
     def test_extract_real_scripting_filesystemobject():
         """Integration test with real COM library"""
         try:
             result = TypeLibraryExtractor().extract("Scripting.FileSystemObject")
             assert result.metadata.prog_id == "Scripting.FileSystemObject"
             assert len(result.classes) > 0
         except COMNotRegisteredError:
             pytest.skip("Scripting.FileSystemObject not available")
     ```

**Files to create/modify:**

- `comsense/model.py`
- `comsense/extractor.py`
- `comsense/logging_config.py` (logging integration)
- `tests/test_extractor.py`
- `tests/fixtures/mock_com_objects.py`

**Definition of Done:**

- [x] `TypeLibraryExtractor.extract("Some.ProgId")` returns structured `TypeLibrary` object
- [x] All custom exceptions defined and properly raised
- [x] Type mapping rules implemented per section 4.4
- [x] Warnings populated for unmapped types
- [x] All unit tests pass: `pytest -m unit tests/test_extractor.py`
- [x] Integration test passes on Windows (if applicable)
- [x] No hard-coded product-specific logic (Visio, Excel, etc.)

**Documentation updates required:**

- Update `docs/CHECKLIST.md`: Mark Milestone 3 complete
- Update `docs/ARCHITECTURE.md`: Document `TypeLibraryExtractor` class and extraction pipeline
- Update `docs/ERROR_HANDLING.md`: Add all custom exceptions and HRESULT mappings
- Update `README.md`: Add note about extraction process

---

### Milestone 4: JSON Serialization & Schema

**Goal:** Define a stable JSON schema for extracted data and implement serializer + validator.

**Tasks:**

1. **Define JSON Schema in `schemas/typelib-schema.v1.json`**
   - Use JSON Schema Draft 2020-12
   - Root schema: `$schema`, `$id`, `title`, `description`, `type: object`
   - Required properties: `metadata`, `classes`, `enums`
   - Define nested schemas:
     - `metadata`: All fields required except `warnings` (array)
     - `classes`: `patternProperties` for class names mapping to class schema
     - `class`: `kind`, `documentation`, `members` (properties, methods)
     - `property`: `type`, `readonly`, `documentation`
     - `method`: `return_type`, `parameters`, `documentation`
     - `parameter`: `name`, `type`, `optional`, `by_ref`, `default`
     - `enums`: `patternProperties` for enum names mapping to enum schema
     - `enum`: `documentation`, `values` (object with string keys, integer values)
   - Add schema examples in `$examples` annotation

2. **Implement serializer in `serializer.py`**
   - Function: `to_json_dict(type_lib: TypeLibrary) -> Dict[str, Any]`
     - Convert all dataclasses to dictionaries
     - Ensure datetime formatted as ISO 8601 UTC
     - Remove `None` values from optional fields
     - Preserve order: `metadata`, `classes`, `enums`
   - Function: `save_json( Dict[str, Any], path: Path, pretty: bool = True) -> None`
     - Write JSON with `indent=2` if `pretty=True`
     - Ensure UTF-8 encoding
     - Create parent directories if needed
   - Add type hints and docstrings

3. **Implement validator in `validator.py`**
   - Function: `validate_json( Dict[str, Any], schema_version: str = "1.0.0") -> None`
     - Load schema from `schemas/typelib-schema.v{schema_version}.json`
     - Use `jsonschema.validate(data, schema)`
     - Raise `ValidationError` with detailed path and message on failure
   - Function: `validate_json_file(path: Path) -> None`
     - Load JSON file
     - Detect `schema_version` from `metadata.schema_version`
     - Call `validate_json()` with appropriate schema
   - Add helpful error messages for common issues

4. **Add unit tests**
   - `tests/test_serializer.py`:
     - `test_to_json_dict_preserves_structure()`: Check keys and structure
     - `test_to_json_dict_removes_none_values()`: Optional fields handled
     - `test_save_json_creates_valid_file()`: File written correctly
     - `test_save_json_creates_parent_dirs()`: Directory creation works
   - `tests/test_validator.py`:
     - `test_validate_valid_json()`: Valid data passes
     - `test_validate_missing_metadata()`: Raises error
     - `test_validate_invalid_schema_version()`: Raises error
     - `test_validate_extra_fields_allowed()`: Extensibility check
   - All tests marked `@pytest.mark.unit`

5. **Add example files under `examples/`**
   - `examples/simple-mock.json`: Minimal valid example (manually created)
   - `examples/README.md`: Explains each example and intended use

**Files to create/modify:**

- `schemas/typelib-schema.v1.json`
- `comsense/serializer.py`
- `comsense/validator.py`
- `tests/test_serializer.py`
- `tests/test_validator.py`
- `tests/fixtures/valid_simple.json`
- `tests/fixtures/invalid_missing_metadata.json`
- `examples/simple-mock.json`
- `examples/README.md`

**Definition of Done:**

- [x] JSON schema validates all example files
- [x] `to_json_dict()` produces schema-compliant output
- [x] `validate_json()` detects invalid JSON with clear error messages
- [x] All unit tests pass: `pytest -m unit tests/test_serializer.py tests/test_validator.py`
- [x] Schema version field validated correctly
- [x] Code formatted and type-checked

**Documentation updates required:**

- Update `docs/CHECKLIST.md`: Mark Milestone 4 complete
- Update `docs/ARCHITECTURE.md`: Document serialization pipeline and schema versioning strategy
- Update `examples/README.md`: Describe example files
- Add `docs/MIGRATIONS.md`: Placeholder for future schema migrations

---

### Milestone 5: CLI – `info`, `extract`, `validate`

**Goal:** Expose extractor, serializer, and validator through the CLI.

**Tasks:**

1. **Implement `comsense info <prog_id>`**
   - Use `RegistryScanner` to retrieve details
   - Display:
     ```
     ProgID:      Visio.Application
     CLSID:       {00021A98-0000-0000-C000-000000000046}
     Description: Microsoft Visio Application
     Version:     16.0
     Path:        C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE
     ```
   - Handle missing ProgID gracefully: "ProgID not found: <prog_id>"

2. **Implement `comsense extract <prog_id>`**
   - Options:
     - `--output/-o <directory>`: Output directory (default: `./output`)
     - `--filename <name>`: Explicit filename (default: auto-generated)
     - `--debug`: Enable debug logging
     - `--pretty/--no-pretty`: Pretty-print JSON (default: pretty)
   - Filename convention: `<prog_id>` lowercased, dots → hyphens (e.g. `visio-application.json`)
   - Steps:
     1. Call `configure_logging(debug=<debug_flag>)`
     2. Run `TypeLibraryExtractor().extract(prog_id)`
     3. Convert to JSON dict: `to_json_dict(result)`
     4. Save to file: `save_json(data, output_path, pretty=<pretty_flag>)`
     5. Print summary:
        ```
        Extraction complete!
        Output: ./output/visio-application.json
        Classes: 42
        Enums: 8
        Warnings: 2 (see metadata.warnings in output)
        ```
   - Error handling:
     - Catch all `COMExtractionError` subclasses
     - Print user-friendly error message
     - Exit with code 1

3. **Implement `comsense validate <path>`**
   - Load JSON file from `<path>`
   - Run `validate_json_file(path)`
   - Output:
     - Success: `✓ Valid: <path>`
     - Failure: 
       ```
       ✗ Invalid: <path>
       Error at $.metadata.schema_version: '2.0.0' is not a valid version
       ```
   - Exit codes:
     - 0: Valid
     - 1: Invalid or file not found

4. **Extend CLI tests in `tests/test_cli.py`**
   - Use `click.testing.CliRunner`
   - Test cases:
     - `test_list_command()`: Mock registry scanner
     - `test_info_command_found()`: Shows info for known ProgID
     - `test_info_command_not_found()`: Handles missing ProgID
     - `test_extract_command_success()`: Mocks extractor, checks file created
     - `test_extract_command_handles_com_error()`: Error handling works
     - `test_validate_command_valid()`: Validates valid file
     - `test_validate_command_invalid()`: Detects invalid file
   - All tests marked `@pytest.mark.unit` (fully mocked)

**Files to create/modify:**

- `comsense/cli.py`
- `tests/test_cli.py`

**Definition of Done:**

- [x] `comsense info <known_prog_id>` prints meaningful info on Windows
- [x] `comsense extract <prog_id>` generates valid JSON file
- [x] `comsense validate <file>` correctly validates files
- [x] All CLI options work as specified
- [x] Error messages are user-friendly and actionable
- [x] All unit tests pass: `pytest -m unit tests/test_cli.py`
- [x] Help text is clear: `comsense --help`, `comsense extract --help`

**Documentation updates required:**

- Update `docs/CHECKLIST.md`: Mark Milestone 5 complete
- Update `README.md`: Add comprehensive CLI usage examples for all commands
- Update `docs/CONTRIBUTING.md`: Add section on testing CLI changes

---

### Milestone 6: Documentation & Examples

**Goal:** Make the project understandable and usable by others (and future agents).

**Tasks:**

1. **Update `README.md`**
   - Sections:
     - **Overview**: What `comsense` does, why it exists
     - **Installation**:
       ```
       pip install comsense
       ```
     - **Quick Start**:
       ```
       # List all COM libraries
       comsense list
       
       # Filter for Office apps
       comsense list --filter "*.Application"
       
       # Get info about a specific library
       comsense info "Visio.Application"
       
       # Extract to JSON
       comsense extract "Visio.Application" --output ./extractions
       
       # Validate JSON
       comsense validate ./extractions/visio-application.json
       ```
     - **Use Cases**: VS Code IntelliSense, API documentation, etc.
     - **Requirements**: Windows, Python 3.10+
     - **Known Limitations**: Windows-only, requires COM registration
     - **Contributing**: Link to `docs/CONTRIBUTING.md`
     - **License**: MIT

2. **Enhance `examples/README.md`**
   - Explain each example JSON file:
     - `simple-mock.json`: Minimal example for testing
     - `visio-application.json`: Real extraction from Visio (if available)
     - `excel-application.json`: Real extraction from Excel (if available)
   - Show how JSON is consumed:
     ```
     // VS Code extension example
     import * as visioApi from './extractions/visio-application.json';
     
     // Provide completion for Visio API
     vscode.languages.registerCompletionItemProvider('vba', {
       provideCompletionItems() {
         return visioApi.classes['Application'].members.properties;
       }
     });
     ```

3. **Finalize `docs/ARCHITECTURE.md`**
   - Complete sections:
     - System overview diagram (ASCII art)
     - Module responsibilities with detailed descriptions
     - Data flow: CLI → Registry → Extractor → Serializer → JSON file
     - Technology choices: Why pywin32, why click, why jsonschema
     - Extension points: How to add new type mappers, new output formats

4. **Finalize `docs/ERROR_HANDLING.md`**
   - Complete HRESULT table (see section 6.1)
   - Add troubleshooting guide:
     - "ProgID not found" → Check registry, check spelling
     - "Permission denied" → Run as administrator (rare)
     - "Type library DLL missing" → Reinstall application
   - Add debugging procedures:
     - Enable `--debug` flag
     - Check `metadata.warnings` in output JSON
     - Use `win32com` utilities to inspect COM registration

5. **Finalize `docs/CONTRIBUTING.md`**
   - Code style: black, mypy, docstrings required
   - Testing requirements: 80%+ coverage for new code
   - PR process:
     1. Create feature branch
     2. Implement with tests
     3. Update documentation
     4. Run full test suite
     5. Submit PR with description
   - Agent execution guidelines:
     - Follow milestone order
     - Update CHECKLIST.md after each task
     - Commit frequently with descriptive messages
     - Run tests before committing

6. **Update `PROJECT_PLAN.md`**
   - Add changelog section (section 10)
   - Mark all completed milestones

**Files to create/modify:**

- `README.md`
- `examples/README.md`
- `docs/ARCHITECTURE.md`
- `docs/ERROR_HANDLING.md`
- `docs/CONTRIBUTING.md`
- `PROJECT_PLAN.md` (add changelog)

**Definition of Done:**

- [x] A new contributor can read README, install, and extract a type library
- [x] All example JSON files validate against schema
- [x] Documentation is comprehensive and up-to-date
- [x] No broken links in documentation
- [x] Code examples in docs are tested and work

**Documentation updates required:**

- Update `docs/CHECKLIST.md`: Mark Milestone 6 complete
- Final review of all documentation for consistency

---

### Milestone 7: First Real Extractions

**Goal:** Produce and ship a first set of real JSON extractions for common Office apps, to validate the tool in real-world usage.

**Target ProgIDs (if installed):**

- `Scripting.FileSystemObject` (always available on Windows)
- `Visio.Application` (if Visio installed)
- `Excel.Application` (if Excel installed)
- `Word.Application` (if Word installed)
- `PowerPoint.Application` (if PowerPoint installed)
- `Outlook.Application` (if Outlook installed)

**Tasks:**

1. **Run extractions**
   - For each available ProgID:
     ```
     comsense extract "<ProgID>" --output ./examples --debug
     ```
   - Capture console output and logs

2. **Validate outputs**
   - For each generated file:
     ```
     comsense validate ./examples/<prog-id>.json
     ```
   - Check `metadata.warnings` for issues

3. **Manual review**
   - Open each JSON file
   - Check structure is reasonable:
     - Classes have meaningful names
     - Methods have parameters
     - Enums have values
   - Verify no obvious noise or incomplete data

4. **Adjust extractor if needed**
   - If issues found (e.g. enums not extracted, methods missing):
     - Fix extractor code
     - Update tests
     - Re-run extraction
   - Document any known limitations in `docs/ERROR_HANDLING.md`

5. **Commit generated files**
   - Add to `examples/` directory
   - Update `examples/README.md` to list all real extractions
   - Commit with message: `feat(examples): add real extractions for <ProgIDs>`

**Files to create/modify:**

- `examples/scripting-filesystemobject.json` (guaranteed)
- `examples/visio-application.json` (if available)
- `examples/excel-application.json` (if available)
- `examples/word-application.json` (if available)
- `examples/powerpoint-application.json` (if available)
- `examples/outlook-application.json` (if available)
- `examples/README.md`
- `comsense/extractor.py` (if fixes needed)
- `docs/ERROR_HANDLING.md` (if limitations discovered)

**Definition of Done:**

- [x] At least one realistic JSON extraction committed (`Scripting.FileSystemObject` minimum)
- [x] All extractions validate against schema
- [x] Manual review confirms plausible, useful data
- [x] Known limitations documented
- [x] Extractor refined based on real-world testing

**Documentation updates required:**

- Update `docs/CHECKLIST.md`: Mark Milestone 7 complete
- Update `README.md`: Add note about available example extractions
- Update `examples/README.md`: List all committed extractions with descriptions

---

## 10. Changelog

This section tracks major changes to the project plan and implementation status.

**Format:**

```
[YYYY-MM-DD] Milestone N completed
- Task X finished
- Task Y finished
- Documentation updated: <files>
```

**Log:**

```
[2025-11-18] PROJECT_PLAN.md created
- Initial version 1.0.0
- All 7 milestones defined
- Documentation strategy added

[Placeholder for future entries]
```

---

## 11. Future Work (Beyond Initial Implementation)

These are **not** required for the initial implementation, but should guide design decisions:

### 11.1 Performance Optimization

- **Caching**: Cache `win32com` wrapper modules to avoid regeneration
- **Incremental extraction**: Update only changed classes
- **Parallel extraction**: Extract multiple ProgIDs concurrently

### 11.2 Enhanced Type Information

- **Parameter attributes**: Map `in`, `out`, `in/out`, `retval` correctly
- **Array types**: Handle `VT_ARRAY` with element types
- **Nested types**: Support complex type hierarchies
- **Custom interfaces**: Extract IUnknown-derived interfaces

### 11.3 Output Formats

- **TypeScript declarations**: Generate `.d.ts` files directly
- **Markdown documentation**: Human-readable API docs
- **OpenAPI/Swagger**: REST API schemas for COM wrappers

### 11.4 VS Code Integration (separate repo)

- Extension: `comsense-vscode`
- Features:
  - Load JSON from `comsense` extractions
  - Provide IntelliSense for VBA/VBScript
  - Hover documentation
  - Signature help
  - Go to definition (within JSON)

### 11.5 Community Data Repository (separate repo)

- Repository: `comsense-data`
- Curated JSON extractions for popular libraries
- Automated validation in CI
- PR-based contributions with review process
- Versioned releases matching application versions

### 11.6 Cross-Platform Support

- **Linux/macOS**: Read extractions but not create them
- **Wine compatibility**: Explore COM extraction under Wine

---

## 12. Coding Agent Guidelines

When implementing this plan, a coding agent should:

### 12.1 Execution Principles

1. **Follow milestones in order**, unless explicitly instructed otherwise
2. **Complete all tasks** in a milestone before proceeding to the next
3. **Update documentation** after every milestone completion
4. **Run tests** before and after each task
5. **Commit frequently** with descriptive messages following format in section 8.2

### 12.2 Code Quality Standards

1. **Type hints**: All functions and methods must have complete type annotations
2. **Docstrings**: All public APIs must have Google-style docstrings
3. **Testing**: Minimum 80% code coverage, prefer 90%+
4. **Formatting**: Run `black comsense/ tests/` before every commit
5. **Type checking**: Run `mypy comsense/` before every commit
6. **No warnings**: Code must pass all linters without warnings

### 12.3 Documentation Discipline

1. **Update CHECKLIST.md** after completing each task
2. **Update ARCHITECTURE.md** when structure changes
3. **Update ERROR_HANDLING.md** when new errors discovered
4. **Update README.md** when user-facing features change
5. **Add changelog entry** to PROJECT_PLAN.md after each milestone

### 12.4 Decision-Making Guidelines

**If ambiguity arises, prefer:**

- Clear, simple data structures over complex abstractions
- Explicit error handling over silent failures
- Minimal but documented APIs over feature-rich undocumented ones
- Extensibility over premature optimization
- Standards compliance (JSON Schema, Python PEPs) over custom solutions

**When facing technical choices:**

1. Check if this document provides guidance → follow it
2. Check Python PEPs and best practices → align with them
3. Check existing codebase patterns → maintain consistency
4. Document the decision in commit message and update docs

**When encountering blockers:**

1. Check `docs/ERROR_HANDLING.md` for known issues
2. Add detailed information to `metadata.warnings` in output
3. Log at appropriate level (ERROR for failures, WARNING for workarounds)
4. Document limitation in `docs/ERROR_HANDLING.md` under "Known Limitations"

### 12.5 Testing Requirements

**Every new feature must have:**

- Unit tests with mocks (`@pytest.mark.unit`)
- Integration test if applicable (`@pytest.mark.windows_only`)
- Tests for error conditions and edge cases
- Clear test names describing scenario and expectation

**Before marking a milestone complete:**

```
# Run full test suite
pytest

# Check coverage
pytest --cov=comsense --cov-report=term-missing

# Verify type hints
mypy comsense/

# Format code
black comsense/ tests/

# Verify no issues
echo "All checks passed!"
```

---

## 13. Success Criteria

The project is considered successfully implemented when:

- [x] All 7 milestones completed
- [x] All tests passing (unit + integration where applicable)
- [x] Code coverage ≥80%
- [x] Documentation complete and synchronized with code
- [x] At least one real extraction committed and validated
- [x] CLI fully functional with all commands working
- [x] JSON schema stable and versioned
- [x] README provides clear getting started guide
- [x] Repository ready for public release

**Next steps after completion:**

1. Tag release: `git tag v0.1.0`
2. Publish to PyPI: `python -m build && twine upload dist/*`
3. Create GitHub release with changelog
4. Share with community for feedback
5. Begin work on Future Work items (section 11)

---

**End of PROJECT_PLAN.md**
```

Sources
[1] Claude Code: Best practices for agentic coding https://www.anthropic.com/engineering/claude-code-best-practices
[2] A practical guide to building agents https://cdn.openai.com/business-guides-and-resources/a-practical-guide-to-building-agents.pdf
[3] Optimizing Your Codebase for AI Coding Agents https://dev.to/aarongustafson/optimizing-your-codebase-for-ai-coding-agents-4ndm
[4] Coding Guidelines for Your AI Agents | The IntelliJ IDEA Blog https://blog.jetbrains.com/idea/2025/05/coding-guidelines-for-your-ai-agents/
[5] Agentic Coding Best Practices - Ben Houston's Blog https://benhouston3d.com/blog/agentic-coding-best-practices
[6] Error loading type library/DLL. (Exception from HRESULT ... https://stackoverflow.com/questions/50210215/vb-error-loading-type-library-dll-exception-from-hresult-0x80029c4a-type-e
[7] Working with custom markers https://docs.pytest.org/en/stable/example/markers.html
[8] Best practices for using AI coding Agents https://www.augmentcode.com/blog/best-practices-for-using-ai-coding-agents
[9] Error Codes in COM - Win32 apps https://learn.microsoft.com/en-us/windows/win32/learnwin32/error-codes-in-com
[10] How to Skip Tests in Pytest: Markers, Conditions, and ... https://www.browserstack.com/guide/pytest-skip
