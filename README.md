# comsense

> Automated extraction of Windows COM type libraries to JSON for code completion and automation tooling.

**Status:** MVP Phase 1 complete — script-based extraction [docs/user_todo.md](https://github.com/twobeass/comsense/blob/main/docs/user_todo.md)  
**License:** MIT

---

## What is comsense?

**comsense** is a developer tool that enables advanced static analysis and code completion for Windows automation by extracting the full API of registered COM libraries on your system (e.g. Visio, Excel, FileSystemObject) into a standardized, machine-readable JSON format.

- Extracts class, method, property, and enum information from COM libraries.
- JSON outputs can be used for editor IntelliSense, reference docs, or community sharing.
- CLI tool provides commands for interactive discovery, extraction, and validation.

Planned future extensions include a VSCode extension for immediate code completion and a curated repo of shared extractions.

---

## Key Features (Final Product Vision)

- **COM Library Enumeration:** Discover all ProgIDs on your system (e.g. `Visio.Application`, `Excel.Application`, etc.).
- **Intelligent Extraction:** Output a normalized API snapshot (classes, methods, properties, enums) as JSON.
- **Schema Validation:** Ensure all extracted JSONs meet a stable, evolving schema.
- **CLI UX:** `comsense list`, `comsense info`, `comsense extract`, `comsense validate`.
- **Integration Ready:** Output designed for editor plugins (e.g., VSCode), documentation systems, or code generators.
- **Community Sharing:** Encourage sharing/exchange of extracted JSONs for broader COM/VBA developer benefit.
- **Documented, Test-Covered:** Architecture, error handling and contribution docs scaffolded for future growth.

---

## How It Works

1. **Scan:** Enumerate all registered COM type libraries.
2. **Select & Extract:** Choose a ProgID to extract (e.g., `Visio.Application`), and output its type info to schema-compliant JSON.
3. **Validate:** Use the built-in validator to certify output compatibility for downstream tools.
4. **Leverage:** Use JSON in editors, scripting helpers, or contribute findings to a community repo.

---

## Current Status and MVP Workflow

> The following workflow and code represent the initial “MVP” phase. This validates extraction on a per-script basis before building the full CLI, schema, and package.

### MVP Extraction Guide (Phase 1)

**Requirements:**

- Windows OS
- Python 3.10+
- pywin32 (`pip install pywin32`)

**Basic Usage:**

Extract a COM library’s API surface to JSON:

```sh
python extract_com.py "Scripting.FileSystemObject" "examples/scripting-fso.json"
python extract_com.py "Visio.Application" "examples/visio-api.json"  # if Visio is installed
```

**How to Continue:**

- See [docs/user_todo.md](https://github.com/twobeass/comsense/blob/main/docs/user_todo.md) for step-by-step setup, extraction, and validation.
- Output JSON is validated by manually confirming structure and checking properties/methods for each class.
- Once MVP extraction is working, move on to Phase 2 (VSCode extension) and continue development per [PROJECT_PLAN_MVP.md](https://github.com/twobeass/comsense/blob/main/PROJECT_PLAN_MVP.md).

---

## Target Architecture (Final)

```
comsense/
├── comsense/                # Python package and CLI
│   ├── registry.py          # Registry scanner
│   ├── extractor.py         # TypeLib extractor
│   ├── serializer.py        # JSON output
│   ├── validator.py         # Schema validation
│   └── cli.py               # Command-line interface
├── schemas/                 # JSON schema for outputs
├── tests/
├── examples/                # Sample extracted JSONs
├── docs/                    # All project documentation
│   └── user_todo.md
├── PROJECT_PLAN.md
├── PROJECT_PLAN_MVP.md
├── README.md
└── LICENSE
```

---

## Planned CLI (Not Yet Implemented in MVP)

- `comsense list` — List discoverable COM libraries (ProgIDs)
- `comsense info <prog_id>` — Show details for a COM library
- `comsense extract <prog_id>` — Extract to schema-compliant JSON
- `comsense validate <file>` — Validate a JSON against schema

---

## Use Cases

- **Editor IntelliSense:** VSCode extension parses output for live VBA/COM code completion.
- **API Documentation:** Generate API documentation for automation libraries.
- **Community Repo:** Pool and share extractions for hard-to-inspect libraries.

---

## Requirements

- Windows 10+ (due to registry and COM)
- Python 3.10 or newer
- pywin32 library

---

## Contributing and Documentation

- See [docs/CONTRIBUTING.md](docs/CONTRIBUTING.md) for contribution guidelines.
- For in-depth plans, see [PROJECT_PLAN.md](PROJECT_PLAN.md) and [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).
- All next steps and ongoing tasks tracked in [docs/CHECKLIST.md](docs/CHECKLIST.md).

---

## License

MIT

---

**For setup and current step-by-step, start with [docs/user_todo.md](https://github.com/twobeass/comsense/blob/main/docs/user_todo.md).**  
Check [PROJECT_PLAN.md](https://github.com/twobeass/comsense/blob/main/PROJECT_PLAN.md) for detailed project milestones.
