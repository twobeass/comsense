# comsense

> Automated extraction of Windows COM type libraries to JSON for code completion and automation tooling.

**Status:** Phase 2 completed — VSCode extension works with mock API; extraction script production-ready; end-to-end MVP flow demoable

---

## What is comsense?

**comsense** is a developer tool that enables advanced static analysis and code completion for Windows automation by extracting the full API of registered COM libraries on your system (e.g. Visio, Excel, FileSystemObject) into a standardized, machine-readable JSON format.

- Extracts class, method, property, and enum information from COM libraries.
- JSON outputs can be used for editor IntelliSense, reference docs, or community sharing.
- CLI tool provides commands for interactive discovery, extraction, and validation (planned post-MVP).
- MVP includes automated completion for VSCode via JSON, with the ability to demo using mock or production data.

---

## Current Status

- ✅ **Phase 1 Complete**: Extraction script (`extract_com.py`) extracts a chosen ProgID to correct JSON format with properties/methods; utility script lists available COM ProgIDs; see [docs/user_todo.md](docs/user_todo.md)
- ✅ **Phase 2 Complete**: VSCode extension loads any extracted (or mock) API JSON and provides working completions for `.vba`, `.bas` etc. out of the box — ready for demo, testing, and integration
- 🕓 Extracting real COM libraries not required to test extension: simply edit/copy mock files
- 🕓 Further CLI, validation, and JSON schema planned per [PROJECT_PLAN.md](PROJECT_PLAN.md)

---

## Quick Start for Both Extraction and Extension

See [docs/user_todo.md](docs/user_todo.md) for step-by-step setup for both script-based extraction and VSCode IntelliSense extension prototyping (with or without real API data).

### Extension demo on any platform

```sh
cd vscode-extension
npm install
npm run compile
# Launch extension dev mode in VSCode (see user_todo.md)
```

### Extraction (requires Windows)

```sh
python list_com_progids.py     # See ProgIDs available for extraction
python extract_com.py "<ProgID>" "examples/your-api.json"
# Copy resulting .json to vscode-extension/data/apis/ for immediate testing in the extension
```

---

## Architecture and Roadmap

See [PROJECT_PLAN.md](PROJECT_PLAN.md) for full feature set, future plans, and upcoming CLI/validation features. The MVP flow is now demoable end-to-end:

- **Discover COM APIs**: Script lists ProgIDs (Windows)
- **Extract**: Script extracts to JSON
- **Demo/Integrate**: VSCode extension consumes any JSON and provides completions for class members, properties, and methods
- **Swap in real or mock data instantly**

You can contribute or test the end-to-end experience with or without real extraction output.

---

## License

MIT
