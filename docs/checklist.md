# comsense Project Task Checklist

> Tracks all core MVP and final product tasks, aligned to PROJECT_PLAN_MVP.md and PROJECT_PLAN.md. Each task is marked as "User" (requires human) or "Agent" (AI/coding agent automated), including current progress status.

---

## Legend
- **[x]** Complete
- **[ ]** In progress / To do
- **User:** Manual/validation step (requires human)
- **Agent:** Automated/coding agent task

---

## Phase 1: Extraction Script (MVP)

- **[x]** **Agent**: Create extraction script skeleton (`extract_com.py`)
- **[x]** **Agent**: Implement COM library loading and type/class discovery
- **[x]** **Agent**: Extract properties and methods for each class (basic structure)
- **[x]** **User**: Test extraction script on Windows (Python 3.10+, pywin32)
- **[x]** **Agent**: Commit extracted JSON (mock and/or real) to `examples/`
- **[x]** **User**: Validate JSON structure, check for properties/methods, basic data quality
- **[x]** **Agent**: Add utility to list all available ProgIDs (`list_com_progids.py`)

## Phase 2: VSCode Extension (MVP)

- **[x]** **Agent**: Scaffold extension structure and setup (package.json, tsconfig, etc.)
- **[x]** **Agent**: Define API data types (`types.ts`)
- **[x]** **Agent**: Implement API loader (`apiLoader.ts`)
- **[x]** **Agent**: Implement completion provider (class/method/property)
- **[x]** **Agent**: Implement extension entrypoint
- **[x]** **User**: Test extension with mock-api.json for completions in VSCode
- **[x]** **User**: Validate completion behavior, test swapping in real extracted JSON

## Phase 3: Validation & Community Feedback (MVP Finalization)

- **[ ]** **User**: Extract and test with ≥1 real COM library (Visio, FSO, Office, etc.)
- **[ ]** **User**: Validate end-to-end workflow on Windows and in VSCode
- **[ ]** **User**: Document any issues/lessons in VALIDATION.md
- **[ ]** **User**: Decide to proceed to production plan or review MVP

---

## Final Product Roadmap (Highlights from PROJECT_PLAN.md)

### Registry & Extraction Tools
- **[ ]** **Agent**: Full registry scanner with filtering (`RegistryScanner`)
- **[ ]** **Agent**: Full-featured extraction core with enum/type/param mapping
- **[ ]** **Agent**: JSON Schema validation, error/warning propagation in output
- **[ ]** **Agent**: CLI tool: `comsense list`, `comsense info`, `comsense extract`, `comsense validate`

### Documentation, Schemas, and QA
- **[x]** **Agent**: User onboarding docs (`user_todo.md`, README)
- **[ ]** **Agent/User**: Finalize and keep up-to-date: `CHECKLIST.md`, `ARCHITECTURE.md`, `CONTRIBUTING.md`, `ERROR_HANDLING.md`, `VALIDATION.md`
- **[ ]** **Agent**: JSON schema placed in `schemas/`, used for validation as well as docs/test
- **[ ]** **User**: Validation against schema for all production JSON extractions

### Extension and Integration
- **[ ]** **Agent**: VSCode extension production features (hover, signature, multi-API, settings)
- **[ ]** **User**: Community test: contribute, submit or consume extraction JSONs
- **[ ]** **Agent/User**: Publish extension + extraction tool, maintain releases

---

## Status Summary

- **MVP extraction and Phase 2 extension demo workflow complete** (mock- and real-data ready)
- **Awaiting user testing with real extractions** for full MVP validation
- **Production CLI, schema validation, and advanced extension features to be built post-validation**

---
