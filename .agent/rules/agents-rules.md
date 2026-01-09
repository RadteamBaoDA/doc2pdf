---
trigger: always_on
---

# Antigravity Rules: doc2pdf

## Project Overview
`doc2pdf` is a modern Python CLI application designed for document conversion. The project follows a strict `src` layout with feature-driven modularization and a focus on clean, maintainable CLI interfaces.

---

## 1. Architectural Principles
- **Feature-Driven Development:** Logic should be encapsulated within feature modules under `src/core/` or `src/utils/`. Avoid monolithic files. 
- **Source Layout:** All application code resides in the `src/` directory.
- **Dependency Management:** Use `pyproject.toml` for project metadata and dependencies. `requirements.txt` should be kept in sync for legacy compatibility.
- **Design pattern:** Using Layered Architecture with the Strategy Pattern for conversion

---

## 2. Python Modernization Rules
- **Type Hinting:** All new functions and classes **must** include PEP 484 type hints.
- **Async/Await:** Use asynchronous patterns for I/O bound tasks (file conversion, network requests) where applicable.
- **Pathlib:** Use `pathlib.Path` for all filesystem operations instead of `os.path`.
- **Packaging:** Adhere to modern standards defined in `pyproject.toml`. Prefer `importlib.metadata` for versioning.

---

## 3. CLI Development Rules
- **Entry Points:** The primary entry point is `src/__main__.py`. This should invoke the CLI handler defined in `src/cli.py`.
- **Framework:** Default to using `Typer` or `Click` for CLI implementation to ensure consistent help menus and argument parsing.
- **Configuration:** Use `src/config.py` to load settings from `config.yml`. Environment variables should override YAML settings.
- **Logging:** Direct all application logs to the `logs/` directory using the standard `logging` library or `loguru`.
- **Console UI:** Console UI Pattern using the Rich library.
---

## 4. File Structure Rules
- **src/core:** Contains the "heart" of the application (e.g., conversion engines, business logic).
- **src/utils:** Contains shared utility functions (e.g., file validation, formatting).
- **tests:** All tests must mirror the `src` structure and use `pytest`.
- **Init Files:** Every directory in `src/` must contain an `__init__.py` to maintain package integrity.
- **Output folder:** Keep folder structure and folder/file name base on structure of input
---

## 5. Development Workflow
- **Linting/Formatting:** Code must be formatted with `black` and linted with `ruff` before submission.
- **Testing:** Ensure `pytest` passes with 80%+ coverage. Check `.pytest_cache` is ignored by Git.
- **Git:** Never commit `__pycache__`, `.venv`, or `logs`. These are strictly local.

---

## 6. Antigravity Agent Instructions
- When adding a new conversion feature, create a new module in `src/core/`.
- When modifying CLI arguments, update `src/cli.py` and the `README.md` documentation simultaneously.
- If a new dependency is required, add it to `pyproject.toml` first.