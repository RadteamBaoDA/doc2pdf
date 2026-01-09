---
description: Init project
---
1. Create the project directory structure if it doesn't exist:
   - `src/core`
   - `src/utils`
   - `tests`
   - `logs`

2. Initialize `pyproject.toml` with build system (setuptools), project metadata, and dependencies (typer, rich, pyyaml, loguru).

3. Create `src/__init__.py` to make `src` a package.

4. Create `src/__main__.py` as the entry point:
   - Should import `cli` from `src.cli` and run it.

5. Create `src/cli.py` to handle command line interface:
   - Use `Typer` app.
   - Define a callback for global settings (config).
   - Setup logging initialization.

6. Create `src/config.py` to handle configuration:
   - Load from `config.yml`.
   - Use `pydantic` or simple class for config validation.

7. Create `src/utils/logger.py` or similar for logging setup using `loguru`.

8. Create `README.md` with usage instructions.
