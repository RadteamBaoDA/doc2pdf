from typer.testing import CliRunner
from src.cli import app
from src import __version__

runner = CliRunner()

def test_version():
    result = runner.invoke(app, ["--version"])
    assert result.exit_code == 0
    assert f"version: {__version__}" in result.stdout

def test_convert_missing_file():
    result = runner.invoke(app, ["convert", "non_existent_file.txt"])
    assert result.exit_code == 1
    assert "Input file not found" in result.stdout
