$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$pythonCommand = Get-Command py -ErrorAction SilentlyContinue
if ($null -eq $pythonCommand) {
    $pythonCommand = Get-Command python -ErrorAction Stop
}

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    & $pythonCommand.Source -m venv .venv
}

& ".\.venv\Scripts\python.exe" -m pip install -r requirements.txt
& ".\.venv\Scripts\python.exe" -m streamlit run app.py
