@echo off
cd /d "%~dp0"
where pyw >nul 2>nul && (start "" pyw fatura_doldur.py) || (start "" pythonw fatura_doldur.py)
