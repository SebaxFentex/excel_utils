@echo off
cd /d "%~dp0..\.."
uv run python tests/test_consulta_power_query.py
set EXIT_CODE=%ERRORLEVEL%
echo.
echo Codigo de salida: %EXIT_CODE%
pause
exit /b %EXIT_CODE%
