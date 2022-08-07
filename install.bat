:begin
@echo off

@setlocal enableextensions
@cd /d "%~dp0"

where python > NUL 2> NUL

if %errorlevel% NEQ 0 (
    @REM set retmsg= PYTHON NON INSTALLATO
    @REM goto end
    echo  PYTHON NON INSTALLATO
    python-3.10.2-amd64.exe /passive
)

python-3.10.2-amd64.exe /passive /quiet

set /a found=0
call :loop
goto check

:loop
for /f "delims=" %%i in ('where python') do (
    set python=%%i
    if not "%python%"==[] (
        set /a found=1
        exit /b
    )
)

:check
if %found% NEQ 1 (
    set retmsg= ERRORE: "python not found"
    goto end
)

@REM pip show xlsxwriter
pip install xlsxwriter > NUL 2> NUL

if %errorlevel% NEQ 0 (
    set retmsg= ERRORE: "pip install xlsxwriter" failed
    goto end
)

pip install pywin32 > NUL 2> NUL

if %errorlevel% NEQ 0 (
    set retmsg= ERRORE: "pip install pywin32" failed
    goto end
)

@REM pip install pypiwin32

set script="fe.py"
xcopy %script% %userprofile% /y > NUL 2> NUL

if %errorlevel% NEQ 0 (
    set retmsg= ERRORE: "copy failed"
    goto end
)

reg add HKCR\SystemFileAssociations\.xml\shell\FatturaElettronica /d "Fattura Elettronica: xml -> xlsx" /f > NUL 2> NUL

if %errorlevel% NEQ 0 (
    set retmsg= ERRORE: "reg add failed"
    goto end
)

reg add HKCR\SystemFileAssociations\.xml\shell\FatturaElettronica\command /d "%python% %userprofile%\\%script% \"%%1\"" /f > NUL 2> NUL

if %errorlevel% NEQ 0 (
    set retmsg= ERRORE: "reg add failed"
    goto end
)

set retmsg= INSTALLAZIONE COMPLETATA

:end
echo.
echo %retmsg%
echo.
echo  premi INVIO per chiudere
pause > NUL