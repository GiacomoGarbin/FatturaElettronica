:begin
@echo off

@setlocal enableextensions
@cd /d "%~dp0"

@setlocal enabledelayedexpansion

set scriptname="fe.py"

echo checking if python and pip are installed...
set /a bInstallPython=0

where python > NUL 2> NUL
if %errorlevel% NEQ 0 (
    echo  python NOT installed
    set /a bInstallPython=1
) else (
    echo  python installed
)

where pip > NUL 2> NUL
if %errorlevel% NEQ 0 (
    echo  pip NOT installed
    set /a bInstallPython=1
) else (
    echo  pip installed
)

if %bInstallPython% EQU 1 (
    echo installing python...
    python-3.10.2-amd64.exe /passive /quiet
    
    if %errorlevel% NEQ 0 (
        echo python install FAILED
        goto end
    ) else (
        echo python install COMPLETE
    )
)

@REM sanity check

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
    echo.
    echo python NOT found
    goto end
)

echo.
echo upgrading pip...
python -m pip install --upgrade pip > NUL 2> NUL

echo.
echo checking if xlsxwriter is installed...
pip show xlsxwriter > NUL 2> NUL
if %errorlevel% NEQ 0 (
    echo  xlsxwriter NOT installed
    echo  installing xlsxwriter...
    pip install xlsxwriter > NUL 2> NUL

    pip show xlsxwriter > NUL 2> NUL
    if !errorlevel! NEQ 0 (
        echo  xlsxwriter install FAILED
        goto end
    ) else (
        echo  xlsxwriter install complete
    )
) else (
    echo  xlsxwriter installed
)

echo.
echo checking if pywin32 is installed...
pip show pywin32 > NUL 2> NUL
if %errorlevel% NEQ 0 (
    echo  pywin32 NOT installed
    echo  installing pywin32...
    pip install pywin32 > NUL 2> NUL

    pip show pywin32 > NUL 2> NUL
    if !errorlevel! NEQ 0 (
        echo  pywin32 install FAILED
        goto end
    ) else (
        echo  pywin32 install complete
    )
) else (
    echo  pywin32 installed
)

echo.
echo coping %scriptname% to "%userprofile%"...
xcopy %scriptname% %userprofile% /y > NUL 2> NUL
if %errorlevel% NEQ 0 (
    echo  copy %scriptname% to "%userprofile%" FAILED
    goto end
) else (
    echo  copy complete
)

echo.
echo adding registers...

reg add HKCR\SystemFileAssociations\.xml\shell\FatturaElettronica /d "Fattura Elettronica: xml -> xlsx" /f > NUL 2> NUL
if %errorlevel% NEQ 0 (
    echo  1st reg add FAILED
    goto end
)

reg add HKCR\SystemFileAssociations\.xml\shell\FatturaElettronica\command /d "\"%python%\" \"%userprofile%\\%scriptname%\" \"%%1\"" /f > NUL 2> NUL
if %errorlevel% NEQ 0 (
    echo  2nd reg add FAILED
    goto end
)

echo  registers added

echo.
echo install COMPLETE

:end
echo.
echo press ENTER to quit
pause > NUL