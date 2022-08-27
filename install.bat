:begin
@echo off

@setlocal enableextensions
@cd /d "%~dp0"

@setlocal enabledelayedexpansion

set pythonurl=https://www.python.org/ftp/python/3.10.6/python-3.10.6-amd64.exe
set scripturl=https://raw.githubusercontent.com/GiacomoGarbin/FatturaElettronica/main/fe.py

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
    echo.
    echo downloading python...
    bitsadmin /transfer "DownloadPython" %pythonurl% "%cd%\python-3.10.6-amd64.exe" > NUL 2> NUL
    if !errorlevel! NEQ 0 (
        echo  python download FAILED
        goto end
    ) else (
        echo  python download COMPLETE
    )

    echo.
    echo installing python...
    python-3.10.6-amd64.exe /passive /quiet
    if !errorlevel! NEQ 0 (
        echo  python install FAILED
        goto end
    ) else (
        echo  python install COMPLETE
    )

    del "%cd%\python-3.10.6-amd64.exe"
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
        echo  xlsxwriter install COMPLETE
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
        echo  pywin32 install COMPLETE
    )
) else (
    echo  pywin32 installed
)

echo.
echo checking if pyOpenSSL is installed...
pip show pyOpenSSL > NUL 2> NUL
if %errorlevel% NEQ 0 (
    echo  pyOpenSSL NOT installed
    echo  installing pyOpenSSL...
    pip install pyOpenSSL > NUL 2> NUL

    pip show pyOpenSSL > NUL 2> NUL
    if !errorlevel! NEQ 0 (
        echo  pyOpenSSL install FAILED
        goto end
    ) else (
        echo  pyOpenSSL install COMPLETE
    )
) else (
    echo  pyOpenSSL installed
)

echo.
echo downloading FatturaElettronica script...
bitsadmin /transfer "DownloadFatturaElettronica" %scripturl% "%userprofile%\FatturaElettronica.py" > NUL 2> NUL
if %errorlevel% NEQ 0 (
    echo  FatturaElettronica download FAILED
    goto end
) else (
    echo  FatturaElettronica download COMPLETE
)

echo.
echo adding registers...

reg add HKCR\SystemFileAssociations\.xml\shell\FatturaElettronica /d "Fattura Elettronica: xml -> xlsx" /f > NUL 2> NUL
if %errorlevel% NEQ 0 (
    echo  1st reg add FAILED
    goto end
)

reg add HKCR\SystemFileAssociations\.xml\shell\FatturaElettronica\command /d "\"%python%\" \"%userprofile%\\FatturaElettronica.py\" \"%%1\"" /f > NUL 2> NUL
if %errorlevel% NEQ 0 (
    echo  2nd reg add FAILED
    goto end
)

reg add HKCR\SystemFileAssociations\.p7m\shell\FatturaElettronica /d "Fattura Elettronica: xml -> xlsx" /f > NUL 2> NUL
if %errorlevel% NEQ 0 (
    echo  3rd reg add FAILED
    goto end
)

reg add HKCR\SystemFileAssociations\.p7m\shell\FatturaElettronica\command /d "\"%python%\" \"%userprofile%\\FatturaElettronica.py\" \"%%1\"" /f > NUL 2> NUL
if %errorlevel% NEQ 0 (
    echo  4th reg add FAILED
    goto end
)

echo  registers ADDED

echo.
echo install COMPLETE

:end
echo.
echo press ENTER to quit
pause > NUL