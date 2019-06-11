if "%~1"=="" GOTO blank
PowerShell.exe -file %1 ;exit $LASTEXITCODE
EXIT %ERRORLEVEL%

:blank
EXIT -99
