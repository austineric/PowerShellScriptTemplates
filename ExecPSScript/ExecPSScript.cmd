if "%~1"=="" GOTO blank
pwsh.exe -file %1
EXIT %ERRORLEVEL%

:blank
EXIT -99
