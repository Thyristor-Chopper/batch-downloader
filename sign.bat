@ECHO OFF
REM G:\BuildBinaries\BIN\idw\signtool.exe
signtool sign /f F:\�⵿���\cert\gec275.pfx /p %CERTPASSWORD% %1
