@ECHO OFF
REM G:\BuildBinaries\BIN\idw\signtool.exe
signtool sign /f F:\잡동사니\cert\gec275.pfx /p %CERTPASSWORD% %1
