@ECHO OFF
upx.exe -9 %1
sign.bat %1
