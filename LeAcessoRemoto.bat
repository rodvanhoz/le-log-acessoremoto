@ECHO OFF

copy /b \\srvv09\sys\BIN\Acesso_Remoto_Clientes\Conectados\AcessoRemoto.log
echo.

cscript //nologo LeAcessoRemoto.vbs %1 AcessoRemoto.log %2

del %cd%\AcessoRemoto.log

