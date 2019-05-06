@ECHO OFF

SET SERVIDOR_TS=%1

setlocal ENABLEDELAYEDEXPANSION

:: verifica se foi digitado parametro
if "!SERVIDOR_TS!" equ "" (
	echo Informe o TS.
	SET /P SERVIDOR_TS=
	
	if "!SERVIDOR_TS!" equ "" GOTO :EOF
)

copy /b \\srvv09\sys\BIN\Acesso_Remoto_Clientes\Conectados\AcessoRemoto.log
echo.

cscript //nologo LeAcessoRemoto.vbs !SERVIDOR_TS! AcessoRemoto.log

endlocal

del %cd%\AcessoRemoto.log
