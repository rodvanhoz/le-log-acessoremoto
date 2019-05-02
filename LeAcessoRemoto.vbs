' Atulaliza o Obj da base teste e gera o arquivo
'
' retorno:
'	0 -> Sem erros
'	1 -> Erro de processamento
'
' Data: 21/06/2018
' Author: Rodrigo Vanhoz Ribeiro
'
' Versão: 1
'
' Alteracoes informar abaixo
'

Option Explicit

' checagem de parametros
If WScript.Arguments.Count <> 1 AND WScript.Arguments.Count <> 5 Then
	WScript.Echo "uso: LeAcessoRemoto [Conexao]"
	WScript.Echo ""
	WScript.Quit( 1 )
End if

' Constantes para uso de manipulacao de arquivos
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
 
' parametros recebidos
dim nomearqremoto
nomearqremoto = wscript.arguments.parameters(0)

' variaveis de arquivo
dim fs, arqremoto, linhas, linha
Set fs        = CreateObject( "scripting.filesystemobject" )
Set arqremoto = fs.OpenTextFile( nomearqremoto, ForReading, TristateFalse )

' carregando arquivo
arqremoto.realall
arqremoto.close


