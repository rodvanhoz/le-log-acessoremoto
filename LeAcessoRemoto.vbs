' Le o arquivo de log dos acesso remotos e retorna os acessos de um ts passado por parametro do dia
'
' retorno:
'	0 -> Sem erros
'	1 -> Erro de processamento
'
' Data: 02/05/
' Author: Rodrigo Vanhoz Ribeiro
'
' Vers√£o: 1
'
' Alteracoes informar abaixo
'

Option Explicit

' checagem de parametros
If WScript.Arguments.Count <> 2 And WScript.Arguments.Count <> 3 Then
	WScript.Echo "uso: LeAcessoRemoto [Conexao] [Caminho LOG] [Data dd/MM/yyyy (deixar em branco traz a data do dia)]"
	WScript.Echo ""
	WScript.Quit( 1 )
End if

' Constantes para uso de manipulacao de arquivos
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
 
' parametros recebidos
dim nomearqremoto, data, nomeserver
nomeserver = acertaNomeServer(WScript.Arguments.Unnamed(0))
nomearqremoto = WScript.Arguments.Unnamed(1)


If WScript.Arguments.Count = 3 Then
	data = WScript.Arguments.Unnamed(2)
Else
	data = Date 
End If

' variaveis de arquivo
dim fs, arqremoto, linhas, linha, totalLinhas, cont, continf, contAcesso, achou, x
Set fs        = CreateObject( "scripting.filesystemobject" )
Set arqremoto = fs.OpenTextFile( nomearqremoto, ForReading, TristateFalse )

achou = False
cont = 0
continf = 0
contAcesso = 0

' variaveis do log
Dim login, servidor, usr, dtAcesso, hrAcesso, hlogin, hservidor, husr, hdtAcesso, hhrAcesso
Set hlogin    = CreateObject("scripting.dictionary")
Set hservidor = CreateObject("scripting.dictionary")
Set husr      = CreateObject("scripting.dictionary")
Set hdtAcesso = CreateObject("scripting.dictionary")
Set hhrAcesso = CreateObject("scripting.dictionary")


' carregando arquivo
linhas = Split(arqremoto.ReadAll, Chr(13) & Chr(10))
totalLinhas = arqremoto.Line
arqremoto.Close

For cont = 0 to (UBound(linhas) - 1)
	linha = Trim(linhas(cont))
	If Left(linha, 2) <> "--" And Left(linha, 2) <> "" Then
		continf = continf + 1
		If Left(linha, 5) = "login" Then
		    If login <> "" And achou = True Then
		    	hlogin.Add contAcesso, login
		    	hservidor.Add contAcesso, servidor
		    	husr.Add contAcesso, usr
		    	hdtAcesso.Add contAcesso, dtAcesso
		    	hhrAcesso.Add contAcesso, hrAcesso
		    	achou = False
		    End If
		    
		    contAcesso = contAcesso + 1
			login = linha
			
		Elseif continf = 2 Then
			servidor = linha
			
		Elseif continf = 3 Then
			usr = linha
			
		Elseif Mid(linha, 3, 1) = "/" Then
			dtAcesso = linha
			If DateDiff("y", data, dtAcesso) = 0 Then
				achou = True
			End If
		
		Elseif Mid(linha, 3, 1) = ":" Then
			hrAcesso = linha
			continf = 0
		End if
	End if
Next

Dim chaves
chaves = hlogin.Keys

For each x in chaves
	If UCase(hservidor.Item(x)) = UCase(nomeserver) Then
		WScript.Echo "-------------------"
		WScript.Echo hlogin.Item(x)
		WScript.Echo hservidor.Item(x)
		WScript.Echo husr.Item(x)
		WScript.Echo hdtAcesso.Item(x)
		WScript.Echo hhrAcesso.Item(x)
	End if
Next

WScript.StdOut.WriteBlankLines 2
WScript.Quit(0)


' FUN«’ES

Function acertaNomeServer(nome)
	Dim tkn, tmp
	acertaNomeServer = nome
	
	If InStr(nome, "/") > 0 Then
		tkn = Split(nome, "/")
		tmp = tkn(UBound(tkn))
		tkn = Split(tmp, ".")
		acertaNomeServer = tkn(0)
	End If
End function