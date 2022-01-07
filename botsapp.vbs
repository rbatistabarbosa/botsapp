dim mensagem
dim numero

set arquivoMensagem = CreateObject("Scripting.FileSystemObject").OPenTextFile("mensagem.txt", 1)
mensagem = arquivoMensagem.ReadLine()
arquivoMensagem.close

set shell = WScript.CreateObject("WScript.Shell")

set arquivoNumeros = CreateObject("Scripting.FileSystemObject").OPenTextFile("numeros.txt", 1)
do while not arquivoNumeros.AtEndOfStream
	numero = arquivoNumeros.ReadLine()
	shell.run("https://web.whatsapp.com/send?phone="&numero&"&text="&mensagem)
	WScript.sleep(8000)
	shell.SendKeys "{TAB}"
	WScript.sleep(100)
	shell.SendKeys "{TAB}"
	WScript.sleep(100)
	shell.SendKeys "{TAB}"
	WScript.sleep(100)
	shell.SendKeys "{TAB}"
	WScript.sleep(100)
	shell.SendKeys "{TAB}"
	WScript.sleep(100)
	shell.SendKeys "{TAB}"
	WScript.sleep(100)
	shell.SendKeys "{TAB}"
	WScript.sleep(100)
	shell.SendKeys "{TAB}"
	WScript.sleep(100)
	shell.SendKeys "{TAB}"
	WScript.sleep(100)
	shell.SendKeys "{TAB}"
	WScript.sleep(100)
	shell.SendKeys "{ENTER}"
	WScript.sleep(2000)
	shell.SendKeys "^w"
Loop
arquivoNumeros.close

MsgBox "Script concluído!"