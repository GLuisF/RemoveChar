Entrada = InputBox("Cole o código aqui.")

For i = 1 to Len(Entrada)
	If InStr("0123456789",Mid(Entrada,i,1)) then
		Saida = Saida & Mid(Entrada,i,1)
	End If
Next

Resp = InputBox ("Clique em OK para copiar o " & _ 
		 "código corrigido para a "& _
		 "Área de transferência",,Saida)
If Resp Then
	Set WshShell = WScript.CreateObject("WScript.Shell")
	WshShell.Run "cmd.exe /c echo " & Saida & " | clip", 0, TRUE
	Set WshShell = Nothing
End If