Entrada = InputBox("Paste the text here.","Remove Char")

For i = 1 to Len(Entrada)
	If InStr("0123456789",Mid(Entrada,i,1)) then
		Saida = Saida & Mid(Entrada,i,1)
	End If
Next

Resp = InputBox ("Click OK to copy the adjusted text to the Clipboard","Remove Char",Saida)

If Resp Then
	Set WshShell = WScript.CreateObject("WScript.Shell")
	WshShell.Run "cmd.exe /c echo " & Saida & " | clip", 0, TRUE
	Set WshShell = Nothing
End If
