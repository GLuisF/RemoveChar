Const title = "Remove Char"
Const inputMsg = "Paste the text here."
inputText = InputBox(inputMsg, title)
If Not inputText = "" Then
	Const resMsg = "Click OK to copy the adjusted text to the Clipboard"
	outputText = ""
	For i = 1 to Len(inputText) Step 1
		If InStr("0123456789", Mid(inputText, i, 1)) Then outputText = outputText & Mid(inputText,i,1)
	Next
	Resp = InputBox(resMsg, title, outputText)
	If Resp Then
		Set WshShell = WScript.CreateObject("WScript.Shell")
		WshShell.Run "cmd.exe /c echo " & outputText & " | clip", 0, TRUE
		Set WshShell = Nothing
	End If
End If
