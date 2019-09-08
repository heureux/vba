Dim wshShell
Set wshShell = WScript.CreateObject("WScript.Shell")

wshShell.Run "iexplore -nomerge http://www.google.com"

Dim objShell
Set objShell = CreateObject("Shell.Application")

Dim objShellWindows
Set objShellWindows = objShell.Windows

Dim i
Dim ieObject
For i = 0 To objShellWindows.Count - 1
    If InStr(objShellWindows.Item(i).FullName, "iexplore.exe") <> 0 Then
        Set ieObject = objShellWindows.Item(i)
        If VarType(ieObject.Document) = 8 Then
            MsgBox "Loaded " & ieObject.Document.Title
            Exit For
        End If
    End If
Next

Set ieObject = Nothing
Set objShellWindows = Nothing
Set objShell = Nothing
Set wshShell = Nothing