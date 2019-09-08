Option Explicit
On Error Resume Next

Dim objFSO          ' FileSystemObject
Dim strTempFolder   ' 一時フォルダ名
Dim strTempFile     ' 一時ファイル名

Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
If Err.Number = 0 Then
    strTempFolder = objFSO.GetSpecialFolder(2)
    strTempFile = objFSO.GetTempName()
    WScript.Echo "一時ファイル名: " & strTempFolder & "\" & strTempFile
Else
    WScript.Echo "エラー: " & Err.Description
End If

Set objFSO = Nothing
