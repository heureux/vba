Option Explicit
On Error Resume Next

Dim objFSO          ' FileSystemObject
Dim strTempFolder   ' �ꎞ�t�H���_��
Dim strTempFile     ' �ꎞ�t�@�C����

Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
If Err.Number = 0 Then
    strTempFolder = objFSO.GetSpecialFolder(2)
    strTempFile = objFSO.GetTempName()
    WScript.Echo "�ꎞ�t�@�C����: " & strTempFolder & "\" & strTempFile
Else
    WScript.Echo "�G���[: " & Err.Description
End If

Set objFSO = Nothing
