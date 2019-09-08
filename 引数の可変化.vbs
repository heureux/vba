'引数の可変化
Function sample1(ParamArray arrs())
    t = 0
    For i = 0 To UBound(arrs)
        If IsNumeric(arrs(i)) Then
            t = t + arrs(i)
        End If
    Next i
    sample1 = t
End Function



Dim objIE As New SHDocVw.InternetExplorer
Dim obj As Object 'ＩＥオブジェクト参照用
Dim strURL As String = "http://www.hogehoge.com"

'インターネットエクスプローラーのオブジェクトを作る
obj = CreateObject("Shell.Application")
System.Diagnostics.Process.Start("C:Program FilesInternet Exploreriexplore.exe", "-noframemerging" & " " & strURL)
System.Threading.Thread.Sleep(1000)
objIE = CType(obj.Windows(obj.Windows.Count - 1), SHDocVw.InternetExplorer)
　
System.Diagnostics.Process.Start(“C:Program FilesInternet Exploreriexplore.exe”, “-noframemerging” & ” ” & strURL)
ここで、internetexploreを起動しますが、”-noframemerging”を第一引数につけることで、「新規セッション」で起動させます。




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




If Intersect(Target, Columns(1)) Is Nothing Then Cancel = True: Exit Sub
Cancel = True
If Target.Interior.ColorIndex = 6 Then Target.ClearFormats Else Target.Interior.ColorIndex = 6


Sub MakeMesssage_Click()
Dim LastRow As Long     ' 行数
Dim strMessage As String
Dim i As Long
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
strMessage = ""
For i = 1 To LastRow
        If Cells(i, 1).Interior.ColorIndex = 6 Then strMessage = strMessage & Cells(i, 2).Value
Next i
MsgBox strMessage
End Sub




Dim windows As Object = Activator.CreateInstance( _
Type.GetTypeFromCLSID(Guid.Parse("{9BA05972-F6A8-11CF-A442-00A0C90A8F39}")))








