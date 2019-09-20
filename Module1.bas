Attribute VB_Name = "Module1"
Option Explicit
Dim ary() As Variant
Const P As String = "E:\Downloads\csv_zenkoku\zenkoku.csv"
Sub sample()
Debug.Print Time & " - スタート"
'  Dim ary() As Variant
    Debug.Print IsInitArray(ary())
  If IsInitArray(ary()) = False Then Call SetCSV


Dim i As Long, max_n As Long
Dim text As String
max_n = UBound(ary())
For i = 0 To max_n
    If ary(i, 9) Like "*西牟婁郡すさみ町*" Then text = text & ary(i, 11) & vbLf
Next i
MsgBox text
  Debug.Print Time & " - 列幅の自動調整"
End Sub
'// 配列初期化判定関数
'// 引数    ：(IN)  配列変数
'// 戻り値  ：Boolean 初期化済み＝True、未初期化＝False
Function IsInitArray(ary()) As Boolean
    If Sgn(ary) <> 0 Then
        IsInitArray = True
    Else
        IsInitArray = False
    End If
End Function


Sub SetCSV()
Dim col As VBA.Collection
Set col = New VBA.Collection
  Dim file As String, max_n As Long
  Dim buf As String, tmp As Variant
  Dim i As Long, n As Long, val As Long
  Debug.Print Time & " - CSV処理開始"

  '準備
  file = P 'ファイル指定
  max_n = CreateObject("Scripting.FileSystemObject").OpenTextFile(P, 8).Line 'ファイルの行数取得
  
  ReDim ary(max_n - 1, 21) As Variant '取得した行数で2次元配列の再定義
  'CSVファイルを配列へ
  Open file For Input As #1 'CSVファイルを開く
  Do Until EOF(1) '最終行までループ
    Line Input #1, buf '読み込んだデータを1行ずつみていく
    tmp = Split(buf, ",") 'カンマで分割
    For i = 0 To UBound(tmp) '項目数ぶんループ
      ary(n, i) = tmp(i) '分割した内容を配列の項目へ入れる（0→ID, 1→名称, 2→値）
    Next i
    n = n + 1 '配列の次の行へ
  Loop
  Close #1 'CSVファイルを閉じる

End Sub
Sub test1()
Dim a As Long, ret As Long
Dim Path As String, WSH As Variant
    Set WSH = CreateObject("WScript.Shell")
    Path = WSH.SpecialFolders("Desktop") & "\"
a = 2
Path = Path & "back.vbs"
ret = CreateObject("Wscript.Shell").Run(Path & " " & a, 0, True)
'Range("B3").CurrentRegion.Offset(1).Resize(Range("B3").CurrentRegion.Rows.Count - 1).Select
Debug.Print ret


'https://ateitexe.com/excel-vba-csv-to-class-collection/
'Dim file As String: file = "C:\test.csv" 'CSVファイル指定
'  Dim Items As New Collection 'コレクションを生成
'
'  Open file For Input As #1 'CSVファイルを開く
'  Do Until EOF(1) '最終行までループ
'    Dim buf As String: Line Input #1, buf '読み込んだデータを1行ずつみていく
'    Dim tmp As Variant: tmp = Split(buf, ",") 'カンマで分割
'
'    With New Class1 'インスタンスの生成
'      .Name = CStr(tmp(0)) '名称
'      .Price = CInt(tmp(1)) '値段
'      .Number = CInt(tmp(2)) '個数
'      Items.Add .Self 'コレクションに追加
'    End With
'  Loop
'  Close #1 'CSVファイルを閉じる
'
'  Dim item As Class1 'ループ用の変数
'  For Each item In Items 'コレクション内をループ
'    Debug.Print item.Name, item.Price, item.Number, item.Sale 'プロパティを取得
'  Next
End Sub
