Option Explicit

' ArrayList作成
Dim ary
Set ary = CreateObject("System.Collections.ArrayList")

' 要素追加
ary.add "AK-47"
ary.add "M4"
ary.add "G3"

' 要素数取得
Dim num
num = ary.Count

WScript.Echo "要素数: " & num

' For Each文によるループ処理
Dim item
For Each item In ary
    WScript.Echo item
Next

' 要素ソート
ary.Sort

For Each item In ary
    WScript.Echo item
Next

' 要素クリア
ary.Clear


' 破棄
Set ary = Nothing