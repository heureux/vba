Option Explicit

getIE2 'サンプル1.3.2_既に開かれているInternetExplorerを取得する(エラー処理済み)
Sub getIE2()
    Dim ie
    Dim sh
    Dim win   'ShellWindowを格納する変数
    Dim document_title  'ドキュメントタイトルの一時格納変数
    Dim objU
    '起動中のShellWindow一式を変数winsに格納
    Set sh = CreateObject("Shell.Application")
    'ShellWindowから1つずつ取得して処理
    For Each win In sh.windows
        'ドキュメントタイトル取得失敗を無視(処理継続)
        On Error Resume Next
	win.Visible = True
        document_title = ""
        document_title = win.document.Title
        On Error GoTo 0
        'タイトルバーにGoogleが含まれるかチェック
	'MsgBox document_title
        If InStr(document_title, "Google") > 0 Then
		Set objU = win.Document.getElementsByName("q")(0)
	    	If Not objU Is Nothing Then
	   		objU.Value = "a"
	   		Set objU = Nothing
		Else
			MsgBox "IEを起動しなおししてください"
			WScript.Quit
		End If

            '変数ieに取得したwinを格納
            Set ie = win
            'ループを抜ける
            'Exit For
        End If
    Next
    'URLを表示する
    MsgBox ie.LocationURL
End Sub