Option Explicit

getIE2 '�T���v��1.3.2_���ɊJ����Ă���InternetExplorer���擾����(�G���[�����ς�)
Sub getIE2()
    Dim ie
    Dim sh
    Dim win   'ShellWindow���i�[����ϐ�
    Dim document_title  '�h�L�������g�^�C�g���̈ꎞ�i�[�ϐ�
    Dim objU
    '�N������ShellWindow�ꎮ��ϐ�wins�Ɋi�[
    Set sh = CreateObject("Shell.Application")
    'ShellWindow����1���擾���ď���
    For Each win In sh.windows
        '�h�L�������g�^�C�g���擾���s�𖳎�(�����p��)
        On Error Resume Next
	win.Visible = True
        document_title = ""
        document_title = win.document.Title
        On Error GoTo 0
        '�^�C�g���o�[��Google���܂܂�邩�`�F�b�N
	'MsgBox document_title
        If InStr(document_title, "Google") > 0 Then
		Set objU = win.Document.getElementsByName("q")(0)
	    	If Not objU Is Nothing Then
	   		objU.Value = "a"
	   		Set objU = Nothing
		Else
			MsgBox "IE���N�����Ȃ������Ă�������"
			WScript.Quit
		End If

            '�ϐ�ie�Ɏ擾����win���i�[
            Set ie = win
            '���[�v�𔲂���
            'Exit For
        End If
    Next
    'URL��\������
    MsgBox ie.LocationURL
End Sub