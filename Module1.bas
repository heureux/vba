Attribute VB_Name = "Module1"
Option Explicit
Dim ary() As Variant
Const P As String = "E:\Downloads\csv_zenkoku\zenkoku.csv"
Sub sample()
Debug.Print Time & " - �X�^�[�g"
'  Dim ary() As Variant
    Debug.Print IsInitArray(ary())
  If IsInitArray(ary()) = False Then Call SetCSV


Dim i As Long, max_n As Long
Dim text As String
max_n = UBound(ary())
For i = 0 To max_n
    If ary(i, 9) Like "*�����K�S�����ݒ�*" Then text = text & ary(i, 11) & vbLf
Next i
MsgBox text
  Debug.Print Time & " - �񕝂̎�������"
End Sub
'// �z�񏉊�������֐�
'// ����    �F(IN)  �z��ϐ�
'// �߂�l  �FBoolean �������ς݁�True�A����������False
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
  Debug.Print Time & " - CSV�����J�n"

  '����
  file = P '�t�@�C���w��
  max_n = CreateObject("Scripting.FileSystemObject").OpenTextFile(P, 8).Line '�t�@�C���̍s���擾
  
  ReDim ary(max_n - 1, 21) As Variant '�擾�����s����2�����z��̍Ē�`
  'CSV�t�@�C����z���
  Open file For Input As #1 'CSV�t�@�C�����J��
  Do Until EOF(1) '�ŏI�s�܂Ń��[�v
    Line Input #1, buf '�ǂݍ��񂾃f�[�^��1�s���݂Ă���
    tmp = Split(buf, ",") '�J���}�ŕ���
    For i = 0 To UBound(tmp) '���ڐ��Ԃ񃋁[�v
      ary(n, i) = tmp(i) '�����������e��z��̍��ڂ֓����i0��ID, 1������, 2���l�j
    Next i
    n = n + 1 '�z��̎��̍s��
  Loop
  Close #1 'CSV�t�@�C�������

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
'Dim file As String: file = "C:\test.csv" 'CSV�t�@�C���w��
'  Dim Items As New Collection '�R���N�V�����𐶐�
'
'  Open file For Input As #1 'CSV�t�@�C�����J��
'  Do Until EOF(1) '�ŏI�s�܂Ń��[�v
'    Dim buf As String: Line Input #1, buf '�ǂݍ��񂾃f�[�^��1�s���݂Ă���
'    Dim tmp As Variant: tmp = Split(buf, ",") '�J���}�ŕ���
'
'    With New Class1 '�C���X�^���X�̐���
'      .Name = CStr(tmp(0)) '����
'      .Price = CInt(tmp(1)) '�l�i
'      .Number = CInt(tmp(2)) '��
'      Items.Add .Self '�R���N�V�����ɒǉ�
'    End With
'  Loop
'  Close #1 'CSV�t�@�C�������
'
'  Dim item As Class1 '���[�v�p�̕ϐ�
'  For Each item In Items '�R���N�V�����������[�v
'    Debug.Print item.Name, item.Price, item.Number, item.Sale '�v���p�e�B���擾
'  Next
End Sub
