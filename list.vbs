Option Explicit

' ArrayList�쐬
Dim ary
Set ary = CreateObject("System.Collections.ArrayList")

' �v�f�ǉ�
ary.add "AK-47"
ary.add "M4"
ary.add "G3"

' �v�f���擾
Dim num
num = ary.Count

WScript.Echo "�v�f��: " & num

' For Each���ɂ�郋�[�v����
Dim item
For Each item In ary
    WScript.Echo item
Next

' �v�f�\�[�g
ary.Sort

For Each item In ary
    WScript.Echo item
Next

' �v�f�N���A
ary.Clear


' �j��
Set ary = Nothing