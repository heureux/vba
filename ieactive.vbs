Const strIEexe = "iexplore.exe" 'IEのプロセス名
Dim intProcID

'IEのウィンドウをアクティブ化する
call ActiveIE

sub ActiveIE()
    Dim objWshShell

    GetProcID(strIEexe)
    Set objWshShell = CreateObject("Wscript.Shell")
    objWshShell.AppActivate intProcID
    Set objWshShell = Nothing
End Sub

Function GetProcID(ProcessName)
    Dim Service
    Dim QfeSet
    Dim Qfe

    Set Service = WScript.CreateObject("WbemScripting.SWbemLocator").ConnectServer
    Set QfeSet = Service.ExecQuery("Select * From Win32_Process Where Caption='"& ProcessName &"'")

    intProcID = 0
	'msgbox QfeSet.count
    For Each Qfe in QfeSet
	msgbox Qfe.Name
        intProcID = Qfe.ProcessId
        'Exit For
    Next

    GetProcID = intProcID <> 0
End Function