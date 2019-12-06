Option Explicit
If AppPrevInstance() Then 
    MsgBox "There is an existing proceeding !" & VbCrLF & CommandLineLike(WScript.ScriptName),VbExclamation,"There is an existing proceeding !"    
    WScript.Quit   
Else
    Dim ws,fso,Srcimage,Temp,PathOutPutHTML,fhta
    Set ws = CreateObject("wscript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Temp = WS.ExpandEnvironmentStrings("%Temp%")
    PathOutPutHTML = Temp & "\image.hta"
    Set fhta = fso.OpenTextFile(PathOutPutHTML,2,True)
    Srcimage = "https://scontent-waw1-1.xx.fbcdn.net/v/t1.15752-9/69103267_1159992570862409_7726830110363353088_n.jpg?_nc_cat=108&_nc_ohc=_Q15XGpTedwAQlqzxs2L3wr7X0FdVzzHl2GAcSmrr5v5WfHVKNGr4JsHg&_nc_ht=scontent-waw1-1.xx&oh=35a1e6a8e6cc5082260d9a1744cc530e&oe=5E74535B"
    Do
        Call LoadImage(Srcimage)
        ws.run "mshta.exe " & PathOutPutHTML
        Call Pause(20)
    Loop
End If
'********************************************************************************************************
Sub LoadImage(Srcimage)
    fhta.WriteLine "<html>"
    fhta.WriteLine "    <hta:application id=""oHTA"" "
    fhta.WriteLine "        border=""none"" "
    fhta.WriteLine "        caption=""no"" "
    fhta.WriteLine "        contextmenu=""no"" "
    fhta.WriteLine "        innerborder=""no"" "
    fhta.WriteLine "        scroll=""no"" "
    fhta.WriteLine "        showintaskbar=""no"" "
    fhta.WriteLine "    />"
    fhta.WriteLine "    <script language=""VBScript"">"
    fhta.WriteLine "        Sub Window_OnLoad"
    fhta.WriteLine "            'Resize and position the window"
    fhta.WriteLine "            width = 460 : height = 510"
    fhta.WriteLine "            window.resizeTo width, height"
    fhta.WriteLine "            window.moveTo screen.availWidth\2 - width\2, screen.availHeight\2 - height\2"
    fhta.WriteLine "            'Automatically close the windows after 5 seconds"
    fhta.WriteLine "            idTimer = window.setTimeout(""vbscript:window.close"",10000)"
    fhta.WriteLine "        End Sub"
    fhta.WriteLine "    </script>"
    fhta.WriteLine "<body>"
    fhta.WriteLine "    <table border=0 width=""100%"" height=""100%"">"
    fhta.WriteLine "        <tr>"
    fhta.WriteLine "            <td align=""center"" valign=""middle"">"
    fhta.WriteLine "                <img src= "& Srcimage & ">"
    fhta.WriteLine "            </td>"
    fhta.WriteLine "        </tr>"
    fhta.WriteLine "    </table>"
    fhta.WriteLine "</body>"
    fhta.WriteLine "</html>"
End Sub
'**********************************************************************************************s
Sub Pause(Min)
    Wscript.Sleep(5000)
End Sub  
'**********************************************************************************************
Function AppPrevInstance()   
    With GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")   
        With .ExecQuery("SELECT * FROM Win32_Process WHERE CommandLine LIKE " & CommandLineLike(WScript.ScriptFullName) & _
            " AND CommandLine LIKE '%WScript%' OR CommandLine LIKE '%cscript%'")   
            AppPrevInstance = (.Count > 1)   
        End With   
    End With   
End Function   
'**************************************************************************
Function CommandLineLike(ProcessPath)   
    ProcessPath = Replace(ProcessPath, "\", "\\")   
    CommandLineLike = "'%" & ProcessPath & "%'"   
End Function