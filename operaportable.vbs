Option Explicit
Dim sAction, sAppPath, sExecPath, sIconPath, objFile, sbaseKey, sbaseKey2, sAppDesc
Dim sClsKey, ArrKeys, regkey
Dim WshShell : Set WshShell = CreateObject("WScript.Shell")
Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = oFSO.GetFile(WScript.ScriptFullName)
sAppPath = oFSO.GetParentFolderName(objFile)
sExecPath = sAppPath & "\opera.exe"
sIconPath = sAppPath & "\opera.exe"
sAppDesc = "Opera delivers fast, secure, and customizable web browsing experience. " & _
"A familiar user interface, built-in ad blocker, and integrated messengers " & _
"let you browse the web efficiently and stay connected."

'Quit if opera.exe is missing in the current folder!
If Not oFSO.FileExists(sExecPath) Then
   MsgBox "Please run this script from Opera Portable folder. The script will now quit.", _
   vbOKOnly + vbInformation, "Register Opera Portable with Default Apps"
   WScript.Quit
End If

If InStr(sExecPath, " ") > 0 Then
   sExecPath = """" & sExecPath & """"
   sIconPath = """" & sIconPath & """"
End If

sbaseKey = "HKCU\Software\"
sbaseKey2 = sbaseKey & "Clients\StartmenuInternet\Opera Portable\"
sClsKey = sbaseKey & "Classes\"

If WScript.Arguments.Count > 0 Then
   If UCase(Trim(WScript.Arguments(0))) = "-REG" Then Call RegisterOperaPortable
   If UCase(Trim(WScript.Arguments(0))) = "-UNREG" Then Call UnRegisterOperaPortable
Else
   sAction = InputBox("Type REGISTER to add Opera Portable to Default Apps. " & _
   "Type UNREGISTER to remove.", "Opera Portable Registration", "REGISTER")
   If UCase(Trim(sAction)) = "REGISTER" Then Call RegisterOperaPortable
   If UCase(Trim(sAction)) = "UNREGISTER" Then Call UnRegisterOperaPortable
End If

Sub RegisterOperaPortable   
   WshShell.RegWrite sbaseKey & "RegisteredApplications\Opera Portable", _
   "Software\Clients\StartMenuInternet\Opera Portable\Capabilities", "REG_SZ"
   
   'OperaHTML registration
   WshShell.RegWrite sClsKey & "OperaHTML2\", "Opera HTML Document", "REG_SZ"
   WshShell.RegWrite sClsKey & "OperaHTML2\EditFlags", 2, "REG_DWORD"
   WshShell.RegWrite sClsKey & "OperaHTML2\FriendlyTypeName", "Opera HTML Document", "REG_SZ"
   WshShell.RegWrite sClsKey & "OperaHTML2\DefaultIcon\", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sClsKey & "OperaHTML2\shell\", "open", "REG_SZ"
   WshShell.RegWrite sClsKey & "OperaHTML2\shell\open\command\", sExecPath & _
   " ""%1""", "REG_SZ"
   WshShell.RegWrite sClsKey & "OperaHTML2\shell\open\ddeexec\", "", "REG_SZ"
   
   'OperaPDF registration
   WshShell.RegWrite sClsKey & "OperaPDF2\", "Opera PDF Document", "REG_SZ"
   WshShell.RegWrite sClsKey & "OperaPDF2\EditFlags", 2, "REG_DWORD"
   WshShell.RegWrite sClsKey & "OperaPDF2\FriendlyTypeName", "Opera PDF Document", "REG_SZ"
   WshShell.RegWrite sClsKey & "OperaPDF2\DefaultIcon\", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sClsKey & "OperaPDF2\shell\open\", "open", "REG_SZ"
   WshShell.RegWrite sClsKey & "OperaPDF2\shell\open\command\", sExecPath & _
   " ""%1""", "REG_SZ"
   
   'OperaURL registration
   WshShell.RegWrite sClsKey & "OperaURL2\", "Opera URL", "REG_SZ"
   WshShell.RegWrite sClsKey & "OperaURL2\EditFlags", 2, "REG_DWORD"
   WshShell.RegWrite sClsKey & "OperaURL2\FriendlyTypeName", "Opera URL", "REG_SZ"
   WshShell.RegWrite sClsKey & "OperaURL2\URL Protocol", "", "REG_SZ"
   WshShell.RegWrite sClsKey & "OperaURL2\DefaultIcon\", sIconPath & ",0", "REG_SZ"
   WshShell.RegWrite sClsKey & "OperaURL2\shell\open\", "open", "REG_SZ"
   WshShell.RegWrite sClsKey & "OperaURL2\shell\open\command\", sExecPath & _
   " ""%1""", "REG_SZ"
   WshShell.RegWrite sClsKey & "OperaURL2\shell\open\ddeexec\", "", "REG_SZ"   
   
   'Default Apps Registration/Capabilities
   WshShell.RegWrite sbaseKey2, "Opera Portable", "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "Capabilities\ApplicationDescription", sAppDesc, "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "Capabilities\ApplicationIcon", sIconPath, "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "Capabilities\ApplicationName", "Opera Portable", "REG_SZ" 
   WshShell.RegWrite sbaseKey2 & "Capabilities\FileAssociations\.pdf", "OperaPDF2", "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "Capabilities\StartMenu", "Opera Portable", "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "DefaultIcon\", sIconPath, "REG_SZ"
   WshShell.RegWrite sbaseKey2 & "shell\open\command\", sExecPath, "REG_SZ"
   
   ArrKeys = Array ( _
   "FileAssociations\.htm", _
   "FileAssociations\.html", _
   "URLAssociations\http", _
   "URLAssociations\https", _
   "URLAssociations\mailto" _
   )
   
   For Each regkey In ArrKeys
      WshShell.RegWrite sbaseKey2 & "Capabilities\" & regkey, "OperaHTML2", "REG_SZ"
   Next      
   
   'Override the default app name by which the program appears in Default Apps (*Optional*)
   WshShell.RegWrite sClsKey & "OperaHTML2\Application\ApplicationName", "Opera Portable", "REG_SZ"
   
   'Launch Default Programs or Default Apps after registering Opera Portable   
   WshShell.Run "control /name Microsoft.DefaultPrograms /page pageDefaultProgram"
End Sub


Sub UnRegisterOperaPortable
   sbaseKey = "HKCU\Software\"
   sbaseKey2 = "HKCU\Software\Clients\StartmenuInternet\Opera Portable"   
   
   On Error Resume Next
   WshShell.RegDelete sbaseKey & "RegisteredApplications\Opera Portable"
   On Error GoTo 0
   
   WshShell.Run "reg.exe delete " & sClsKey & "OperaHTML2" & " /f", 0
   WshShell.Run "reg.exe delete " & sClsKey & "OperaPDF2" & " /f", 0
   WshShell.Run "reg.exe delete " & sClsKey & "OperaURL2" & " /f", 0
   WshShell.Run "reg.exe delete " & chr(34) & sbaseKey2 & chr(34) & " /f", 0
   
   'Launch Default Apps after unregistering Opera Portable   
   WshShell.Run "control /name Microsoft.DefaultPrograms /page pageDefaultProgram"   
End Sub
