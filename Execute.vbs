'COLETA DISPOSITIVOS NOT PRESENTS
On Error Resume Next
Set WshShell = CreateObject("WScript.Shell")
Const DeleteReadOnly = True
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strScriptPath = objFSO.GetParentFolderName(Wscript.ScriptFullName)
Set objFolder = objFSO.CreateFolder(strScriptPath & "\TEMP")
Set objTextALL = objFSO.CreateTextFile(strScriptPath & "\TEMP\ALL.txt", ForWriting)
Set objTextNotPresent = objFSO.CreateTextFile(strScriptPath & "\TEMP\NOTPRESENTS.txt", ForWriting)
Set objTextToRemove = objFSO.CreateTextFile(strScriptPath & "\TEMP\TOREMOVE.txt", ForWriting)

Set colOSes = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each objOS In colOSes
    strSO = objOS.Caption
    strArc = objOS.OSArchitecture
    If instr(1, Ucase(strSO), "64") <> 0 Then
       strSOArc = 64
    ElseIf instr(1, Ucase(strArc), 64) <> 0 Then
       strSOArc = 64
    Else
       strSOArc = 32
    End If
Next

If strSOArc = 64 Then
   strDevBin = "devconx64.exe"
Else
   strDevBin = "devcon.exe"
End If

Set oExec = WshShell.Exec(strScriptPath & "\" & strDevBin & " -FINDALL *")
objTextALL.Write oExec.StdOut.ReadAll
Set oExec1 = WshShell.Exec(strScriptPath & "\" & strDevBin & " -FIND *")
strPresents = oExec1.StdOut.ReadAll
objTextALL.Close

Set objTextALL = objFSO.OpenTextFile(strScriptPath & "\TEMP\ALL.txt", ForReading)
Do Until objTextALL.AtEndOfStream
strLine = objTextALL.ReadLine
If UCase(Left(strLine, 11)) <> "ROOT\LEGACY" Then
str1 = InStr(1, strLine, ":")
If str1 = 0 Then
   strDEVID = strLine
Else
   strDEVID = Left(strLine, str1 - 1)
End If
strFind = InStr(1, strPresents, strDEVID)
If strFind = 0 Then
   str2 = InStr(1, strLine, "matching device")
   str3 = InStr(1, strLine, "HTREE\ROOT\0")
   If str2 = 0 And str3 = 0 Then
      objTextNotPresent.WriteLine strLine
      objTextToRemove.WriteLine strDEVID
   End If
End If
End If
Loop
objTextNotPresent.Close
objTextToRemove.Close

'EXCLUI DISPOSITIVOS NOTPRESENTS
Set objTextLOG = objFSO.CreateTextFile(strScriptPath & "\LOG.txt", ForWriting)
Set objTextToRemove = objFSO.OpenTextFile(strScriptPath & "\TEMP\TOREMOVE.txt", ForReading)
Do Until objTextToRemove.AtEndOfStream
strToRemove = objTextToRemove.ReadLine
Set oExec2 = WshShell.Exec(strScriptPath & "\" & strDevBin & " remove @" & strToRemove)
strOUT = oExec2.StdOut.ReadAll
str4 = InStr(1, strOUT, "1 device(s) were removed")
Set objTextNotPresent = objFSO.OpenTextFile(strScriptPath & "\TEMP\NOTPRESENTS.txt", ForReading)
Do Until objTextNotPresent.AtEndOfStream
strLine1 = objTextNotPresent.ReadLine
str5 = InStr(1, strLine1, strToRemove)
If str5 <> 0 Then
   objTextLOG.Write strLine1
   Exit Do
End If
Loop

If str4 = 0 Then
   objTextLOG.WriteLine ";Não Removido!"
Else
   objTextLOG.WriteLine ";Removido com sucesso!"
End If

Loop
objTextALL.Close
objTextNotPresent.Close
objTextToRemove.Close
objTextLOG.Close

objFSO.DeleteFile (strScriptPath & "\TEMP\*.txt"), DeleteReadOnly

'MONTAGEM DO RELATÓRIO
Set objTextHTML = objFSO.CreateTextFile(strScriptPath & "\LOG.html", ForWriting)
objTextHTML.Write "<html><font face=""Georgia""><center><b><u>Relatório de Execução</u></b></center></font></br><table border=""1"" align=""center"" cellpadding=""8"" cellspacing=""1""><tr><th>ID</th><th>Descrição</th><th>Status</th></tr>"
Set objTextLOG = objFSO.OpenTextFile(strScriptPath & "\LOG.txt", ForReading)
Do Until objTextLOG.AtEndOfStream
   strLine2 = objTextLOG.ReadLine
   str1 = InStr(1, strLine2, ":")
   str2 = InStr(1, strLine2, ";")
   strID = Left(strLine2, str1 - 1)
   str3 = Left(strLine2, str2 - 1)
   str4 = Len(str3)
   strDesc = Right(str3, str4 - str1 - 1)
   str5 = Len(strLine2)
   strStatus = Right(strLine2, str5 - str2)
   If strStatus = "Não Removido!" Then
      objTextHTML.Write "<tr><td>" & strID & "</td><td>" & strDesc & "</td><td><font color=#FF0000>" & strStatus & "</font></td></tr>"
   Else
      objTextHTML.Write "<tr><td>" & strID & "</td><td>" & strDesc & "</td><td>" & strStatus & "</td></tr>"
   End If

Loop

oExec5 = WshShell.Run("iexplore.exe " & strScriptPath & "\LOG.html")

Wscript.Quit