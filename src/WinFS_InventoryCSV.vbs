Option Explicit

' ============================================
' WinFS_InventoryCSV_V0.7.vbs
' - Reads settings from <scriptname>.config
' - Produces Inventory CSV + Log CSV
' - Filenames: <System>_<LastFolder>_<Mode>_<Timestamp>.csv
'              <System>_<LastFolder>_<Mode>_<Timestamp>_log.csv
' - DEBUGMODE controls log verbosity
' ============================================

Const ForReading = 1
Const ForWriting = 2

Const MODE_FILES   = "FILES"
Const MODE_FOLDERS = "FOLDER"
Const MODE_BOTH    = "BOTH"

' 0 = production (only START/ERROR/WARN/SUMMARY/END/OUTPUT)
' 1 = debug (includes FILE/FOLDER-level logs)
Const DEBUGMODE = 1 '0

Dim g_logTS, g_logPath, g_sourcePath


' -------------------------
' ENTRY POINT
' -------------------------
Call Main()


Sub Main()
    Dim fso, scriptFullName, scriptFolder, scriptFile, dotPos, baseName
    Dim configPath, cfgScanFolder, cfgOutputFolder, cfgMode
    Dim outputFolder, mode, csvPath, baseFileName
    Dim net, computerName, ts, timestamp, lastFolder
    Dim startTime, endTime, counter

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' ---- Identify script folder + config ----
    scriptFullName = WScript.ScriptFullName
    scriptFolder   = fso.GetParentFolderName(scriptFullName)
    scriptFile     = fso.GetFileName(scriptFullName)

    dotPos = InStrRev(scriptFile, ".")
    If dotPos > 0 Then
        baseName = Left(scriptFile, dotPos - 1)
    Else
        baseName = scriptFile
    End If

    configPath = fso.BuildPath(scriptFolder, baseName & ".config")

    ' ---- Read configuration ----
    If Not fso.FileExists(configPath) Then
        ' No echo/log yet because log file not opened
        WScript.Quit 1
    End If

    If Not ReadConfig(configPath, cfgScanFolder, cfgOutputFolder, cfgMode) Then
        WScript.Quit 1
    End If

    g_sourcePath = Trim(cfgScanFolder)
    outputFolder = Trim(cfgOutputFolder)
    mode         = UCase(Trim(cfgMode))

    ' Validate scan path
    If g_sourcePath = "" Then WScript.Quit 1
    If Not fso.FolderExists(g_sourcePath) Then WScript.Quit 1

    ' Validate/create output folder
    If outputFolder = "" Then
        outputFolder = fso.BuildPath(scriptFolder, "Output")
    End If
    If Not fso.FolderExists(outputFolder) Then fso.CreateFolder(outputFolder)

    ' Validate mode
    Select Case mode
        Case MODE_FILES, MODE_FOLDERS, MODE_BOTH
        Case Else
            mode = MODE_BOTH
    End Select

    ' ---- Prepare naming components ----
    Set net = CreateObject("WScript.Network")
    computerName = net.ComputerName

    lastFolder = GetLastFolderName(g_sourcePath)
    timestamp  = BuildTimestamp(Now)  ' YYYYMMDD_HHMM

    baseFileName = computerName & "_" & lastFolder & "_" & mode & "_" & timestamp

    ' ---- Build filenames ----
    csvPath   = fso.BuildPath(outputFolder, baseFileName & ".csv")
    g_logPath = fso.BuildPath(outputFolder, baseFileName & "_log.csv")

    ' ---- Open log CSV ----
    Set g_logTS = fso.OpenTextFile(g_logPath, ForWriting, True, 0)
    g_logTS.WriteLine "EventTime,SourcePath,Level,EventType,Message,Details"

    WriteLog "INFO", "START", "Scan started", "Mode=" & mode

    ' ---- Open inventory CSV ----
    Set ts = fso.OpenTextFile(csvPath, ForWriting, True, 0)
    ts.WriteLine "SlNo,ItemType,FullPath,Name,Extension,ParentFolder,SizeBytes,CreatedDate,ModifiedDate,Attributes"

    ' ---- Begin scan ----
    startTime = Now
    counter = 0

    ScanFolderRecursive g_sourcePath, mode, fso, ts, counter

    endTime = Now
    ts.Close

    ' ---- Summary ----
    WriteLog "INFO", "SUMMARY", "Total scanned", "Items=" & counter
    WriteLog "INFO", "OUTPUT", "CSV created", csvPath
    WriteLog "INFO", "OUTPUT", "Log created", g_logPath
    WriteLog "INFO", "END", "Scan completed", _
             "Duration=" & DateDiff("s", startTime, endTime) & " seconds"

    g_logTS.Close
End Sub



' ----------------------------
' LOG FUNCTION (with DEBUG filter)
' ----------------------------
Sub WriteLog(level, eventType, message, details)
    ' Filter noisy events when DEBUGMODE = 0
    If DEBUGMODE = 0 Then
        If level = "DEBUG" Then Exit Sub
        If eventType = "FILE" Then Exit Sub
        If eventType = "FOLDER" Then Exit Sub
    End If

    g_logTS.WriteLine _
        BuildTimestamp(Now) & "," & _
        EscapeCSV(g_sourcePath) & "," & _
        level & "," & _
        eventType & "," & _
        EscapeCSV(message) & "," & _
        EscapeCSV(details)
End Sub



' ----------------------------
' RECURSIVE SCAN
' ----------------------------
Sub ScanFolderRecursive(currentPath, mode, fso, ts, ByRef counter)
    Dim folder, subFolder, fileObj

    On Error Resume Next
    Set folder = fso.GetFolder(currentPath)
    If Err.Number <> 0 Then
        WriteLog "WARN", "SKIP", "Access denied folder", currentPath
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    WriteLog "DEBUG", "FOLDER", "Scanning folder", currentPath

    ' --- Folder row ---
    If mode = MODE_FOLDERS Or mode = MODE_BOTH Then
        counter = counter + 1
        WriteFolderRow ts, counter, folder
    End If

    ' --- Files ---
    If mode = MODE_FILES Or mode = MODE_BOTH Then
        For Each fileObj In folder.Files
            counter = counter + 1
            WriteFileRow ts, counter, fileObj
            WriteLog "DEBUG", "FILE", "File scanned", fileObj.Path
        Next
    End If

    ' --- Recurse subfolders ---
    For Each subFolder In folder.SubFolders
        ScanFolderRecursive subFolder.Path, mode, fso, ts, counter
    Next
End Sub



' ----------------------------
' WRITE FOLDER CSV ROW
' ----------------------------
Sub WriteFolderRow(ts, slno, folder)
    ts.WriteLine Join(Array( _
        slno, _
        "Folder", _
        EscapeCSV(folder.Path), _
        EscapeCSV(folder.Name), _
        "", _
        EscapeCSV(folder.ParentFolder), _
        "", _
        EscapeCSV(CStr(folder.DateCreated)), _
        EscapeCSV(CStr(folder.DateLastModified)), _
        folder.Attributes _
    ), ",")
End Sub



' ----------------------------
' WRITE FILE CSV ROW
' ----------------------------
Sub WriteFileRow(ts, slno, fileObj)
    Dim ext
    ext = GetExtension(fileObj.Name)

    ts.WriteLine Join(Array( _
        slno, _
        "File", _
        EscapeCSV(fileObj.Path), _
        EscapeCSV(fileObj.Name), _
        EscapeCSV(ext), _
        EscapeCSV(fileObj.ParentFolder), _
        fileObj.Size, _
        EscapeCSV(CStr(fileObj.DateCreated)), _
        EscapeCSV(CStr(fileObj.DateLastModified)), _
        fileObj.Attributes _
    ), ",")
End Sub



' ----------------------------
' UTILITIES
' ----------------------------

Function GetExtension(name)
    Dim p
    p = InStrRev(name, ".")
    If p > 0 Then
        GetExtension = Mid(name, p + 1)
    Else
        GetExtension = ""
    End If
End Function


Function EscapeCSV(v)
    v = CStr(v)
    v = Replace(v, """", """""")
    EscapeCSV = """" & v & """"
End Function


' YYYYMMDD_HHMM (no seconds)
Function BuildTimestamp(dt)
    BuildTimestamp = _
        Year(dt) & _
        Right("0" & Month(dt), 2) & _
        Right("0" & Day(dt), 2) & "_" & _
        Right("0" & Hour(dt), 2) & _
        Right("0" & Minute(dt), 2)
End Function


' Get LAST folder name
' Example: C:\Tools\Utilities\  → Utilities
Function GetLastFolderName(fullPath)
    fullPath = Trim(fullPath)

    If Right(fullPath, 1) = "\" Then
        fullPath = Left(fullPath, Len(fullPath) - 1)
    End If

    Dim p
    p = InStrRev(fullPath, "\")
    If p > 0 Then
        GetLastFolderName = Mid(fullPath, p + 1)
    Else
        GetLastFolderName = fullPath
    End If
End Function



' -----------------------------------
' CONFIG PARSER
' -----------------------------------
Function ReadConfig(path, ByRef scan, ByRef output, ByRef mode)
    Dim fso, ts, line, key, value, pos

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts  = fso.OpenTextFile(path, ForReading)

    Do While Not ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If line <> "" Then
            If Left(line,1) <> "#" And Left(line,2) <> "//" Then
                pos = InStr(line, "=")
                If pos > 0 Then
                    key   = Trim(Left(line, pos - 1))
                    value = Trim(Mid(line, pos + 1))

                    Select Case UCase(key)
                        Case "SCANFOLDER":   scan  = value
                        Case "OUTPUTFOLDER": output = value
                        Case "MODE":         mode  = value
                    End Select
                End If
            End If
        End If
    Loop

    ts.Close
    ReadConfig = True
End Function
