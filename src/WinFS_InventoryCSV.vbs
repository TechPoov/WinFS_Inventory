Option Explicit

' ============================================
' WinFS_InventoryCSV_V1.1.vbs
' Multi-job scanning (Folders/Files/Both)
' Strict MODE validation
' Per-job CSV + LOG output
' Run-level summary file
' Metadata footer row
' ============================================

Const ForReading = 1
Const ForWriting = 2

Const MODE_FILES   = "FILES"
Const MODE_FOLDERS = "FOLDERS"
Const MODE_BOTH    = "BOTH"

Const DEBUGMODE = 0   ' 0 = Production, 1 = Debug

Const VERSION      = "1.1"
Const RELEASE_DATE = "2025-11-21"
Const UTILITY_NAME = "WinFS_InventoryCSV"

Dim g_logTS, g_logPath, g_sourcePath
Dim g_scriptFolder
Dim g_sumTS
Dim g_runTimestamp


' ----------------------------------------------------
' ENTRY POINT
' ----------------------------------------------------
Call Main()


Sub Main()
    Dim fso, scriptFullName, scriptFile, dotPos, baseName
    Dim configPath, jobs, jobKey
    Dim summaryPath
    Dim utilityBanner

    Set fso = CreateObject("Scripting.FileSystemObject")

    scriptFullName = WScript.ScriptFullName
    g_scriptFolder = fso.GetParentFolderName(scriptFullName)
    scriptFile     = fso.GetFileName(scriptFullName)

    dotPos = InStrRev(scriptFile, ".")
    If dotPos > 0 Then
        baseName = Left(scriptFile, dotPos - 1)
    Else
        baseName = scriptFile
    End If

    configPath = fso.BuildPath(g_scriptFolder, baseName & ".config")

    If Not fso.FileExists(configPath) Then
        WScript.Quit 1
    End If

    Set jobs = ParseMultiJobConfig(configPath)
    If jobs Is Nothing Then WScript.Quit 1
    If jobs.Count = 0 Then WScript.Quit 1

    g_runTimestamp = BuildTimestamp(Now)
    summaryPath = fso.BuildPath(g_scriptFolder, "scanutility_" & g_runTimestamp & ".csv")

    Set g_sumTS = fso.OpenTextFile(summaryPath, ForWriting, True, 0)
    g_sumTS.WriteLine "RunTimestamp,JobName,SourcePath,Mode,OutputFolder,Status,ItemsScanned,DataFile,LogFile,Details"

    For Each jobKey In jobs.Keys
        RunJob jobs(jobKey), jobKey
    Next

    utilityBanner = UTILITY_NAME & "_v" & VERSION & " - Completed (Release: " & RELEASE_DATE & ")"
    WriteSummaryRow g_runTimestamp, "UTILITY", "", "", g_scriptFolder, "COMPLETED", 0, "", "", utilityBanner

    g_sumTS.Close
End Sub



' ----------------------------------------------------
' RUN ONE JOB
' ----------------------------------------------------
Sub RunJob(jobDict, jobName)
    Dim fso, net
    Dim sourcePath, outputFolder, modeRaw, mode, email
    Dim computerName, lastFolder, timestamp, baseFileName
    Dim csvPath, jobLogPath
    Dim ts, startTime, endTime, counter
    Dim status, details, itemsScanned
    Dim utilityBanner

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set net = CreateObject("WScript.Network")

    sourcePath   = GetJobValue(jobDict, "SCANFOLDER", "")
    outputFolder = GetJobValue(jobDict, "OUTPUTFOLDER", "")
    modeRaw      = Trim(UCase(GetJobValue(jobDict, "MODE", "")))
    email        = GetJobValue(jobDict, "EMAIL", "")

    g_sourcePath = Trim(sourcePath)
    status       = "NOTRUN"
    details      = ""
    itemsScanned = 0
    csvPath      = ""
    jobLogPath   = ""

    If g_sourcePath = "" Then
        status  = "CONFIG_ERROR"
        details = "ScanFolder missing in config"
        WriteSummaryRow g_runTimestamp, jobName, sourcePath, modeRaw, outputFolder, status, itemsScanned, csvPath, jobLogPath, details
        Exit Sub
    End If

    If Not fso.FolderExists(g_sourcePath) Then
        status  = "CONFIG_ERROR"
        details = "ScanFolder does not exist"
        WriteSummaryRow g_runTimestamp, jobName, sourcePath, modeRaw, outputFolder, status, itemsScanned, csvPath, jobLogPath, details
        Exit Sub
    End If

    If Trim(outputFolder) = "" Then
        outputFolder = fso.BuildPath(g_scriptFolder, "Output")
    End If

    On Error Resume Next
    If Not fso.FolderExists(outputFolder) Then
        fso.CreateFolder(outputFolder)
        If Err.Number <> 0 Then
            status  = "ERROR"
            details = "Unable to create OutputFolder: " & outputFolder & " (" & Err.Description & ")"
            Err.Clear
            On Error GoTo 0
            WriteSummaryRow g_runTimestamp, jobName, sourcePath, modeRaw, outputFolder, status, itemsScanned, csvPath, jobLogPath, details
            Exit Sub
        End If
    End If
    On Error GoTo 0


    ' ----------------------------------------------------
    ' STRICT MODE VALIDATION
    ' ----------------------------------------------------
    Dim tmpLog, tmpLogPath

    Select Case modeRaw
        Case MODE_FILES
            mode = MODE_FILES
        Case MODE_FOLDERS
            mode = MODE_FOLDERS
        Case MODE_BOTH
            mode = MODE_BOTH

        Case Else
            status  = "CONFIG_ERROR"
            details = "Invalid MODE value: [" & modeRaw & "]. Allowed values: Files, Folders, Both."

            tmpLogPath = fso.BuildPath(outputFolder, _
                "INVALIDMODE_" & jobName & "_" & BuildTimestamp(Now) & "_log.csv")

            Set tmpLog = fso.OpenTextFile(tmpLogPath, ForWriting, True, 0)
            tmpLog.WriteLine "EventTime,SourcePath,Level,EventType,Message,Details"
            tmpLog.WriteLine BuildTimestamp(Now) & "," & _
                EscapeCSV(g_sourcePath) & ",ERROR,CONFIG,Invalid MODE," & _
                EscapeCSV(details)
            tmpLog.Close

            WriteSummaryRow g_runTimestamp, jobName, g_sourcePath, modeRaw, outputFolder, status, 0, "", tmpLogPath, details
            Exit Sub
    End Select
    ' ----------------------------------------------------



    computerName = net.ComputerName
    lastFolder   = GetLastFolderName(g_sourcePath)
    timestamp    = BuildTimestamp(Now)

    baseFileName = computerName & "_" & lastFolder & "_" & mode & "_" & timestamp

    csvPath    = fso.BuildPath(outputFolder, baseFileName & ".csv")
    jobLogPath = fso.BuildPath(outputFolder, baseFileName & "_log.csv")

    On Error Resume Next
    Set g_logTS = fso.OpenTextFile(jobLogPath, ForWriting, True, 0)
    If Err.Number <> 0 Then
        status  = "ERROR"
        details = "Cannot create job log file: " & jobLogPath & " (" & Err.Description & ")"
        Err.Clear
        On Error GoTo 0
        WriteSummaryRow g_runTimestamp, jobName, sourcePath, mode, outputFolder, status, itemsScanned, csvPath, jobLogPath, details
        Exit Sub
    End If
    On Error GoTo 0

    g_logTS.WriteLine "EventTime,SourcePath,Level,EventType,Message,Details"
    WriteLog "INFO", "START", "Job started: " & jobName, "Mode=" & mode

    On Error Resume Next
    Set ts = fso.OpenTextFile(csvPath, ForWriting, True, 0)
    If Err.Number <> 0 Then
        status  = "ERROR"
        details = "Cannot create data CSV: " & csvPath & " (" & Err.Description & ")"
        Err.Clear
        On Error GoTo 0
        WriteLog "ERROR", "DATA", "Unable to open data CSV", csvPath
        g_logTS.Close
        WriteSummaryRow g_runTimestamp, jobName, sourcePath, mode, outputFolder, status, itemsScanned, "", jobLogPath, details
        Exit Sub
    End If
    On Error GoTo 0

    ts.WriteLine "SlNo,ItemType,FullPath,Name,Extension,ParentFolder,SizeBytes,CreatedDate,ModifiedDate,Attributes"

   ' Dim startTime, endTime, counter
    startTime = Now
    counter   = 0

    ScanFolderRecursive g_sourcePath, mode, fso, ts, counter

    endTime = Now
    ts.Close

    itemsScanned = counter
    status       = "SUCCESS"
    details      = "Completed in " & DateDiff("s", startTime, endTime) & " seconds"

    WriteLog "INFO", "SUMMARY", "Total items scanned", "Items=" & itemsScanned
    WriteLog "INFO", "OUTPUT", "CSV created", csvPath
    WriteLog "INFO", "OUTPUT", "Log created", jobLogPath
    WriteLog "INFO", "END", "Job completed: " & jobName, details

    utilityBanner = UTILITY_NAME & "_v" & VERSION & " - Completed (Release: " & RELEASE_DATE & ")"
    WriteLog "INFO", "UTILITY", utilityBanner, ""

    g_logTS.Close

    WriteSummaryRow g_runTimestamp, jobName, g_sourcePath, mode, outputFolder, status, itemsScanned, csvPath, jobLogPath, details
End Sub



' ----------------------------------------------------
' SUMMARY ROW
' ----------------------------------------------------
Sub WriteSummaryRow(runTimestamp, jobName, sourcePath, mode, outputFolder, status, itemsScanned, dataFile, logFile, details)
    g_sumTS.WriteLine _
        runTimestamp & "," & _
        EscapeCSV(jobName) & "," & _
        EscapeCSV(sourcePath) & "," & _
        mode & "," & _
        EscapeCSV(outputFolder) & "," & _
        status & "," & _
        itemsScanned & "," & _
        EscapeCSV(dataFile) & "," & _
        EscapeCSV(logFile) & "," & _
        EscapeCSV(details)
End Sub



' ----------------------------------------------------
' LOG FUNCTION
' ----------------------------------------------------
Sub WriteLog(level, eventType, message, details)
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



' ----------------------------------------------------
' RECURSIVE SCAN
' ----------------------------------------------------
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

    If mode = MODE_FOLDERS Or mode = MODE_BOTH Then
        counter = counter + 1
        WriteFolderRow ts, counter, folder
    End If

    If mode = MODE_FILES Or mode = MODE_BOTH Then
        For Each fileObj In folder.Files
            counter = counter + 1
            WriteFileRow ts, counter, fileObj
            WriteLog "DEBUG", "FILE", "File scanned", fileObj.Path
        Next
    End If

    For Each subFolder In folder.SubFolders
        ScanFolderRecursive subFolder.Path, mode, fso, ts, counter
    Next
End Sub



' ----------------------------------------------------
' WRITE FOLDER ROW
' ----------------------------------------------------
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



' ----------------------------------------------------
' WRITE FILE ROW
' ----------------------------------------------------
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



' ----------------------------------------------------
' UTIL FUNCTIONS
' ----------------------------------------------------
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


Function BuildTimestamp(dt)
    Dim y, m, d, hh, nn
    y  = Year(dt)
    m  = Right("0" & Month(dt), 2)
    d  = Right("0" & Day(dt), 2)
    hh = Right("0" & Hour(dt), 2)
    nn = Right("0" & Minute(dt), 2)
    BuildTimestamp = y & m & d & "_" & hh & nn
End Function


Function GetLastFolderName(fullPath)
    Dim p
    fullPath = Trim(fullPath)

    If Right(fullPath, 1) = "\" Then
        fullPath = Left(fullPath, Len(fullPath) - 1)
    End If

    p = InStrRev(fullPath, "\")
    If p > 0 Then
        GetLastFolderName = Mid(fullPath, p + 1)
    Else
        GetLastFolderName = fullPath
    End If
End Function



' ----------------------------------------------------
' PARSE CONFIG (multi-job)
' ----------------------------------------------------
Function ParseMultiJobConfig(path)
    Dim fso, ts, line
    Dim currentSection, jobs, jobDict
    Dim pos, key, value

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts  = fso.OpenTextFile(path, ForReading)

    Set jobs = CreateObject("Scripting.Dictionary")
    Set jobDict = Nothing
    currentSection = ""

    Do While Not ts.AtEndOfStream
        line = Trim(ts.ReadLine)

        If line = "" Then
        ElseIf Left(line,1) = "#" Or Left(line,2) = "//" Then
        ElseIf Left(line,1) = "[" And Right(line,1) = "]" Then

            If Not jobDict Is Nothing Then
                If currentSection <> "" Then
                    jobs.Add currentSection, jobDict
                End If
            End If

            currentSection = Mid(line, 2, Len(line) - 2)
            Set jobDict = CreateObject("Scripting.Dictionary")

        Else
            pos = InStr(line, "=")
            If pos > 0 Then
                key   = UCase(Trim(Left(line, pos - 1)))
                value = Trim(Mid(line, pos + 1))

                If Not jobDict Is Nothing Then
                    jobDict(key) = value
                End If
            End If
        End If
    Loop

    ts.Close

    If Not jobDict Is Nothing Then
        If currentSection <> "" Then
            jobs.Add currentSection, jobDict
        End If
    End If

    Set ParseMultiJobConfig = jobs
End Function



Function GetJobValue(jobDict, keyName, defaultValue)
    Dim uKey
    uKey = UCase(keyName)
    If jobDict.Exists(uKey) Then
        GetJobValue = jobDict(uKey)
    Else
        GetJobValue = defaultValue
    End If
End Function
