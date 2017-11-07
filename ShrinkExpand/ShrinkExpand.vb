Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Timers
Public Class ShrinkExpand
    Dim dsFolders As New DataSet
    Dim ConfigPath As String
    Dim LogPath As String
    Dim LogFile As String
    Dim appName As String = "ShrinkExpand"
    Dim eventSource As String = appName + " Service"
    Dim appCompany As String = "FFXNZ"
    Dim CheckTimer As New Timer
    Dim TimerInterval As Integer = 30 'Default 30 Seconds
    Dim LMRegPath As String = "Software\" + appCompany + "\" + appName
    Dim HKLM As String = "HKEY_LOCAL_MACHINE\"
    Dim MaxLogFileSize As Integer = 10000 'Default is 10K
    Dim ServiceDirectory As String
    Dim myMessage As String = ""

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        AutoLog = False
        If Not EventLog.SourceExists(eventSource) Then
            EventLog.CreateEventSource(eventSource, appCompany)
        End If
        ' Configure the event log instance to use this source name
        EventLog1.Source = eventSource.ToString
        '
        ServiceDirectory = IO.Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory).ToString
        ConfigPath = (ServiceDirectory + "\ShrinkExpand-Folders.xml").ToString
        LogPath = (ServiceDirectory + "\Logs").ToString
        LogFile = LogPath + appName + ".log"

        Try
            If My.Computer.Registry.LocalMachine.OpenSubKey("Software\FFXNZ\ShrinkExpand") IsNot Nothing Then
                ConfigPath = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\Software\FFXNZ\ShrinkExpand", "ConfigFile", ConfigPath.ToString).ToString
                LogPath = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\Software\FFXNZ\ShrinkExpand", "LogPath", ConfigPath.ToString).ToString
                TimerInterval = CInt(My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\Software\FFXNZ\ShrinkExpand", "TimerInterval", TimerInterval).ToString)
                MaxLogFileSize = CInt(My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\Software\FFXNZ\ShrinkExpand", "LogSize", TimerInterval).ToString)
            Else
                My.Computer.Registry.LocalMachine.CreateSubKey("Software\FFXNZ\ShrinkExpand")
                My.Computer.Registry.SetValue("HKEY_LOCAL_MACHINE\Software\FFXNZ\ShrinkExpand", "ConfigFile", ConfigPath)
                My.Computer.Registry.SetValue("HKEY_LOCAL_MACHINE\Software\FFXNZ\ShrinkExpand", "LogPath", LogPath)
                My.Computer.Registry.SetValue("HKEY_LOCAL_MACHINE\Software\FFXNZ\ShrinkExpand", "TimerInterval", TimerInterval)
                My.Computer.Registry.SetValue("HKEY_LOCAL_MACHINE\Software\FFXNZ\ShrinkExpand", "LogSize", MaxLogFileSize)
            End If
        Finally
            My.Computer.Registry.LocalMachine.Close()
        End Try

    End Sub

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things in motion so your service can do its work.
        LogFile = LogPath + "\" + appName + ".Log"
        WriteLog("ShrinkExpand Starting")
        EventLog1.WriteEntry("ShrinkExpand Starting from " + IO.Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory).ToString)
        EventLog1.WriteEntry("ShrinkExpand Config File: " + ConfigPath, EventLogEntryType.Information)
        WriteLog("ShrinkExpand Config File: " + ConfigPath)
        If File.Exists(ConfigPath) Then
            dsFolders.Clear()
            dsFolders.ReadXml(ConfigPath)
            EventLog1.WriteEntry("Config Read")
            StartTimer()
        Else
            EventLog1.WriteEntry("Config File Missing:" + vbNewLine + ConfigPath)
            OnStop()
        End If
    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        EventLog1.WriteEntry("ShrinkExpand Stopping")
        WriteLog("ShrinkExpand Stopping")
        CheckTimer.Stop()
    End Sub

    ' Start the timer with the given interval
    Private Sub StartTimer()
        AddHandler CheckTimer.Elapsed, AddressOf CheckTimer_Tick
        CheckTimer.Interval = CInt(TimerInterval * 1000)
        CheckTimer.Start()
    End Sub

    Private Sub CheckTimer_Tick(ByVal obj As Object, ByVal e As EventArgs)
        '  writeLog("ShrinkExpand Tick")
        'EventLog1.WriteEntry("ShrinkExpand Tick", EventLogEntryType.Information)
        For Each dsRow As DataRow In dsFolders.Tables("Folder").Rows
            If dsRow.Item("Enabled") = "true" Then
                Dim Source As String = dsRow.Item("Source")
                Dim Destination As String = dsRow.Item("Destination")
                Dim Shrink As String = dsRow.Item("Shrink")
                'EventLog1.WriteEntry(("S: " + Source + vbNewLine + "D: " + Destination + vbNewLine + "P: " + Shrink).ToString, EventLogEntryType.Information)
                'EventLog1.WriteEntry("Process folder: " + Source + " As " + Shrink)
                If Shrink = "true" Then
                    CheckFolder(Source, Destination, True)
                Else
                    CheckFolder(Source, Destination, False)
                End If
            End If

        Next

    End Sub

    Private Sub WriteLog(Message As String)
        Dim timeMessage = Now.ToString("yyyyMMdd HH:mm:ss") + " - " + Message + vbNewLine
        If (My.Computer.FileSystem.FileExists(LogFile)) Then
            Dim logInfo As FileInfo = My.Computer.FileSystem.GetFileInfo(LogFile)
            If logInfo.Length > MaxLogFileSize Then
                Dim log5 As String = LogPath + "\" + appName + ".5.log"
                Dim log4 As String = LogPath + "\" + appName + ".4.log"
                Dim log3 As String = LogPath + "\" + appName + ".3.log"
                Dim log2 As String = LogPath + "\" + appName + ".2.log"
                Dim log1 As String = LogPath + "\" + appName + ".1.log"
                If (My.Computer.FileSystem.FileExists(log5)) Then
                    My.Computer.FileSystem.DeleteFile(log5)
                End If
                If (My.Computer.FileSystem.FileExists(log4)) Then
                    My.Computer.FileSystem.MoveFile(log4, log5)
                End If
                If (My.Computer.FileSystem.FileExists(log3)) Then
                    My.Computer.FileSystem.MoveFile(log3, log4)
                End If
                If (My.Computer.FileSystem.FileExists(log2)) Then
                    My.Computer.FileSystem.MoveFile(log2, log3)
                End If
                If (My.Computer.FileSystem.FileExists(log1)) Then
                    My.Computer.FileSystem.MoveFile(log1, log2)
                End If
                My.Computer.FileSystem.MoveFile(LogFile, log1)
                My.Computer.FileSystem.WriteAllText(LogFile, timeMessage, False)
            Else
                My.Computer.FileSystem.WriteAllText(LogFile, timeMessage, True)
            End If
        Else
            If Not My.Computer.FileSystem.DirectoryExists(LogPath) Then
                My.Computer.FileSystem.CreateDirectory(LogPath)
            End If
            My.Computer.FileSystem.WriteAllText(LogFile, timeMessage, False)
        End If

    End Sub

    Private Sub CheckFolder(SourcePath As String, DestinationPath As String, Shrink As Boolean)
        Dim files As ReadOnlyCollection(Of String)
        If Shrink Then
            'Do Subfolders
            files = My.Computer.FileSystem.GetFiles(SourcePath, FileIO.SearchOption.SearchAllSubDirectories, "*")
            For Each file In files
                'EventLog1.WriteEntry("Shrink - " + file, EventLogEntryType.Information)
                If IsNotLocked(file) Then
                    'EventLog1.WriteEntry(file + " is not in Use - Shrink", EventLogEntryType.Information)
                    ShrinkMe(file, SourcePath, DestinationPath)
                Else
                    EventLog1.WriteEntry(file + " in Use - Do Nothing", EventLogEntryType.Information)
                End If
            Next
        Else
            'Root Folder Only
            files = My.Computer.FileSystem.GetFiles(SourcePath, FileIO.SearchOption.SearchTopLevelOnly, "*")
            For Each file In files
                'EventLog1.WriteEntry("Expand - " + file, EventLogEntryType.Information)
                If IsNotLocked(file) Then
                    'EventLog1.WriteEntry(file + " is not in Use - Expand", EventLogEntryType.Information)
                    ExpandMe(file, SourcePath, DestinationPath)
                Else
                    EventLog1.WriteEntry(file + " in Use - Do Nothing", EventLogEntryType.Information)
                End If
            Next
        End If
    End Sub

    Private Function IsNotLocked(FilePath As String) As Boolean
        Dim stream As FileStream = Nothing
        Dim testFile As FileInfo = My.Computer.FileSystem.GetFileInfo(FilePath)
        Try
            stream = File.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None)
            stream.Close()
            Return False
        Catch ex As Exception
            Return True
        End Try
    End Function

    Private Sub ShrinkMe(SourceFile As String, FromDir As String, ToDir As String)
        Dim RelativePath As String = SourceFile.Replace(FromDir + "\", "")
        Dim DestinationFile As String = ToDir + "\" + RelativePath.Replace("\", "~")
        'EventLog1.WriteEntry(DestinationFile, EventLogEntryType.Information)
        MoveMe(SourceFile, DestinationFile)
    End Sub

    Private Sub ExpandMe(SourceFile As String, FromDir As String, ToDir As String)
        Dim RelativePath As String = SourceFile.Replace(FromDir + "\", "")
        Dim DestinationFile As String = ToDir + "\" + RelativePath.Replace("~", "\")
        'EventLog1.WriteEntry(DestinationFile, EventLogEntryType.Information)
        MoveMe(SourceFile, DestinationFile)
    End Sub

    Private Sub MoveMe(Source As String, Destination As String)
        Dim DestinationDirectory = Path.GetDirectoryName(Destination)
        If Not (Directory.Exists(DestinationDirectory)) Then
            Directory.CreateDirectory(DestinationDirectory)
        End If
        Try
            My.Computer.FileSystem.MoveFile(Source, Destination)
            'EventLog1.WriteEntry("Moved: " + vbNewLine + Source + vbNewLine + "To: " + vbNewLine + Destination + vbNewLine, EventLogEntryType.Information)
            WriteLog("Moved: " + Source + vbNewLine + "                         To: " + Destination)
        Catch ex As Exception
            EventLog1.WriteEntry(ex.Message + Source, EventLogEntryType.Error)
        End Try
    End Sub

End Class
