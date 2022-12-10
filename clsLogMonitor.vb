Imports System.IO
Imports System.Threading.Tasks

''' <summary>Monitors a text file for changes.</summary>
Public Class LogMonitor

    Private logStream As FileStream = Nothing
    Private logReader As StreamReader = Nothing
    Private _LogFile As String = ""
    Private _MonitorMS As Integer = 2000
    Private LogTask As Task = Nothing

    Private Started As Boolean = False
    ''' <summary>Flags whether logging has started.</summary>
    Public ReadOnly Property IsStarted() As Integer
        Get
            Return Started
        End Get
    End Property

    Private Running As Boolean = False
    ''' <summary>Flags whether logging is currently running.</summary>
    Public ReadOnly Property IsRunning() As Integer
        Get
            Return Running
        End Get
    End Property

    ''' <summary>Monitors a text file for changes.</summary>
    Public Event LogUpdated(ByVal NewLine As String)

    ''' <summary>Initializes a new LogMonitor class.</summary>
    ''' <param name="LogFile">The full path of the file to monitor.</param>
    ''' <param name="MonitorMS">The monitoring rate in milliseconds of delay.</param>
    Public Sub New(ByVal LogFile As String, Optional ByVal MonitorMS As Integer = 2000)
        _LogFile = LogFile
        _MonitorMS = MonitorMS
    End Sub

    ''' <summary>Starts logging of the file.</summary>
    Public Function StartLogging() As String

        StopLogging()

        ' Open the stream, reader, and read the initial log contents
        logStream = New FileStream(_LogFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
        logReader = New StreamReader(logStream)
        Dim InitLog As String = logReader.ReadToEnd()

        ' Start a thread to watch for updates
        Started = True
        Running = True
        LogTask = Task.Factory.StartNew(Sub() Monitor())

        ' Return the initial log contents
        Return InitLog

    End Function

    ''' <summary>Pauses logging of the file.</summary>
    Public Sub PauseLogging()
        If Not Started Then Exit Sub
        Running = False
    End Sub

    ''' <summary>Continues logging after a pause.</summary>
    Public Sub ContinueLogging()
        If Not Started Then Exit Sub
        Running = True
    End Sub

    ''' <summary>Stops logging.</summary>
    Public Sub StopLogging()
        If Not Started Then Exit Sub
        Running = False
        Started = False
        If logReader IsNot Nothing Then
            logReader.Close()
            logReader.Dispose()
        End If
        If logStream IsNot Nothing Then
            logStream.Close()
            logStream.Dispose()
        End If
    End Sub

    ' Check for new log lines
    Private Sub Monitor()
        Do While Started
            If Running Then ReadNewLines()
            Threading.Thread.Sleep(_MonitorMS)
        Loop
    End Sub

    ' Read all new log lines
    Private Sub ReadNewLines()
        While Not logReader.EndOfStream
            RaiseEvent LogUpdated(logReader.ReadLine)
        End While
    End Sub

End Class
