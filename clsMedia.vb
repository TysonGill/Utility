''' <summary>General library for MCI Multimedia routines.</summary>
Public Class clsMedia

    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer
    Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Integer, ByVal lpstrBuffer As String, ByVal uLength As Integer) As Integer

    Private MediaName As String

    ''' <summary>Play a media file</summary>
    ''' <param name="MediaPath">The complete path of the media file to play.</param>
    ''' <returns>A string if an error is encountered, otherwise an empty string.</returns>
    Public Function MediaPlay(ByVal MediaPath As String) As String

        ' If media is opened, close it
        If IsPlaying() Then MediaClose()
        mciSendString("close all", 0, 0, 0)

        ' Open the media file
        MediaName = IO.Path.GetFileNameWithoutExtension(MediaPath)
        Dim cmd As String = "open " + fmt.Quote(MediaPath) + " alias " + MediaName
        Dim retVal As Integer = mciSendString(cmd, 0, 0, 0)
        If retVal <> 0 Then Return GetErrorString(retVal)

        ' Play the media file
        retVal = mciSendString("play " + MediaName, 0, 0, 0)

        Return ""

    End Function

    ''' <summary>Start recording a media file</summary>
    ''' <returns>A string if an error is encountered, otherwise an empty string.</returns>
    Public Function MediaRecordStart() As String

        MediaName = "Recording"
        mciSendString("close all", 0, 0, 0)

        Dim cmd As String = "open new type waveaudio alias Recording"
        Dim retVal As Integer = mciSendString(cmd, 0, 0, 0)
        If retVal <> 0 Then Return GetErrorString(retVal)

        cmd = "set Recording format tag pcm"
        retVal = mciSendString(cmd, 0, 0, 0)
        If retVal <> 0 Then Return GetErrorString(retVal)

        cmd = "set Recording channels 2"
        retVal = mciSendString(cmd, 0, 0, 0)
        If retVal <> 0 Then Return GetErrorString(retVal)

        cmd = "set Recording samplespersec 44100"
        retVal = mciSendString(cmd, 0, 0, 0)
        If retVal <> 0 Then Return GetErrorString(retVal)

        cmd = "set Recording bitspersample 16"
        retVal = mciSendString(cmd, 0, 0, 0)
        If retVal <> 0 Then Return GetErrorString(retVal)

        ' Start recording the media file
        retVal = mciSendString("record Recording", 0, 0, 0)
        If retVal <> 0 Then Return GetErrorString(retVal)

        Return ""

    End Function

    ''' <summary>Record a media file</summary>
    ''' <param name="MediaPath">The complete path of the media file to create.</param>
    ''' <returns>A string if an error is encountered, otherwise an empty string.</returns>
    Public Function MediaRecordStop(ByVal MediaPath As String) As String

        MediaStop()

        ' Save the media file
        If MediaPath.Trim <> "" Then
            Dim cmd As String = "save Recording " + fmt.Quote(MediaPath)
            Dim retVal As Integer = mciSendString(cmd, 0, 0, 0)
            If retVal <> 0 Then Return GetErrorString(retVal)
        End If

        MediaClose()
        Return ""

    End Function

    ''' <summary>Pause playback of the current file.</summary>
    Public Sub MediaPause()
        mciSendString("pause " + MediaName, 0, 0, 0)
    End Sub

    ''' <summary>Resume playback of the current file.</summary>
    Public Sub MediaResume()
        mciSendString("resume " + MediaName, 0, 0, 0)
    End Sub

    ''' <summary>Stop playback of the current file.</summary>
    Public Sub MediaStop()
        mciSendString("stop " + MediaName, 0, 0, 0)
    End Sub

    ''' <summary>Check if a current file is open (playing or paused).</summary>
    Public Function IsPlaying() As Boolean
        Dim ReturnString As String = Space(128)
        mciSendString("status " + MediaName + " mode", ReturnString, ReturnString.Length, 0)
        Return (ReturnString.StartsWith("playing") Or ReturnString.StartsWith("paused"))
    End Function

    Private Sub MediaClose()
        mciSendString("close " + MediaName, 0, 0, 0)
    End Sub

    Private Function GetErrorString(ByVal retVal As Integer) As String
        If retVal = 0 Then Return ""
        Dim ErrorString As String = Space(128)
        mciGetErrorString(retVal, ErrorString, ErrorString.Length)
        Return ErrorString
    End Function

End Class
