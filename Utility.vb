Imports System.Windows.Forms

''' <summary>General reusable utility module.</summary>
Public Module Routines

    ' Public classes
    ''' <summary>Database Library.</summary>
    Public db As New clsDatabase
    ''' <summary>User Interface Library.</summary>
    Public ui As New clsUserInterface
    ''' <summary>Formatting Library.</summary>
    Public fmt As New clsFormatting
    ''' <summary>Mapping Library.</summary>
    Public map As New clsMapping
    ''' <summary>Networking Library.</summary>
    Public nwk As New clsNetworking

    ''' <summary>A structure to store information about the current user.</summary>
    ''' <remarks>See UserInfo</remarks>
    Public Structure UserInfoStruct
        Dim Environment As String
        Dim WindowsName As String
        Dim RoleID As Integer
        Dim Role As String
        Dim Department As String
        Dim IsDeptAdmin As Boolean
        Dim IsSuperAdmin As Boolean
        Dim LanguageGroup As String
        Dim LanguageCodes As String
        Dim dtRoles As DataTable
        Dim dtExtended As DataTable
    End Structure

    ''' <summary>Store information about the current user.</summary>
    ''' <remarks>Call SetUsereInfo to populate.</remarks>
    Public UserInfo As UserInfoStruct

    ''' <summary>Populates Utililty.UserInfo with details about the current Windows User.</summary>
    Public Sub SetUserInfo(Optional ByVal dqhr As clsDataQuery = Nothing)
        UserInfo.Department = ""
        UserInfo.RoleID = 0
        UserInfo.Role = ""
        UserInfo.IsDeptAdmin = False
        UserInfo.IsSuperAdmin = False
        UserInfo.LanguageGroup = ""
        UserInfo.LanguageCodes = ""
        UserInfo.WindowsName = nwk.GetUser()
        UserInfo.dtExtended = nwk.GetUserInfo(dqhr)
    End Sub

    ''' <summary>Allow the calling application to store it's last ping time here for error reporting,</summary>
    Public LastPing As Integer = -1

    ''' <summary>The query reporting threshold in milliseconds, defaults to 5000.</summary>
    Public LogThreshold As Integer = 5000

    ''' <summary>The query logging path. If empty or invalid then no logging.</summary>
    Public LogPath As String = ""

    ''' <summary>The event logging path. If empty or invalid then no logging.</summary>
    Public LogConnection As String = ""

    ''' <summary>Show and optionally mail error information.</summary>
    ''' <param name="exRep">The exception object to report.</param>
    ''' <param name="Fatal">Set to False to suppress exiting application after error.</param>
    ''' <remarks>Mail specifications are taken from My.Settings.</remarks>
    Public Sub ReportError(ByVal exRep As Exception, Optional ByVal Fatal As Boolean = False, Optional ByVal SendTo As String = "")
        Try
            ' Capture the date and time of error
            Dim ErrTime As String = DateTime.Now.ToLongDateString + " " + DateTime.Now.ToLongTimeString
            Dim m1 As Long = GC.GetTotalMemory(False) / 1000
            Dim m2 As Long = My.Computer.Info.AvailablePhysicalMemory() / 1000000

            ' Create html body content
            Dim sErr As String = ""
            If Fatal Then
                sErr = "Fatal Error " + Err.Number.ToString + " in: "
            Else
                sErr = "Non-Fatal Error " + Err.Number.ToString + " in: "
            End If
            sErr += Application.ProductName + " v" + Application.ProductVersion + vbCrLf
            sErr += "Error: " + exRep.Message + vbCrLf + vbCrLf

            sErr += "User: " + nwk.GetUser.ToString + vbCrLf
            sErr += "Date: " + ErrTime + vbCrLf
            If Not UserInfo.Role Is Nothing Then sErr += "Role: " + UserInfo.Role.ToString + vbCrLf
            If Not UserInfo.Environment Is Nothing Then sErr += "Environment: " + UserInfo.Environment.ToString + vbCrLf
            sErr += "Memory: " + m1.ToString("#,###,###") + "KB of " + m2.ToString("#,###,###") + "MB" + vbCrLf
            If LastPing = -1 Then
                sErr += "Ping: n/a" + vbCrLf
            Else
                sErr += "Ping: " + LastPing.ToString + " ms" + vbCrLf
            End If
            sErr += "Source: " + Err.Source + vbCrLf
            If exRep.TargetSite IsNot Nothing Then
                sErr += "Site: " + exRep.TargetSite.Name + vbCrLf + vbCrLf

                sErr += "Stack: " + vbCrLf
                sErr += Replace(exRep.StackTrace.ToString, " at", " at ")

                If Not (exRep.Data Is Nothing) Then
                    Dim de As DictionaryEntry
                    For Each de In exRep.Data
                        sErr += vbCrLf + de.Key + ": " + de.Value.ToString + vbCrLf
                    Next de
                    sErr += vbCrLf
                End If
            End If

            ' Display error report form
            Dim frm As New frmErrorReport
            frm.txtDesc.Text = sErr
            frm.lblFatal.Visible = Fatal
            frm.ShowDialog()

            ' Log error
            EventLogSave("Application Error", Microsoft.VisualBasic.Left(exRep.Message, 2000))

            ' Send mail message
            If SendTo.Trim <> "" Then nwk.MailError(SendTo, "eReport: " + Application.ProductName, sErr, True)

        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Error in ReportError", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If Fatal Then Application.Exit()
        End Try
    End Sub

    ''' <summary>Save event to database.</summary>
    ''' <param name="EventName">Name of the event to log.</param>
    ''' <param name="EventDesc">Optional description of the event.</param>
    ''' <param name="StartTime">Optional start time of event, defaults to Now.</param>
    ''' <param name="EndTime">Optional end time of event, defaults to Nothing.</param>
    ''' <param name="UserName">Optional User Name associated with the event, defaults to current user.</param>
    ''' <param name="AppName">Optional Application Name associated with the event, defaults to current user.</param>
    ''' <remarks>Note you must have an AppShareCon key in the Connections collection to use this method.</remarks>
    Public Sub EventLogSave(ByVal EventName As String, Optional ByVal EventDesc As String = "", Optional ByVal StartTime As DateTime = Nothing, Optional ByVal EndTime As DateTime = Nothing, Optional ByVal UserName As String = "", Optional ByVal AppName As String = "")

        If LogConnection.Trim = "" Then Exit Sub

        If UserName.Trim = "" Then UserName = nwk.GetUser
        If AppName.Trim = "" Then AppName = Application.ProductName

        Dim sql As String
        If StartTime = Nothing Then StartTime = Now
        sql = "INSERT INTO ApplicationEvents (AppName, EventName, EventDesc, UserName, EventStart, EventEnd) VALUES ("
        sql += fmt.q(AppName) + ", "
        sql += fmt.q(EventName) + ", " + fmt.q(EventDesc) + ", "
        sql += fmt.q(UserName) + ", "
        sql += fmt.q(StartTime) + ", " + fmt.q(EndTime) + ")"
        Dim dq As New clsDataQuery(LogConnection)
        dq.Execute(sql)

    End Sub

    ''' <summary>Free a disposable object for garbage collection.</summary>
    ''' <param name="obj">The disposable object to release.</param>
    ''' <remarks>Call when finished with any object that implements IDisposable.</remarks>
    Public Sub ReleaseObj(ByRef obj As Object)
        If obj Is Nothing Then Exit Sub
        Try
            obj.dispose()
        Finally
            obj = Nothing
        End Try
    End Sub

End Module
