Imports System
Imports System.DirectoryServices
Imports System.DirectoryServices.ActiveDirectory
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Printing
Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Threading
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms
Imports System.Xml
Imports Microsoft.Win32
Imports System.Security.Principal
Imports System.Security.Permissions

''' <summary>General library for networking routines.</summary>
Public Class clsNetworking

#Region "Networking"

    ''' <summary>Returns the DNS name of the host machine.</summary>
    Public Function HostNameGet() As String
        Return Dns.GetHostName
    End Function

    ''' <summary>Returns the IP Address for the most machine.</summary>
    Public Function HostAddressGet() As String
        Dim IpAddr As IPHostEntry = Dns.GetHostEntry(HostNameGet)
        Return IpAddr.AddressList(0).ToString
    End Function

    ''' <summary>Returns the user name for the current windows login.</summary>
    ''' <returns>The user name.</returns>
    ''' <remarks>Requires that application be configured to use Windows Authentication and also a connection to the HR database called HRCon in the Connections collection.</remarks>
    Public Function GetUser() As String
        Dim UserName As String = My.User.Name
        If InStr(UserName, "\") > 0 Then
            Dim s() As String
            s = Split(UserName, "\")
            Return s(1).Trim
        Else
            Return UserName.Trim
        End If
    End Function

    ''' <summary>Returns detailed user information.</summary>
    ''' <param name="dqHR">A DataQuery object pointing to the HR database.</param>
    ''' <param name="UserName">The windows system name or the current user if omitted.</param>
    ''' <returns>A datatable of user information.</returns>
    ''' <remarks>
    ''' Requires that application be configured to use Windows Authentication. 
    ''' Note that multiple records can be returned if the person has multiple phone numbers.
    ''' </remarks>
    Public Function GetUserInfo(ByVal dqHR As clsDataQuery, Optional ByVal UserName As String = "") As DataTable
        If dqHR Is Nothing Then Return Nothing
        Try
            UserName = UserName.Trim
            If UserName = "" Then UserName = nwk.GetUser

            ' Get the email address to look up in the HR database
            Dim Email As String = GetUserProperty("mail", UserName).ToString.Trim
            If Email = "" Then Email = "return an empty datatable"

            ' Retrieve the user information
            Dim sql As String = "SELECT DISTINCT CAST('' AS VARCHAR(50)) AS [User Name], FirstName, MiddleName, LastName, NickName AS [Nick Name], CAST('' AS Varchar(100)) AS [Full Name], EmployeeStatus AS [Status], PositionCodeDescription AS [Position], OrganizationDescription AS [Organization], PersonPhoneNo AS [Phone Number], SupervisorLastName, SupervisorFirstName, CAST('' AS Varchar(100)) AS [Supervisor], EmailAddress AS [Email Address], DepartmentDescription AS [Department], LocationDescription AS [Location] from dv_office_phones"
            sql += " WHERE emailaddress = @Email"
            dqHR.ParamSet("@Email", Email, True)
            GetUserInfo = dqHR.GetTable("User Information", sql)

            ' If no data was returned, just create a record and populate with empty strings
            If GetUserInfo.Rows.Count = 0 Then
                GetUserInfo.Rows.Add("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            End If

            ' Format fields
            For Each r As DataRow In GetUserInfo.Rows
                r("User Name") = UserName
                r("Full Name") = fmt.ToFullName(r("FirstName").ToString, r("MiddleName").ToString, r("LastName").ToString)
                r("Supervisor") = fmt.ToFullName(r("SupervisorFirstName").ToString, "", r("SupervisorLastName").ToString)
                r("Phone Number") = fmt.ToPhone(r("Phone Number").ToString)
            Next
            GetUserInfo.Columns.Remove("FirstName")
            GetUserInfo.Columns.Remove("MiddleName")
            GetUserInfo.Columns.Remove("LastName")
            GetUserInfo.Columns.Remove("SupervisorFirstName")
            GetUserInfo.Columns.Remove("SupervisorLastName")
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

#End Region

#Region "File Handling"

    Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, ByRef cbRemoteName As Integer) As Integer

    ''' <summary>Converts a local (mapped) path to a UNC path.</summary>
    ''' <param name="LocalPath">The local path to convert, must start with drive letter followed by a colon.</param>
    ''' <returns>The UNC path.</returns>
    ''' <remarks>It will return the original path if the conversion cannot be performed.</remarks>
    Public Function ToUncPath(ByVal LocalPath As String) As String

        ' Validate the parameter
        LocalPath = LocalPath.Trim
        If LocalPath.StartsWith("\\") Then Return LocalPath
        If LocalPath.Length = 0 Then Return LocalPath
        If LocalPath.Length = 1 Then LocalPath += ":"
        If LocalPath.Substring(1, 1) <> ":" Then Return LocalPath

        ' Get the UNC Path
        Dim UncPath As String = New String(" ", 520)
        If WNetGetConnection(LocalPath.Substring(0, 2), UncPath, 520) <> 0 Then Return LocalPath

        ' Format the response
        UncPath = UncPath.Substring(0, UncPath.IndexOf(vbNullChar)).Trim
        If UncPath.Length = 0 Then Return LocalPath
        Return UncPath + LocalPath.Substring(2, LocalPath.Length - 2)

    End Function

    ''' <summary>Check if a path point to a file or a URL.</summary>
    ''' <param name="Path">The URL or file path.</param>
    ''' <returns>Returns True if Path is a URL.</returns>
    Public Function IsURL(ByVal Path As String) As Boolean
        Path = Path.Trim
        If Path.ToLower().StartsWith("http://") Or Path.ToLower().StartsWith("www.") Then Return True
        Return False
    End Function

    ''' <summary>Returns an image pointed to by a URL or File Path.</summary>
    ''' <param name="Path">The URL or file path to the image.</param>
    ''' <returns>The Image.</returns>
    Public Function GetImage(ByVal Path As String) As Image
        If IsURL(Path) Then
            GetImage = GetWebImage(Path)
        Else
            GetImage = GetFileImage(Path)
        End If
    End Function

    ''' <summary>Returns an image pointed to by a URL.</summary>
    ''' <param name="ImageURL">The URL to the image on the web.</param>
    ''' <returns>The Image.</returns>
    Public Function GetWebImage(ByVal ImageURL As String) As Image
        Dim objImage As MemoryStream
        Dim objwebClient As WebClient
        ImageURL = Trim(ImageURL)
        If Not ImageURL.ToLower().StartsWith("http://") Then ImageURL = "http://" & ImageURL
        objwebClient = New WebClient()
        objImage = New MemoryStream(objwebClient.DownloadData(ImageURL))
        GetWebImage = Image.FromStream(objImage)
    End Function

    ''' <summary>Returns an image pointed to by a file path.</summary>
    ''' <param name="ImagePath">The path the file image file.</param>
    ''' <returns>The Image.</returns>
    Public Function GetFileImage(ByVal ImagePath As String) As Image
        Try
            GetFileImage = New System.Drawing.Bitmap(ImagePath)
        Catch
            Throw New Exception("Could not load the image: " + ImagePath)
        End Try
    End Function

    ''' <summary>Returns the oldest file in a specified file path.</summary>
    ''' <param name="FilePath">The path to search.</param>
    ''' <returns>The file name.</returns>
    Public Function GetOldestFile(ByVal FilePath As String) As String

        Dim FileList() As String = Directory.GetFiles(FilePath)
        If FileList.Length = 0 Then Return ""

        ' Get oldest file
        Dim InFile As String = ""
        Dim OldestDate As DateTime = DateTime.MaxValue
        For Each sFile As String In FileList
            If File.GetCreationTime(sFile) < OldestDate Then
                InFile = sFile
                OldestDate = File.GetCreationTime(sFile)
            End If
        Next
        Return InFile

    End Function

    ''' <summary>Returns the newest file in a specified file path.</summary>
    ''' <param name="FilePath">The path to search.</param>
    ''' <returns>The file name.</returns>
    Public Function GetNewestFile(ByVal FilePath As String) As String

        Dim FileList() As String = Directory.GetFiles(FilePath)
        If FileList.Length = 0 Then Return ""

        ' Get newest file
        Dim InFile As String = ""
        Dim NewestDate As DateTime = DateTime.MinValue
        For Each sFile As String In FileList
            If File.GetCreationTime(sFile) > NewestDate Then
                InFile = sFile
                NewestDate = File.GetCreationTime(sFile)
            End If
        Next
        Return InFile

    End Function

    ''' <summary>AppendAllText with retry to avoid locking issues.</summary>
    ''' <param name="FileName">The file name to append.</param>
    ''' <param name="Text">The text to append.</param>
    ''' <param name="TimeOutSecs">Optional timeout in seconds.</param>
    ''' <remarks>Will raise an error if unsuccessful.</remarks>
    Public Sub AppendText(ByVal FileName As String, ByVal Text As String, Optional ByVal TimeOutSecs As Integer = 5)
        Dim StartTime As DateTime = Now
        Do While DateDiff(DateInterval.Second, StartTime, Now) < TimeOutSecs
            Try
                File.AppendAllText(FileName, Text)
                Return
            Catch ex As IOException
                Application.DoEvents()
            End Try
        Loop
        ' Retry time exceeded, throw an IO Error
        Throw New IOException
    End Sub

    ''' <summary>Enumerates the Modes for the DirectoryDelete method.</summary>
    Public Enum DeleteMode
        DeleteAll
        DeleteFilesOnly
        DeleteSubfoldersOnly
    End Enum

    ''' <summary>
    ''' Recursively delete all files in a folder
    ''' </summary>
    ''' <param name="TopFolder">The file path to delete</param>
    ''' <param name="Mode">Specifies whether to delete the folder entirely or clean up contents only</param>
    ''' <remarks>Overrides any readonly permission settings</remarks>
    Public Sub DirectoryDelete(ByVal TopFolder As String, Optional ByVal Mode As DeleteMode = DeleteMode.DeleteAll)
            DirectoryDeleteInt(TopFolder, Mode)
            If Mode = DeleteMode.DeleteSubfoldersOnly Then Directory.CreateDirectory(TopFolder)
    End Sub

    Private Sub DirectoryDeleteInt(ByVal TopFolder As String, Optional ByVal Mode As DeleteMode = DeleteMode.DeleteAll)

        If Not Directory.Exists(TopFolder) Then Exit Sub

        ' Delete all files from the Directory
        Dim Files As FileInfo() = New DirectoryInfo(TopFolder).GetFiles("*.*")
        For Each Info As FileInfo In Files
            If (Info.Attributes And FileAttributes.ReadOnly) Then Info.Attributes = (Info.Attributes And Not FileAttributes.ReadOnly)
            File.Delete(Info.FullName)
        Next

        ' Process all child directories
        For Each Folder As String In Directory.GetDirectories(TopFolder)
            DirectoryDeleteInt(Folder, Mode)
        Next

        ' Delete the parent directory
        If Mode <> DeleteMode.DeleteFilesOnly Then Directory.Delete(TopFolder)

    End Sub

#End Region

#Region "Screen Capture"

    ''' <summary>Gets a screen shot.</summary>
    ''' <returns>An image of the current screen.</returns>
    Public Function CaptureScreen() As Image
        Dim ms As New MemoryStream
        Try
            Dim scr As System.Drawing.Rectangle = Screen.PrimaryScreen.Bounds
            Dim b As New Bitmap(scr.Width, scr.Height, PixelFormat.Format32bppArgb)
            Dim g As Graphics = Graphics.FromImage(b)
            g.CopyFromScreen(0, 0, 0, 0, scr.Size)
            b.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            ms.Seek(0, SeekOrigin.Begin)
            CaptureScreen = Image.FromStream(ms)
            ReleaseObj(g)
            ReleaseObj(b)
        Finally
            ReleaseObj(ms)
        End Try
    End Function

    Declare Function BitBlt Lib "gdi32.dll" ( _
        ByVal hDestDC As IntPtr, _
        ByVal x As Int32, _
        ByVal y As Int32, _
        ByVal nWidth As Int32, _
        ByVal nHeight As Int32, _
        ByVal hSrcDC As IntPtr, _
        ByVal xSrc As Int32, _
        ByVal ySrc As Int32, _
        ByVal dwRop As Int32) As Int32

    ''' <summary>Gets a screen shot of an application form.</summary>
    ''' <param name="frm">The form to capture.</param>
    ''' <returns>An image of the form.</returns>
    Public Function CaptureForm(ByVal frm As Form) As Image
        Dim g As Graphics = frm.CreateGraphics()
        Dim b = New Bitmap(frm.ClientRectangle.Width, frm.ClientRectangle.Height, g)
        Dim g2 As Graphics = Graphics.FromImage(b)
        Dim dc1 As IntPtr = g.GetHdc()
        Dim dc2 As IntPtr = g2.GetHdc()
        BitBlt(dc2, 0, 0, frm.ClientRectangle.Width, frm.ClientRectangle.Height, dc1, 0, 0, 13369376)
        g.ReleaseHdc(dc1)
        g2.ReleaseHdc(dc2)
        CaptureForm = b
    End Function

#End Region

#Region "Email"

    ''' <summary>Send an email message using from the local system.</summary>
    ''' <param name="SendTo">Mail destination address.</param>
    ''' <param name="Subject">The subject line of the email message.</param>
    ''' <param name="Body">HTML Body of the email message.</param>
    ''' <param name="Attachments">Optional array of attachment file network path specifications.</param>
    ''' <param name="SendAs">Optional sender address.</param>
    Public Function MailSend(ByVal SendTo As String, ByVal Subject As String, ByVal Body As String, Optional ByVal Attachments() As String = Nothing, Optional ByVal SendAs As String = "automailer@healthfirst.org", Optional ByVal CopyToSentFolder As Boolean = False) As String
        Try
            ' Validate mail parameters
            Dim MailHost As String = "sv-hfi-exchange.hfms.healthfirst.org"

            Dim MailPort As Integer = 0
            Dim MailUser As String = ""
            Dim MailPass As String = ""
            Dim MailSecure As Boolean = False

            ' Create mail message and smtp client
            Dim smtp As New SmtpClient(MailHost, MailPort)
            smtp.Credentials = New System.Net.NetworkCredential(MailUser, MailPass)
            smtp.EnableSsl = MailSecure
            If CopyToSentFolder Then smtp.DeliveryMethod = SmtpDeliveryMethod.Network

            ' Address mail
            Dim mail As New MailMessage()
            mail.From = New MailAddress(SendAs)
            Dim recepient As String
            For Each recepient In Split(SendTo, ";")
                If recepient.Trim <> "" Then mail.To.Add(recepient)
            Next
            mail.Subject = Subject
            mail.IsBodyHtml = True

            ' Create plain view body
            Dim plainView As AlternateView = AlternateView.CreateAlternateViewFromString(Body, Nothing, "text/plain")
            mail.AlternateViews.Add(plainView)

            ' Create HTML view body
            Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(Body, Nothing, "text/html")
            mail.AlternateViews.Add(htmlView)

            ' Add any attachments
            If Not Attachments Is Nothing Then
                For Each sFile As String In Attachments
                    mail.Attachments.Add(New Attachment(sFile))
                Next
            End If

            ' Send mail message
            smtp.Send(mail)

            ' Release any disposable objects
            ReleaseObj(plainView)
            ReleaseObj(htmlView)
            ReleaseObj(mail)
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    ''' <summary>Send an email report of an error.</summary>
    ''' <param name="SendTo">Mail destination address.</param>
    ''' <param name="Subject">The subject line of the email message.</param>
    ''' <param name="HTMLBody">HTML Body of the email message.</param>
    ''' <param name="Screenshot">True to include a screenshot.</param>
    Public Sub MailError(ByVal SendTo As String, ByVal Subject As String, ByVal HTMLBody As String, Optional ByVal Screenshot As Boolean = False)
        Try
            ' Validate mail parameters
            Dim MailHost As String = "sv-hfi-exchange.hfms.healthfirst.org"
            Dim MailPort As Integer = 0
            Dim MailUser As String = ""
            Dim MailPass As String = ""
            Dim MailSecure As Boolean = False

            ' Create mail message and smtp client
            Dim smtp As New SmtpClient(MailHost, MailPort)
            smtp.Credentials = New System.Net.NetworkCredential(MailUser, MailPass)
            smtp.EnableSsl = MailSecure

            ' Address mail
            Dim mail As New MailMessage()
            mail.From = New MailAddress("automailer@healthfirst.org")
            Dim recepient As String
            For Each recepient In Split(SendTo, ";")
                If recepient.Trim <> "" Then mail.To.Add(recepient)
            Next
            mail.Subject = Subject
            mail.IsBodyHtml = True

            ' Create plain view body
            Dim plainView As AlternateView = AlternateView.CreateAlternateViewFromString("Screen Shot", Nothing, "text/plain")
            mail.AlternateViews.Add(plainView)

            ' Create HTML view body
            Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(HTMLBody, Nothing, "text/html")

            ' Add screen shot to HTML view
            Dim ms As New MemoryStream
            If Screenshot Then
                Dim scr As System.Drawing.Rectangle = Screen.PrimaryScreen.Bounds
                Dim b As New Bitmap(scr.Width, scr.Height, PixelFormat.Format32bppArgb)
                Dim g As Graphics = Graphics.FromImage(b)
                g.CopyFromScreen(0, 0, 0, 0, scr.Size)
                b.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
                ms.Seek(0, SeekOrigin.Begin)
                Dim rsc As New LinkedResource(ms, MediaTypeNames.Image.Jpeg)
                HTMLBody += "<br><br><img src=cid:screenshot>"
                rsc.ContentId = "screenshot"
                htmlView.LinkedResources.Add(rsc)
                ReleaseObj(g)
                ReleaseObj(b)
            End If

            ' Link up view
            mail.AlternateViews.Add(htmlView)

            ' Send mail message
            smtp.Send(mail)

            ' Release any disposable objects
            ReleaseObj(ms)
            ReleaseObj(plainView)
            ReleaseObj(htmlView)
            ReleaseObj(mail)
        Catch ex As Exception
            ReportError(ex)
        End Try
    End Sub

    ''' <summary>Validates an email address</summary>
    ''' <param name="EmailAddress">The email address to validate</param>
    ''' <param name="TryConnect">False to check format only, True to attempt to get an SMTP response.</param>
    ''' <returns>An empty string on success or an error string describing the failure.</returns>
    Public Function EmailValidate(ByVal EmailAddress As String, Optional ByVal TryConnect As Boolean = False) As String

        ' Validate Email address for correct format
        Dim pattern As String = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
        If Not System.Text.RegularExpressions.Regex.IsMatch(EmailAddress, pattern) Then Return "Bad Format"

        If TryConnect Then

            ' Get the mail server
            Dim DomainName As String = Split(EmailAddress, "@")(1)
            Dim MailServer As String = NSLookup(DomainName)
            If DomainName.ToUpper <> "HEALTHFIRST.ORG" AndAlso MailServer = "" Then Return "Mail Server Lookup Failed"

            ' Try to connect to the server and send to the address
            Dim oStream As NetworkStream = Nothing
            Dim oConnection As New TcpClient()
            Dim sResponse As String
            Dim MyEmailAddress As String = "automailer@healthfirst.org"
            Try
                ' Create a connection stream
                oConnection.SendTimeout = 3000
                oConnection.Connect(MailServer, 25)
                oStream = oConnection.GetStream()
                sResponse = GetData(oStream)

                'Collect Response
                sResponse = SendData(oStream, "HELO healthfirst.org" + vbCrLf)
                If Not ValidResponse(sResponse) Then Return sResponse
                sResponse = SendData(oStream, "MAIL FROM: <" + MyEmailAddress + ">" + vbCrLf)
                If Not ValidResponse(sResponse) Then Return sResponse
                sResponse = SendData(oStream, "RCPT TO: " & EmailAddress + vbCrLf)
                If Not ValidResponse(sResponse) Then Return sResponse
                SendData(oStream, "QUIT" & vbCrLf)
                Return sResponse

            Catch ex As Exception
                Return "The email host did not respond."
            Finally
                oConnection.Close()
                ReleaseObj(oStream)
            End Try
        End If
        Return ""
    End Function

    Private Function SendData(ByRef oStream As NetworkStream, ByVal sToSend As String) As String
        Dim bArray() As Byte = Encoding.ASCII.GetBytes(sToSend.ToCharArray)
        oStream.Write(bArray, 0, bArray.Length())
        Return GetData(oStream)
    End Function

    Private Function GetData(ByRef oStream As NetworkStream) As String
        Dim bResponse(1024) As Byte
        Dim lenStream As Integer = oStream.Read(bResponse, 0, 1024)
        Dim sResponse As String = ""
        If lenStream > 0 Then
            sResponse = Encoding.ASCII.GetString(bResponse, 0, 1024)
        End If
        Return sResponse
    End Function

    Private Function ValidResponse(ByVal sResult As String) As Boolean
        Return (sResult.StartsWith("250 "))
    End Function

    Private Function NSLookup(ByVal sDomain As String) As String
        Dim info As New ProcessStartInfo()
        info.UseShellExecute = False
        info.CreateNoWindow = True
        info.RedirectStandardInput = True
        info.RedirectStandardOutput = True
        info.FileName = "nslookup"
        info.Arguments = "-type=MX " + sDomain.ToUpper.Trim

        Dim ns As Process
        ns = Process.Start(info)
        Dim sout As StreamReader
        sout = ns.StandardOutput
        Dim reg As Regex = New Regex("mail exchanger = (?<server>[^\\\s]+)")
        Dim mailserver As String = ""
        Dim response As String = ""
        Do While (sout.Peek() > -1)
            response = sout.ReadLine()
            Dim amatch As Match = reg.Match(response)
            If (amatch.Success) Then
                mailserver = amatch.Groups("server").Value
                Exit Do
            End If
        Loop
        Return mailserver
    End Function

#End Region

#Region "Interactive"

    ''' <summary>Listens for messages from the Interactive call handling system.</summary>
    ''' <remarks>Use WithEvents when you declare an InteractiveListener and handle the InteractiveReceived event.</remarks>
    ''' <example>
    ''' Private WithEvents Interactive As New clsNetworking.InteractiveListener
    ''' Interactive.IsListening = True
    ''' </example>
    Public Class InteractiveListener

        ''' <summary>Raised when a message is received from the call handling system.</summary>
        ''' <param name="dt">The datatable containing information received from the Interactive system.</param>
        Public Event InteractiveReceived(ByVal dt As DataTable)

        ''' <summary>Set to True to monitor Interactive calls raise InteractiveReceived events.</summary>
        Public IsListening As Boolean = False

        ' A new thread will be started to monitor incoming messages
        Private th As Thread
        Private Port As Integer = 0

        ''' <summary>Instantiate a new Interactive listener.</summary>
        ''' <param name="ListeningPort"></param>
        Public Sub New(ByVal ListeningPort As Integer)
            Port = ListeningPort
            th = New Thread(AddressOf Me.InteractiveListener)
            th.IsBackground = True
            th.Start()
        End Sub

        Private Sub InteractiveListener()
            Dim server As TcpListener = Nothing

            Try
                server = New TcpListener(IPAddress.Any, Port)
                server.Start()

                ' Declare reusable variables
                Dim blen As Integer
                Dim bytes(1024) As Byte
                Dim doc As New XmlDocument

                ' Enter the listening loop
                While True

                    ' Do nothing while listening is not enabled
                    Do While Not IsListening
                        Thread.Sleep(1000)
                    Loop

                    ' Perform a blocking call to accept requests
                    Dim client As TcpClient = server.AcceptTcpClient()
                    Dim xmlMessage As String = String.Empty

                    ' Get a stream object for reading and writing
                    Dim stream As NetworkStream = client.GetStream()

                    ' Loop to receive all the data sent by the client
                    blen = stream.Read(bytes, 0, bytes.Length)
                    While (blen <> 0)

                        ' Translate data bytes to a ASCII string
                        xmlMessage = xmlMessage + System.Text.Encoding.ASCII.GetString(bytes, 0, blen)

                        If xmlMessage.ToLower.Contains("</oncontactlookup>") Then

                            Dim dt As New DataTable
                            dt.Columns.Add("Call Type", GetType(String))
                            dt.Columns.Add("Member ID", GetType(String))
                            dt.Columns.Add("Claim Number", GetType(String))
                            dt.Columns.Add("Date of Service", GetType(DateTime))
                            dt.Columns.Add("Authorization Number", GetType(String))
                            dt.Columns.Add("Workgroup", GetType(String))
                            dt.Columns.Add("Call ID", GetType(String))
                            dt.Columns.Add("User ID", GetType(String))
                            dt.Columns.Add("Call XML", GetType(String))
                            Dim xDoc As New XmlDocument()
                            xDoc.LoadXml(xmlMessage)
                            Dim n As XmlElement = xDoc.DocumentElement
                            dt.Rows.Add(fmt.GetAttribute(n, "LookupType"), _
                                        n.SelectSingleNode("MemberID").InnerText.Trim, _
                                        n.SelectSingleNode("ClaimNumber").InnerText.Trim, _
                                        IIf(IsDate(n.SelectSingleNode("DateOfService").InnerText), n.SelectSingleNode("DateOfService").InnerText, Nothing), _
                                        n.SelectSingleNode("AuthorizationNumber").InnerText.Trim, _
                                        n.SelectSingleNode("Workgroup").InnerText.Trim, _
                                        n.SelectSingleNode("CallIDKey").InnerText.Trim, _
                                        n.SelectSingleNode("UserID").InnerText.Trim, _
                                        xDoc.OuterXml)
                            RaiseEvent InteractiveReceived(dt)
                            blen = 0
                            Exit While
                        Else
                            ' Get next stream
                            blen = stream.Read(bytes, 0, bytes.Length)

                        End If
                    End While

                    ' Shutdown and end connection
                    client.Close()

                End While
            Catch ex As Exception
                ' ReportError(ex)
                ' No error if it can't start up
                ' Note that this will happen if user opens a second instance of application
            Finally
                server.Stop()
            End Try
        End Sub
    End Class

#End Region

#Region "Active Directory"

    '''<summary>Get a property for a particular user</summary> 
    ''' <param name="PropertyName">Property name to return</param>
    ''' <param name="UserName">Optional username, defaults to current.</param>
    ''' <param name="path">Optional LDAP path. Defaults to current.</param>
    '''<remarks>Returns only the first value of PropertyName</remarks>
    ''' <returns>The first property value</returns>
    Public Function GetUserProperty(ByVal PropertyName As String, Optional ByVal UserName As String = "", Optional ByVal path As String = "") As String
        Dim root As DirectoryEntry = Nothing
        Dim ds As DirectorySearcher = Nothing
        Try
            ' Get directory
            If path = "" Then
                root = Domain.GetCurrentDomain.GetDirectoryEntry
            Else
                root = New DirectoryEntry(path)
            End If

            ' Get current User if none provided
            If UserName = "" Then UserName = nwk.GetUser()

            ' Create search filter
            ds = New DirectorySearcher(root)
            ds.Filter = "(&(sAMAccountType=805306368)(samaccountname=" + UserName + "))"

            ' Retrieve and return value
            Dim result As SearchResult = ds.FindOne
            If result Is Nothing Then Return ""
            If result.Properties(PropertyName).Count > 0 Then
                Return result.Properties(PropertyName)(0).ToString
            End If
            Return ""

        Finally
            ReleaseObj(ds)
            ReleaseObj(root)
        End Try
    End Function

#End Region

#Region "Printing"

    ''' <summary>Checks if a named printer is installed.</summary>
    ''' <param name="PrinterName">The name of the printer to check.</param>
    ''' <return>True if printer is installed.</return>
    Public Function IsPrinterInstalled(ByVal PrinterName As String) As Boolean
        For Each s As String In Printing.PrinterSettings.InstalledPrinters
            If s.Trim.ToUpper = PrinterName.Trim.ToUpper Then Return True
        Next
        Return False
    End Function

    ''' <summary>Returns a datatable of installed printers.</summary>
    ''' <return>Datatable of printer names.</return>
    Public Function GetPrinters() As DataTable
        GetPrinters = New DataTable("Installed Printers")
        GetPrinters.Columns.Add("Printer Name")
        For Each s As String In Printing.PrinterSettings.InstalledPrinters
            GetPrinters.Rows.Add(s)
        Next
    End Function

    ''' <summary>Sends a text document to the printer.</summary>
    ''' <param name="Text">The text document to print.</param>
    ''' <param name="PrinterName">The name of the printer. Default is default printer.</param>
    ''' <param name="Landscape">Set to True to print landscape.</param>
    Public Sub PrintText(ByVal Text As String, Optional ByVal PrinterName As String = "", Optional ByVal Landscape As Boolean = False)
        Dim prn As New TextPrint(Text)
        Try
            ' Check if the specified printer is installed
            If PrinterName <> "" Then
                If Not IsPrinterInstalled(PrinterName) Then
                    Throw (New Exception("Printer " + PrinterName + " not found."))
                    Exit Sub
                End If
                prn.PrinterSettings.PrinterName = PrinterName
            End If

            ' Print it
            prn.DefaultPageSettings.Landscape = Landscape
            prn.Print()
        Finally
            ReleaseObj(prn)
        End Try
    End Sub

    ''' <summary>Sends a file to a printer.</summary>
    ''' <param name="FileName">The file name to print.</param>
    ''' <param name="PrinterName">Optional name of the printer to send the document to. If no PrinterName is provided it is sent to the default printer.</param>
    '''<remarks>
    ''' To fax, you can simply print to \\sv-hfi-rfax\HPFAX (assuming that is your networked fax printer)
    ''' In your fax, you can embed fax commands for automatic faxing (substitute angle brackets for [])
    ''' [COVER][ToName:Tyson Gill][ToFaxNum:912128095066][Note:Transportation Request]
    ''' Note that you must put this embedded XML in the System font in order for RightFax to recognize it.
    ''' </remarks>
    Public Sub PrintFile(ByVal FileName As String, Optional ByVal PrinterName As String = "")
        Dim p As New Process
        p.StartInfo.FileName = FileName
        If PrinterName.Trim = "" Then
            p.StartInfo.Verb = "Print"
        Else
            If Not IsPrinterInstalled(PrinterName) Then
                Throw (New Exception("Printer " + PrinterName + " not found."))
                Exit Sub
            End If
            p.StartInfo.Verb = "PrintTo"
            p.StartInfo.Arguments = """" + PrinterName + """"
        End If
        p.StartInfo.CreateNoWindow = True
        p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
        p.StartInfo.UseShellExecute = True
        p.Start()
    End Sub

    Private Class TextPrint

        ' Inherits all the functionality of a PrintDocument
        Inherits Printing.PrintDocument

        ' Private variables to hold default font and text
        Private fntPrintFont As Font
        Private strText As String

        Friend Sub New(ByVal Text As String)
            ' Sets the file stream
            MyBase.New()
            strText = Text
        End Sub

        Friend Property Text() As String
            Get
                Return strText
            End Get
            Set(ByVal Value As String)
                strText = Value
            End Set
        End Property

        Protected Overrides Sub OnBeginPrint(ByVal ev _
                                    As Printing.PrintEventArgs)
            ' Run base code
            MyBase.OnBeginPrint(ev)
            ' Sets the default font
            If fntPrintFont Is Nothing Then
                fntPrintFont = New Font("Times New Roman", 12)
            End If
        End Sub

        Friend Property Font() As Font
            ' Allows the user to override the default font
            Get
                Return fntPrintFont
            End Get
            Set(ByVal Value As Font)
                fntPrintFont = Value
            End Set
        End Property

        Protected Overrides Sub OnPrintPage(ByVal ev _
           As Printing.PrintPageEventArgs)
            ' Provides the print logic for our document

            ' Run base code
            MyBase.OnPrintPage(ev)
            ' Variables
            Static intCurrentChar As Integer
            Dim intPrintAreaHeight, intPrintAreaWidth, _
                intMarginLeft, intMarginTop As Integer
            ' Set printing area boundaries and margin coordinates
            With MyBase.DefaultPageSettings
                intPrintAreaHeight = .PaperSize.Height - _
                                   .Margins.Top - .Margins.Bottom
                intPrintAreaWidth = .PaperSize.Width - _
                                  .Margins.Left - .Margins.Right
                intMarginLeft = .Margins.Left 'X
                intMarginTop = .Margins.Top   'Y
            End With
            ' If Landscape set, swap printing height/width
            If MyBase.DefaultPageSettings.Landscape Then
                Dim intTemp As Integer
                intTemp = intPrintAreaHeight
                intPrintAreaHeight = intPrintAreaWidth
                intPrintAreaWidth = intTemp
            End If
            ' Calculate total number of lines
            Dim intLineCount As Int32 = _
                    CInt(intPrintAreaHeight / Font.Height)
            ' Initialise rectangle printing area
            Dim rectPrintingArea As New RectangleF(intMarginLeft, _
                intMarginTop, intPrintAreaWidth, intPrintAreaHeight)
            ' Initialise StringFormat class, for text layout
            Dim objSF As New StringFormat(StringFormatFlags.LineLimit)
            ' Figure out how many lines will fit into rectangle
            Dim intLinesFilled, intCharsFitted As Int32
            ev.Graphics.MeasureString(Mid(strText, _
                        UpgradeZeros(intCurrentChar)), Font, _
                        New SizeF(intPrintAreaWidth, _
                        intPrintAreaHeight), objSF, _
                        intCharsFitted, intLinesFilled)
            ' Print the text to the page
            ev.Graphics.DrawString(Mid(strText, _
                UpgradeZeros(intCurrentChar)), Font, _
                Brushes.Black, rectPrintingArea, objSF)
            ' Increase current char count
            intCurrentChar += intCharsFitted
            ' Check whether we need to print more
            If intCurrentChar < strText.Length Then
                ev.HasMorePages = True
            Else
                ev.HasMorePages = False
                intCurrentChar = 0
            End If
        End Sub

        ' Upgrades all zeros to ones
        ' used instead of defunct IIF or messy If statements
        Private Function UpgradeZeros(ByVal Input As Integer) As Integer
            If Input = 0 Then
                Return 1
            Else
                Return Input
            End If
        End Function
    End Class

#End Region

#Region "FTP"

    ''' <summary>Retrieves an ftp file.</summary>
    ''' <param name="FTPAddress">The full ftp address of the file to retrieve.</param>
    ''' <param name="UserName">The username, or empty if anonymous.</param>
    ''' <param name="Password">The password, or empty if none is needed.</param>
    ''' <param name="LocalFile">The local file to create, or empty if none is needed.</param>
    ''' <returns>The retrieved file.</returns>
    Public Function FTPGet(ByVal FTPAddress As String, Optional ByVal UserName As String = "", Optional ByVal Password As String = "", Optional ByVal LocalFile As String = "") As Object
        Dim reader As StreamReader = Nothing
        Dim tw As StreamWriter = Nothing
        Dim responseStream As Stream = Nothing
        Try
            ' Create the ftp request
            Dim request As FtpWebRequest = WebRequest.Create(FTPAddress)
            request.Method = WebRequestMethods.Ftp.DownloadFile
            request.Credentials = New NetworkCredential(UserName, Password)

            ' Get the file
            Dim response As FtpWebResponse = request.GetResponse()
            responseStream = response.GetResponseStream()
            reader = New StreamReader(responseStream)
            FTPGet = reader.ReadToEnd()
            reader.Close()
            response.Close()

            ' Write to a local file
            If LocalFile = "" Then Exit Function
            tw = New StreamWriter(LocalFile, False)
            tw.Write(FTPGet)
            tw.Close()
        Finally
            ReleaseObj(reader)
            ReleaseObj(responseStream)
            ReleaseObj(tw)
        End Try
    End Function

    ''' <summary>Sends an ftp file.</summary>
    ''' <param name="FTPAddress">The full ftp address of the file to create.</param>
    ''' <param name="LocalFile">The local file to send.</param>
    ''' <param name="UserName">The username, or empty if anonymous.</param>
    ''' <param name="Password">The password, or empty if none is needed.</param>
    ''' <returns>The retrieved file.</returns>
    Public Function FTPPut(ByVal LocalFile As String, ByVal FTPAddress As String, Optional ByVal UserName As String = "", Optional ByVal Password As String = "") As Boolean
        Dim sourceStream As StreamReader = Nothing
        Dim requestStream As Stream = Nothing
        Try
            ' Create the ftp request
            Dim request As FtpWebRequest = WebRequest.Create(FTPAddress)
            request.Method = WebRequestMethods.Ftp.UploadFile
            request.Credentials = New NetworkCredential(UserName, Password)

            ' Get the file to send
            Dim fileContents As Byte() = File.ReadAllBytes(LocalFile)

            ' Send the contents
            request.ContentLength = fileContents.Length
            requestStream = request.GetRequestStream()
            requestStream.Write(fileContents, 0, fileContents.Length)
            requestStream.Close()
            Dim response As FtpWebResponse = request.GetResponse()
            response.Close()
        Finally
            ReleaseObj(sourceStream)
            ReleaseObj(requestStream)
        End Try
    End Function

#End Region

#Region "Registry"

    ''' <summary>Retrieves a registry value.</summary>
    ''' <param name="Hive">The registry hive to search.</param>
    ''' <param name="Key">The registry key to search.</param>
    ''' <param name="ValueName">The registry value to retrieve.</param>
    ''' <returns>The retrieved value.</returns>
    Public Function GetRegValue(ByVal Hive As RegistryHive, ByVal Key As String, ByVal ValueName As String) As String

        ' Get the Registry Hive
        Dim objParent As RegistryKey = Nothing
        Dim objSubkey As RegistryKey = Nothing
        Dim sValue As String = ""
        Select Case Hive
            Case RegistryHive.ClassesRoot
                objParent = Registry.ClassesRoot
            Case RegistryHive.CurrentConfig
                objParent = Registry.CurrentConfig
            Case RegistryHive.CurrentUser
                objParent = Registry.CurrentUser
            Case RegistryHive.PerformanceData
                objParent = Registry.PerformanceData
            Case RegistryHive.LocalMachine
                objParent = Registry.LocalMachine
            Case RegistryHive.PerformanceData
                objParent = Registry.PerformanceData
            Case RegistryHive.Users
                objParent = Registry.Users
        End Select

        ' Get the value
        objSubkey = objParent.OpenSubKey(Key)
        If Not objSubkey Is Nothing Then sValue = (objSubkey.GetValue(ValueName))
        Return sValue

    End Function

#End Region

#Region "TCP/IP"

    ''' <summary>Submits a packet to a host server and returns a the response.</summary>
    ''' <param name="IP">The IP Address to connect to.</param>
    ''' <param name="Port">The IP Port to connect to.</param>
    ''' <param name="Request">The packet to send.</param>
    Public Function TCPSend(ByVal IP As String, ByVal Port As String, ByVal Request As String) As String
        Dim tcpClient As New System.Net.Sockets.TcpClient()
        Try
            ' Open a new stream
            tcpClient.Connect(IP, Val(Port))
            Dim networkStream As NetworkStream = tcpClient.GetStream()

            ' Write the request
            Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(Request)
            networkStream.Write(sendBytes, 0, sendBytes.Length)

            ' Read the response
            Dim bytes(tcpClient.ReceiveBufferSize) As Byte
            networkStream.ReadTimeout = 30000
            networkStream.Read(bytes, 0, CInt(tcpClient.ReceiveBufferSize))
            Return Encoding.ASCII.GetString(bytes)

        Finally
            tcpClient.Close()
        End Try
    End Function

#End Region

#Region "Application Icons"

    Private Structure SHFILEINFO
        Friend hIcon As IntPtr
        Friend iIcon As Integer
        Friend dwAttributes As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)> Public szDisplayName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)> Public szTypeName As String
    End Structure

    Private Const SHGFI_ICON = &H100
    Private Const SHGFI_SMALLICON = &H1
    Private Const SHGFI_LARGEICON = &H0
    Private Declare Auto Function SHGetFileInfo Lib "shell32.dll" _
             (ByVal pszPath As String, _
              ByVal dwFileAttributes As Integer, _
              ByRef psfi As SHFILEINFO, _
              ByVal cbFileInfo As Integer, _
              ByVal uFlags As Integer) As IntPtr

    ''' <summary>Return the icon associated with a file.</summary>
    ''' <param name="FileName">The file name to extract.</param>
    ''' <returns>The icon associated wtih the file.</returns>
    ''' <remarks>Works with file links.</remarks>
    Public Function IconFromFile(ByVal FileName As String, Optional ByVal LargeIcon As Boolean = True) As Icon

        ' Initialize the structure
        Dim shinfo As SHFILEINFO = New SHFILEINFO()
        shinfo.szDisplayName = New String(Chr(0), 260)
        shinfo.szTypeName = New String(Chr(0), 80)

        ' Get the icon
        If LargeIcon Then
            SHGetFileInfo(FileName, 0, shinfo, Marshal.SizeOf(shinfo), SHGFI_ICON Or SHGFI_LARGEICON)
        Else
            SHGetFileInfo(FileName, 0, shinfo, Marshal.SizeOf(shinfo), SHGFI_ICON Or SHGFI_SMALLICON)
        End If

        ' Return the icon
        Return System.Drawing.Icon.FromHandle(shinfo.hIcon)

    End Function

    Friend Declare Auto Function ExtractIcon Lib "shell32" (ByVal hInstance As IntPtr, ByVal lpszExeFileName As String, ByVal nIconIndex As Integer) As IntPtr
    ''' <summary>Return the icon associated with a file.</summary>
    ''' <param name="FileName">The file name to extract.</param>
    ''' <returns>The image associated wtih the file.</returns>
    ''' <remarks>IconFromFile is more general and works with file links.</remarks>
    Private Function IconExtract(ByVal FileName As String) As Image
        Dim hInstance As IntPtr = Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly.GetModules()(0))
        Dim hIcon As IntPtr = ExtractIcon(hInstance, FileName, 0)
        If hIcon.Equals(IntPtr.Zero) Then Return Nothing
        IconExtract = Bitmap.FromHicon(hIcon)
    End Function

#End Region

#Region "Drive Mapping"

    ''' <summary>
    ''' Maps a network drive and assigns permission to a user account
    ''' </summary>
    Private Class MapDrive

        Structure NetResource
            Friend dwScope As Int32
            Friend dwType As Int32
            Friend dwDisplayType As Int32
            Friend dwUsage As Int32
            Friend lpLocalName As String
            Friend lpRemoteName As String
            Friend lpComment As String
            Friend lpProvider As String
        End Structure

        Declare Function WNetAddConnection2 Lib "mpr.dll" Alias _
          "WNetAddConnection2A" (ByRef lpNetResource As NetResource, _
          <MarshalAs(UnmanagedType.LPStr)> ByVal lpPassword As String, <MarshalAs(UnmanagedType.LPStr)> ByVal lpUserName As String, _
          ByVal dwFlags As Int32) As Int32

        Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias _
              "WNetCancelConnection2A" (ByVal lpName As String, _
              ByVal dwFlags As Long, ByVal fForce As Long) As Long

        ''' <summary>Set to true if the drive has been sucessfully mapped</summary>
        Public DriveMapped As Boolean = False

        Private rsc As New NetResource

        ''' <summary>
        ''' Creates a new MapDrive object and creates the system drive mapping
        ''' </summary>
        ''' <param name="NetPath">The network path to map.</param>
        ''' <param name="UserName">The username to access the path.</param>
        ''' <param name="Password">The password to access the path.</param>
        ''' <param name="DriveName">The optional local drive to create, e.g. "z:"</param>
        ''' <remarks>This creates a permanent system drive mapping unless the Remove method is called.</remarks>
        Public Sub New(ByVal NetPath As String, ByVal UserName As String, ByVal Password As String, Optional ByVal DriveName As String = "")
            rsc.dwType = &H1 ' Set resource type to disk
            rsc.lpRemoteName = NetPath
            rsc.lpLocalName = DriveName
            DriveMapped = (WNetAddConnection2(rsc, Password, UserName, 0) = 0)
        End Sub

        ''' <summary>Removes a mapped drive.</summary>
        ''' <remarks>If not called before the object is destroyed the drive will remain mapped.</remarks>
        Public Sub Remove()
            WNetCancelConnection2(rsc.lpRemoteName, 0, 1)
            DriveMapped = False
        End Sub

    End Class

#End Region

#Region "Impersonation"

    ''' <summary>A class to allow the application to impersonate another user.</summary>
    ''' <remarks>The local policy has to allow users to 'Act As Part of the Operating System.</remarks>
    Public Class Impersonator

        Private tokenHandle As New IntPtr(0)
        Private ImpersonatedUser As WindowsImpersonationContext

        Private Declare Auto Function LogonUser Lib "advapi32.dll" (ByVal lpszUsername As [String], _
        ByVal lpszDomain As [String], ByVal lpszPassword As [String], _
        ByVal dwLogonType As Integer, ByVal dwLogonProvider As Integer, _
        ByRef phToken As IntPtr) As Boolean

        Friend Declare Auto Function CloseHandle Lib "kernel32.dll" (ByVal handle As IntPtr) As Boolean

        ''' <summary>
        ''' Impersonate another windows account
        ''' </summary>
        ''' <param name="Domain">The domain on which the user account resides</param>
        ''' <param name="userName">The name of the user account</param>
        ''' <param name="Password">The password for the user account</param>
        ''' <returns>An errorstring, empty if there is no error.</returns>
        ''' <remarks>Impersonation will remain active until ImpersonateStop is called or the application ends.</remarks>
        <PermissionSetAttribute(SecurityAction.Demand, Name:="FullTrust")> _
        Public Function ImpersonateStart(ByVal Domain As String, ByVal userName As String, ByVal Password As String) As String

            ' Call LogonUser to obtain a handle to an access token.
            tokenHandle = IntPtr.Zero
            If Not LogonUser(userName, Domain, Password, 2, 0, tokenHandle) Then
                Return New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error()).Message
            End If

            ' Use the token handle returned by LogonUser
            Dim newID As New WindowsIdentity(tokenHandle)
            ImpersonatedUser = newID.Impersonate()
            Return ""

        End Function

        ''' <summary>
        ''' End the current impersonation and return to the default user
        ''' </summary>
        <PermissionSetAttribute(SecurityAction.Demand, Name:="FullTrust")> _
        Public Sub ImpersonateStop()

            If ImpersonatedUser Is Nothing Then Exit Sub
            ImpersonatedUser.Undo()
            If Not System.IntPtr.op_Equality(tokenHandle, IntPtr.Zero) Then CloseHandle(tokenHandle)

        End Sub

        ''' <summary>
        ''' Get the currently impersonated user.
        ''' </summary>
        Public Function GetCurrentUser() As String
            Return WindowsIdentity.GetCurrent.Name()
        End Function

    End Class

#End Region

#Region "Dynamic Code Execution"

    ''' <summary>
    ''' Dynamically compile and execute a function
    ''' </summary>
    ''' <param name="VBCode">The code to be executed, including the Return statement</param>
    ''' <returns>An object with the function results, or Nothing if there is an error</returns>
    ''' <remarks>Sample VBCode: Return (12+5)*9</remarks>
    Public Function DynamicFunction(ByVal VBCode As String) As Object
        Try
            ' Initialize the compiler
            Dim objCodeCompiler As VBCodeProvider = New VBCodeProvider
            Dim objCompilerParameters As New System.CodeDom.Compiler.CompilerParameters
            objCompilerParameters.GenerateInMemory = True

            ' Create the complete function code to be compiled
            Dim Code As New System.Text.StringBuilder
            Code.Append("Namespace DyNamespace").Append(vbCrLf).Append("Public Class DyClass").Append(vbCrLf).Append("Public Function DyFunction() as Object").Append(vbCrLf)
            Code.Append(VBCode).Append(vbCrLf)
            Code.Append("End Function").Append(vbCrLf).Append("End Class").Append(vbCrLf).Append("End Namespace")

            ' Compile the code
            Dim Compilation As System.CodeDom.Compiler.CompilerResults = objCodeCompiler.CompileAssemblyFromSource(objCompilerParameters, Code.ToString)
            If Compilation.Errors.HasErrors Then Return Nothing

            ' Get a reference to the dynamic assembly
            Dim DyAssem As System.Reflection.Assembly = Compilation.CompiledAssembly

            ' Create an instance of the dynamic class to be invoked
            Dim DyClass As Object = DyAssem.CreateInstance("DyNamespace.DyClass")
            If DyClass Is Nothing Then Return Nothing

            ' Call the dynamic function
            Dim FunctionValue As Object = DyClass.GetType.InvokeMember("DyFunction", System.Reflection.BindingFlags.InvokeMethod, Nothing, DyClass, Nothing)

            ' Return the results
            Return FunctionValue

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

#End Region

#Region "Inter-process Messaging"

    Declare Auto Function SendMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr

    Private Sub Send_Message(ByVal sender As System.Object, ByVal e As System.EventArgs)

        ' Place into the clipboard
        Clipboard.SetText("SF23841V")
        If Not SendToApplication("OnContact", 99) Then
            MessageBox.Show("You need to start OnContact.", "Open OnContact", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

    ' Send a message to an instance of a running application
    Private Function SendToApplication(ByVal AppName As String, ByVal Msg As Integer) As Boolean
        Try
            ' Get the running process
            Dim ocp() As Process = Process.GetProcessesByName(AppName)
            If ocp.Length = 0 Then Return False

            ' Send a message
            SendMessage(ocp(0).MainWindowHandle, Msg, IntPtr.Zero, IntPtr.Zero)
            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function

    ' Listener
    'Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
    '    If m.Msg = 99 Then
    '        MessageBox.Show("Received Message")
    '    End If
    '    MyBase.WndProc(m)
    'End Sub

#End Region

End Class


