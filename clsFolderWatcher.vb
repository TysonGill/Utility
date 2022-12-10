Imports System.IO
Imports System.Diagnostics

''' <summary>Watches a folder and raises an event when a file chanegs.</summary>
Public Class clsFolderWatcher

    ''' <summary>Raised when a file in the watch folder has been modified.</summary>
    ''' <param name="FileName">Name of the file that has changed.</param>
    ''' <param name="Change">The change type, or the new file name.</param>
    ''' <remarks>Use WithEvents when you declare clsFolderWatcher and handle this event.</remarks>
    ''' <example>
    ''' Private WithEvents clsFolderWatcher As New clsFolderWatcher("c:\")
    ''' </example>
    Public Event FileChanged(ByVal FileName As String, ByVal Change As String)

    Private LastChange As DateTime = Now
    Private LastFile As String = ""
    Private fw As FileSystemWatcher

    ''' <summary>Instantiate the class instance.</summary>
    Public Sub New(ByVal Folder As String, Optional ByVal Enabled As Boolean = True)
        Folder = fmt.AppendIfNeeded(Folder, "\")
        If Not Directory.Exists(Folder) Then
            Throw New Exception("Folder Watcher Error: The watch folder " + Folder + " does not exist.")
            Exit Sub
        End If
        fw = New System.IO.FileSystemWatcher(Folder)
        fw.NotifyFilter = NotifyFilters.FileName Or NotifyFilters.LastWrite
        AddHandler fw.Changed, AddressOf WatchFileChanged
        AddHandler fw.Created, AddressOf WatchFileChanged
        AddHandler fw.Deleted, AddressOf WatchFileChanged
        AddHandler fw.Renamed, AddressOf WatchFileRenamed
        fw.EnableRaisingEvents = Enabled
    End Sub

    ''' <summary>The folder to monitor.</summary>
    Public ReadOnly Property Folder() As String
        Get
            Folder = fw.Path
        End Get
    End Property

    ''' <summary>True to raise an event when a file changes.</summary>
    Public Property Enabled() As Boolean
        Get
            Enabled = fw.EnableRaisingEvents
        End Get
        Set(ByVal value As Boolean)
            fw.EnableRaisingEvents = value
        End Set
    End Property


    ''' <summary>Returns a listing of files in the folder.</summary>
    ''' <returns>A DataTable with the File Name and Updated Date.</returns>
    Public Function GetFileTable() As DataTable

        ' Create the tables
        GetFileTable = New DataTable
        GetFileTable.Columns.Add("Full Path", GetType(System.String))
        GetFileTable.Columns.Add("File Name", GetType(System.String))
        GetFileTable.Columns.Add("Updated", GetType(System.DateTime))

        ' Get the file records
        Dim f As String
        For Each f In Directory.GetFiles(fw.Path)
            Dim row As DataRow = GetFileTable.NewRow
            row("Full Path") = f
            row("File Name") = Path.GetFileName(f)
            row("Updated") = File.GetLastWriteTime(f)
            GetFileTable.Rows.Add(row)
        Next

    End Function

    Private Sub WatchFileChanged(ByVal source As Object, ByVal e As System.IO.FileSystemEventArgs)
        RaiseEvent FileChanged(e.Name, e.ChangeType.ToString)
    End Sub

    Private Sub WatchFileRenamed(ByVal source As Object, ByVal e As System.IO.RenamedEventArgs)
        RaiseEvent FileChanged(e.OldName, e.Name)
    End Sub

End Class
