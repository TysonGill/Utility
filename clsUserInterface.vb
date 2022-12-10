Imports System.Drawing
Imports System.Windows.Forms

''' <summary>General library for User Interface routines.</summary>
Public Class clsUserInterface

    ''' <summary>Loops through all child controls to set their editability.</summary>
    Public Sub MakeControlsEditable(ByVal grp As Control, Optional ByVal Editable As Boolean = True)
        For Each ctrl As Control In grp.Controls
            If TypeOf ctrl Is TextBox Then
                CType(ctrl, TextBox).ReadOnly = Not Editable
            ElseIf TypeOf ctrl Is ComboBox Then
                CType(ctrl, ComboBox).Enabled = Editable
            ElseIf TypeOf ctrl Is DateTimePicker Then
                CType(ctrl, DateTimePicker).Enabled = Editable
            ElseIf TypeOf ctrl Is LinkLabel Then
                CType(ctrl, LinkLabel).Enabled = Editable
            ElseIf TypeOf ctrl Is CheckBox Then
                CType(ctrl, CheckBox).Enabled = Editable
            ElseIf TypeOf ctrl Is ListBox Then
                CType(ctrl, ListBox).Enabled = Editable
            Else
                If ctrl.Controls.Count > 0 Then
                    MakeControlsEditable(ctrl, Editable)
                End If
            End If
        Next
    End Sub

    ''' <summary>Loops through all child controls and returns a string of current values.</summary>
    Public Function GetControlString(ByVal grp As Control) As String
        Dim ValStr As String = ""
        For Each ctrl As Control In grp.Controls
            If TypeOf ctrl Is TextBox Then
                ValStr += CType(ctrl, TextBox).Text
            ElseIf TypeOf ctrl Is ComboBox Then
                ValStr += CType(ctrl, ComboBox).Text
            ElseIf TypeOf ctrl Is DateTimePicker Then
                ValStr += CType(ctrl, DateTimePicker).Value.ToUniversalTime
            ElseIf TypeOf ctrl Is CheckBox Then
                ValStr += CType(ctrl, CheckBox).Checked.ToString
            ElseIf TypeOf ctrl Is ListBox Then
                ValStr += CType(ctrl, ListBox).SelectedValue.ToString
            Else
                If ctrl.Controls.Count > 0 Then
                    ValStr += GetControlString(ctrl)
                End If
            End If
        Next
        Return ValStr
    End Function

    ''' <summary>Displays a general dialog for selecting a value from a list.</summary>
    ''' <remarks>Returns the first field in the datatable provided.</remarks>
    Public Function SelectFromList(ByVal dt As DataTable, ByVal ListField As String, Optional ByVal Title As String = "", Optional ByVal Instructions As String = "", Optional ByVal ReturnedField As String = "") As String
        Dim frm As New frmListSelector
        Try
            frm.ListData = dt
            frm.ListField = ListField
            If ReturnedField = "" Then ReturnedField = ListField
            frm.ReturnedField = ReturnedField
            If Title <> "" Then frm.Text = Title
            If Instructions <> "" Then frm.lblInstructions.Text = Instructions
            frm.ShowDialog()
            SelectFromList = frm.Selection
            frm.Close()
        Finally
            frm = Nothing
        End Try
    End Function

    ''' <summary>Displays a general login form for SQL Server and MHS.</summary>
    ''' <remarks>Updates the connection string in Utility.</remarks>
    Public Function ShowAbout() As System.Windows.Forms.DialogResult
        Dim f As New frmAbout
        f.ShowDialog()
    End Function

    ''' <summary>Add an item to the recent menu list.</summary>
    ''' <param name="mnuParent">The parent menu item.</param>
    ''' <param name="Display">The menu text to dispaly.</param>
    ''' <param name="Tag">Key for comparison of menu items.</param>
    ''' <remarks>To configure, add the number of child items you want as place holders and make them invisible initially.</remarks>
    Public Sub MenuAddRecent(ByVal mnuParent As ToolStripMenuItem, ByVal Display As String, ByVal Tag As String)
        Try
            ' If item is already in list then replace it
            Dim i As Integer
            Dim ExistingItem As Integer
            For ExistingItem = 0 To mnuParent.DropDownItems.Count - 1
                If mnuParent.DropDownItems(ExistingItem).Tag = Tag Then
                    Exit For
                End If
            Next
            If ExistingItem = mnuParent.DropDownItems.Count Then ExistingItem -= 1

            ' Move existing items down in the list
            For i = ExistingItem To 1 Step -1
                If mnuParent.DropDownItems(i - 1).Tag <> "" Then
                    mnuParent.DropDownItems(i).Text = mnuParent.DropDownItems(i - 1).Text
                    mnuParent.DropDownItems(i).Tag = mnuParent.DropDownItems(i - 1).Tag
                    mnuParent.DropDownItems(i).Visible = True
                End If
            Next

            ' Add the new item to the top of the list
            mnuParent.DropDownItems(0).Text = Display
            mnuParent.DropDownItems(0).Tag = Tag
            mnuParent.DropDownItems(0).Visible = True
        Catch ex As Exception
            Utility.ReportError(ex, False)
        End Try
    End Sub

    ''' <summary>Find a value in a list or combo box.</summary>
    ''' <param name="lst">The list or combo box to search.</param>
    ''' <param name="strText">The text to find.</param>
    ''' <remarks>More friendly than the built in .Net method.</remarks>
    ''' <returns>The index of the found item or -1.</returns>
    Public Function ListSearch(ByVal lst As Object, ByVal strText As String, Optional ByVal SearchOnKey As Boolean = False) As Integer
        Try
            Cursor.Current = Cursors.WaitCursor

            ' Search for the item
            Dim i As Integer
            For i = 0 To lst.Items.Count - 1

                If TypeOf lst.Items(i) Is DictionaryEntry Then
                    If SearchOnKey = True Then
                        If Trim(CType(lst.Items(i), DictionaryEntry).Key) = Trim(strText) Then
                            Return i
                        End If
                    Else
                        If Trim(CType(lst.Items(i), DictionaryEntry).Value) = Trim(strText) Then
                            Return i
                        End If
                    End If
                Else
                    If Trim(lst.Items(i)) = Trim(strText) Then
                        Return i
                    End If
                End If
            Next
            Return -1

        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Function

    ''' <summary>Execute or open a file.</summary>
    ''' <param name="FileName">Full path and file name for file to open.</param>
    ''' <remarks>Files will open in the default Windows application associated with them.</remarks>
    Public Sub OpenFile(ByVal FileName As String)
        Try
            Process.Start(FileName)
        Catch ex As Exception
            MessageBox.Show("Could not display file:" + vbCrLf + FileName + vbCrLf + vbCrLf + "Make sure that you have a default viewer associated with this file type.", "File Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try
    End Sub

    ''' <summary>Show an HTML string in the default browser.</summary>
    ''' <param name="PageName">The file name to create and pass to browser.</param>
    ''' <param name="sHTML">The HTML string (can be plain text also).</param>
    ''' <remarks>Will write the file in the current directory.</remarks>
    Sub ShowInBrowser(ByVal PageName As String, ByVal sHTML As String)
        Try
            Cursor.Current = Cursors.WaitCursor

            Dim sHTMLFile As String
            sHTMLFile = My.Computer.FileSystem.CurrentDirectory
            sHTMLFile += "\" + PageName + ".html"

            My.Computer.FileSystem.WriteAllText(sHTMLFile, sHTML, False)
            ui.OpenFile(sHTMLFile)
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub


    ''' <summary>
    ''' Prompts the user for a date or optional date range.
    ''' </summary>
    ''' <param name="DateFrom">The starting date, pass Nothing for no starting date.</param>
    ''' <param name="DateTo">The ending date, pass Nothing for no ending date</param>
    ''' <param name="Title">Optional title of the date entry form.</param>
    ''' <returns>True if user has enter a valid date or date range.</returns>
    ''' <remarks>Start and end date are returned in the parameters passed by reference.</remarks>
    Public Function GetDateRange(ByRef DateFrom As DateTime, ByRef DateTo As DateTime, Optional ByVal Title As String = "Date or Date Range Entry") As Boolean

        ' Format the default string
        Dim DefaultRange As String = ""
        If Not DateFrom = Nothing Then DefaultRange = DateFrom.ToShortDateString
        If Not DateTo = Nothing Then DefaultRange += " - " + DateTo.ToShortDateString

        ' Get date or date range from user
        Dim Range As String = InputBox("Enter a date or a date range in the format mm/dd/yyyy - mm/dd/yyyy", Title, DefaultRange).Trim
        If Range = "" Then Return False

        ' Parse the input string
        Dim rdates() As String = Split(Range, "-")
        If rdates.Length = 1 Then
            If Not IsDate(Range) Then Return False
            DateFrom = Date.Parse(Range)
            DateTo = Nothing
            Return True
        ElseIf rdates.Length = 2 Then
            If Not IsDate(rdates(0)) Then Return False
            If Not IsDate(rdates(1)) Then Return False
            DateFrom = Date.Parse(rdates(0))
            DateTo = Date.Parse(rdates(1))
            Return True
        Else
            Return False
        End If

    End Function

End Class
