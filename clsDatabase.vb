Imports System.Data
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports Utility

''' <summary>General library for database routines.</summary>
''' <remarks>Use clsDataQuery for your database access library.</remarks>
Public Class clsDatabase

    ''' <summary>Sorts a datatable.</summary>
    ''' <param name="dt">The datatable to sort.</param>
    ''' <param name="SortString">The sort string.</param>
    ''' <returns>A sorted datatable.</returns>
    ''' <remarks>Sort string should be in standard SQL format like "[Member ID], [Group Number] DESC"</remarks>
    Public Function SortDataTable(ByVal dt As DataTable, Optional ByVal SortString As String = "") As DataTable
        Dim rows() As DataRow = dt.Select("", SortString)
        SortDataTable = dt.Clone()
        For Each r As DataRow In rows
            SortDataTable.ImportRow(r)
        Next
        SortDataTable.AcceptChanges()
    End Function

    ''' <summary>Safely returns a datatable record count.</summary>
    ''' <param name="dt">The datatable to check.</param>
    ''' <returns>A numeric record count.</returns>
    ''' <remarks>Safely returns a record count even if the datatable is Nothing.</remarks>
    Public Function GetRecordCount(ByVal dt As DataTable) As Integer
        If dt Is Nothing Then Return 0
        If dt.Rows Is Nothing Then Return 0
        Return dt.Rows.Count
    End Function

    ''' <summary>Return a delimited string for all unique values in a column.</summary>
    ''' <param name="dt">The DataTable to search for disinct entries.</param>
    ''' <param name="ColName">The name of the column to search for distinct entries.</param>
    ''' <returns>A delimited string of column values.</returns>
    Public Function GetDistinctString(ByVal dt As DataTable, Optional ByVal ColName As String = "", Optional ByVal IncludeQuotes As Boolean = True) As String
        GetDistinctString = ""
        If dt Is Nothing OrElse dt.Columns.Count = 0 OrElse dt.Rows.Count = 0 Then Return ""
        If ColName = "" Then ColName = dt.Columns(0).ColumnName
        For Each s As String In GetDistinct(dt, ColName)
            If IncludeQuotes Then
                GetDistinctString += fmt.q(s) + ", "
            Else
                GetDistinctString += s + ", "
            End If
        Next
        If GetDistinctString.Length > 2 Then GetDistinctString = Microsoft.VisualBasic.Left(GetDistinctString, GetDistinctString.Length - 2)
    End Function

    ''' <summary>Return collection of distinct members of a datatable column.</summary>
    ''' <param name="dt">The DataTable to search for disinct entries.</param>
    ''' <param name="ColName">The name of the column to search for distinct entries.</param>
    ''' <returns>A collection of distinct column values.</returns>
    Public Function GetDistinct(ByVal dt As DataTable, ByVal ColName As String) As Collection
        GetDistinct = New Collection
        For Each row As DataRow In dt.Rows
            If Not GetDistinct.Contains(row(ColName).ToString.Trim) Then GetDistinct.Add(row(ColName), row(ColName).ToString.Trim)
        Next
    End Function

    ''' <summary>Search a datatable for a row and return a column value.</summary>
    ''' <param name="dt">The datatable to saerch.</param>
    ''' <param name="criteria">The search criteria, e.g. "ColumnName = 'Test'.</param>
    ''' <param name="ReturnedColumn">The column name from which to return data.</param>
    ''' <returns>Value of the field.</returns>
    ''' <remarks>Returns Nothing if more than one row is returned from search.</remarks>
    Public Function GetRowValue(ByVal dt As DataTable, ByVal criteria As String, ByVal ReturnedColumn As String) As Object
        Try
            ' Get datarow
            Dim foundRows As DataRow()
            foundRows = dt.Select(criteria)
            If foundRows.Length = 1 Then
                GetRowValue = foundRows(0).Item(ReturnedColumn)
            Else
                GetRowValue = Nothing
            End If
        Finally
        End Try
    End Function

#Region "Import/Export"

    ''' <summary>Writes a DataTable into a CSV file.</summary>
    ''' <param name="dt">The DataTable to export.</param>
    ''' <param name="FullPath">The fully specified path and name of the file to export.</param>
    ''' <param name="IncludeHeader">Set True to include a header.</param>
    ''' <param name="Append">Set to True if you wish to append an existing file.</param>
    Public Sub DataTableExport(ByVal dt As DataTable, ByVal FullPath As String, Optional ByVal IncludeHeader As Boolean = False, Optional ByVal Delimiter As String = ",", Optional ByVal IncludeQuotes As Boolean = True, Optional ByVal Append As Boolean = False)
        If Append And File.Exists(FullPath) Then IncludeHeader = False
        Dim tw As New StreamWriter(FullPath, Append)
        Try
            If dt Is Nothing OrElse dt.Columns.Count = 0 Then Exit Sub
            Dim s As String
            Dim c As DataColumn

            ' Write the header
            If IncludeHeader Then
                s = ""
                For Each c In dt.Columns
                    s += c.ColumnName.Trim + Delimiter
                Next
                tw.Write(Left(s, s.Length - Delimiter.Length) + vbCrLf)
            End If

            ' Write the data
            If dt.Rows.Count = 0 Then Exit Sub
            For Each r As DataRow In dt.Rows
                s = ""
                For Each c In dt.Columns
                    If IncludeQuotes Then
                        s += fmt.Quote(r(c).ToString.Trim) + Delimiter
                    Else
                        s += r(c).ToString.Trim + Delimiter
                    End If
                Next
                tw.Write(Left(s, s.Length - Delimiter.Length) + vbCrLf)
            Next
            tw.Close()
        Finally
            ReleaseObj(tw)
        End Try
    End Sub

    ''' <summary>Writes a DataTable into a text file using the OLEDB driver.</summary>
    ''' <param name="dt">The DataTable to export.</param>
    ''' <param name="FullPath">The fully specified path and name of the file to export.</param>
    ''' <param name="IncludeHeader">Set True if the file should include a header.</param>
    ''' <remarks>
    ''' This routine uses the OLEDB text reader to import any text file.
    ''' By default it will attempt to assign the column types for a comma delimited file.
    ''' If you have any other type of file, you can specify the file format in a Schema.ini file
    ''' in the same folder as the file to be imported. This will handle any type of delimiter or fixed-width files.
    ''' See http://msdn2.microsoft.com/en-us/library/ms709353.aspx for a description of the schema file format.
    ''' </remarks>
    Private Sub DataTableExportOLEDB(ByVal dt As DataTable, ByVal FullPath As String, Optional ByVal IncludeHeader As Boolean = False)
        Dim cn As OleDb.OleDbConnection = Nothing
        Try
            Dim filename As String = System.IO.Path.GetFileName(FullPath)
            Dim path As String = FullPath.Replace(filename, Nothing)

            Dim headers As String
            Dim cnstr As String
            If IncludeHeader Then headers = "Yes" Else headers = "No"
            cnstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Text;FMT=Delimited;HDR=" + headers + "';"

            cn = New OleDb.OleDbConnection(cnstr)
            Dim cmd As New OleDb.OleDbCommand("SELECT * INTO " + filename, cn)
            cn.Open()
            cmd.ExecuteNonQuery()
        Finally
            cn.Close()
        End Try
    End Sub

    ''' <summary>Reads a text file into a DataTable using the OLEDB driver.</summary>
    ''' <param name="FullPath">The fully specified path and name of the file to import.</param>
    ''' <param name="HasHeader">Set True if the file includes a header.</param>
    ''' <returns>A DataTable.</returns>
    ''' <remarks>
    ''' This routine uses the OLEDB text reader to import any text file.
    ''' By default it will attempt to assign the column types for a comma delimited file.
    ''' If you have any other type of file, you can specify the file format in a Schema.ini file
    ''' in the same folder as the file to be imported. This will handle any type of delimiter or fixed-width files.
    ''' See http://msdn2.microsoft.com/en-us/library/ms709353.aspx for a description of the schema file format.
    ''' </remarks>
    Public Function TextImportOLDEB(ByVal FullPath As String, Optional ByVal HasHeader As Boolean = False) As DataTable
        Dim cn As OleDb.OleDbConnection = Nothing
        Try
            Dim filename As String = System.IO.Path.GetFileName(FullPath)
            Dim path As String = FullPath.Replace(filename, Nothing)

            Dim headers As String
            Dim cnstr As String
            If HasHeader Then headers = "Yes" Else headers = "No"
            cnstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Text;FMT=Delimited;HDR=" + headers + "';"
            cn = New OleDb.OleDbConnection(cnstr)
            Dim cmd As New OleDb.OleDbCommand("SELECT * FROM [" + filename + "]", cn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet()
            cn.Open()
            da.Fill(ds, "CSV")
            Return ds.Tables("CSV")

        Finally
            cn.Close()
        End Try
    End Function

    ''' <summary>Reads a delimited string into a DataTable.</summary>
    ''' <param name="FileContents">The delimited file contents as a text string.</param>
    ''' <param name="HasHeader">Set True if the file includes a header.</param>
    ''' <returns>A DataTable.</returns>
    Public Function TextImportDelimitedString(ByVal FileContents As String, Optional ByVal HasHeader As Boolean = True, Optional ByVal Delimiter As String = ",", Optional ByVal StripQuotes As Boolean = True) As DataTable

        ' Parse the string to import
        Dim lines() As String = Split(FileContents, vbCrLf)
        If lines.Length = 0 Then Return Nothing

        ' Convert to a datatable
        Return TextImportDelimitedInt(lines, HasHeader, Delimiter, StripQuotes)

    End Function

    ''' <summary>Reads a delimited text file into a DataTable.</summary>
    ''' <param name="FullPath">The fully specified path and name of the file to import.</param>
    ''' <param name="HasHeader">Set True if the file includes a header.</param>
    ''' <returns>A DataTable.</returns>
    Public Function TextImportDelimited(ByVal FullPath As String, Optional ByVal HasHeader As Boolean = True, Optional ByVal Delimiter As String = ",", Optional ByVal StripQuotes As Boolean = True) As DataTable

        ' Read the file to import
        Dim lines() As String = File.ReadAllLines(FullPath)
        If lines.Length = 0 Then Return Nothing

        ' Convert to a datatable
        Return TextImportDelimitedInt(lines, HasHeader, Delimiter, StripQuotes)

    End Function

    Private Function TextImportDelimitedInt(ByVal lines() As String, Optional ByVal HasHeader As Boolean = True, Optional ByVal Delimiter As String = ",", Optional ByVal StripQuotes As Boolean = True) As DataTable

        ' Replace any quoted delimiters
        For i As Integer = 0 To lines.Length - 1
            lines(i) = HideDelimiters(lines(i), Delimiter)
        Next

        ' Get the max field count
        Dim CurFields As Integer
        Dim MaxFields As Integer = 0
        For i As Integer = 0 To lines.Length - 1
            CurFields = lines(i).Split(Delimiter).Length
            If CurFields > MaxFields Then MaxFields = CurFields
        Next

        ' Add the columns
        Dim dt As New DataTable
        For i As Integer = 1 To MaxFields
            dt.Columns.Add("Column " + i.ToString, GetType(System.String))
        Next

        ' Get the first record
        Dim rec() As String = lines(0).Split(Delimiter)
        If rec.Length = 0 Then Return Nothing

        ' Create the header
        Dim ColName As String
        For i As Integer = 0 To rec.Length - 1
            If HasHeader Then
                ColName = rec(i).ToString.Trim
                If StripQuotes Then ColName = Replace(ColName, """", "")
                dt.Columns(i).ColumnName = ColName
            End If
        Next

        ' Read the data
        Dim StartRec As Integer = 0
        If HasHeader Then StartRec = 1
        Dim fld As String
        For r As Integer = StartRec To lines.Length - 1
            rec = lines(r).Split(Delimiter)
            Dim NewRow As DataRow = dt.NewRow
            For c As Integer = 0 To rec.Length - 1
                fld = rec(c).ToString.Trim
                fld = Replace(fld, Chr(26), Delimiter)
                If StripQuotes Then fld = Replace(fld, """", "")
                NewRow(c) = fld
            Next
            dt.Rows.Add(NewRow)
        Next

        ' Return the DataTable
        Return dt
    End Function

    Private Function HideDelimiters(ByVal DelimStr As String, ByVal Delimiter As String) As String
        HideDelimiters = ""
        Dim fields() As String = Split(DelimStr, Delimiter)
        Dim Unclosed As Boolean = False
        For i As Integer = 0 To fields.Length - 1
            If Unclosed Then
                If fields(i).Trim.EndsWith("""") Then
                    Unclosed = False
                    HideDelimiters += fields(i) + Delimiter
                Else
                    fields(i) += Chr(26)
                    HideDelimiters += fields(i)
                End If
            Else
                If fields(i).Trim.StartsWith("""") And Not fields(i).Trim.EndsWith("""") Then
                    fields(i) += Chr(26)
                    Unclosed = True
                    HideDelimiters += fields(i)
                Else
                    HideDelimiters += fields(i) + Delimiter
                End If
            End If
        Next
        If HideDelimiters.EndsWith(Delimiter) Then HideDelimiters = Left(HideDelimiters, HideDelimiters.Length - Delimiter.Length)
    End Function

    ''' <summary>Reads an Excel worksheet into a DataTable.</summary>
    ''' <param name="FullPath">The fully specified path and name of the Excel file to import.</param>
    ''' <param name="SheetName">The name of the worksheet to import.</param>
    ''' <returns>A DataTable.</returns>
    ''' <remarks>Supports Excel 8.0 (Excel 97 and above)</remarks>
    Public Function ExcelImport(ByVal FullPath As String, ByVal SheetName As String, Optional ByVal HasHeader As Boolean = True) As DataTable
        Dim cn As OleDb.OleDbConnection = Nothing
        Try
            Dim filename As String = System.IO.Path.GetFileName(FullPath)
            Dim path As String = FullPath.Replace(filename, Nothing)

            Dim cnstr As String
            Dim headers As String
            If HasHeader Then headers = "Yes" Else headers = "No"
            ' Note that you would specify Excel 5.0 for Excel 95 workbooks
            'cnstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FullPath + ";Extended Properties='Excel 8.0;HDR=" + headers + ";IMEX=1'"
            cnstr = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + FullPath + ";Extended Properties='Excel 12.0;HDR=" + headers + ";IMEX=1'"
            cn = New OleDb.OleDbConnection(cnstr)
            Dim cmd As New OleDb.OleDbCommand("SELECT * FROM [" + SheetName + "]", cn)
            Dim da As New OleDb.OleDbDataAdapter(cmd)
            Dim ds As New DataSet()
            cn.Open()
            da.Fill(ds, SheetName)
            Return ds.Tables(SheetName)

        Finally
            cn.Close()
        End Try
    End Function

    ''' <summary>Reads the schema of an Excel spreadsheet into a DataTable.</summary>
    ''' <param name="FullPath">The fully specified path and name of the Excel file to read.</param>
    ''' <returns>A DataTable containing the spreadsheet schema.</returns>
    Public Function GetExcelSchema(ByVal FullPath As String) As DataTable
        'Dim cnstr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FullPath + ";Extended Properties=Excel 8.0;"
        Dim cnstr As String = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + FullPath + ";Extended Properties=Excel 12.0;"
        Dim cn As OleDb.OleDbConnection = New OleDb.OleDbConnection(cnstr)
        cn.Open()
        Return cn.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Tables, Nothing)
    End Function

    ''' <summary>Returns an XML string of a dataset.</summary>
    ''' <param name="dt">The dataset to convert.</param>
    ''' <returns>An XML string in standard recordset format.</returns>
    Public Function XmlFromDatatable(ByVal dt As DataTable) As String
        Dim sw As New StringWriter
        dt.WriteXml(sw, XmlWriteMode.WriteSchema)
        Return sw.ToString
    End Function

    ''' <summary>Converts an XML string to a dataset.</summary>
    ''' <param name="sXML">The standard recordset format XML string to convert.</param>
    ''' <returns>A dataset.</returns>
    Public Function XmlToDatatable(ByVal sXML As String) As DataTable
        Dim dt As New DataTable
        Dim rs As New StringReader(sXML)
        dt.ReadXmlSchema(rs)
        rs.Dispose()
        Dim rd As New StringReader(sXML)
        dt.ReadXml(rd)
        rd.Dispose()
        Return dt
    End Function

#End Region

End Class
