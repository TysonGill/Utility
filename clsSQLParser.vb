''' <summary>Parses and reformats SQL queries.</summary>
''' <remarks>Supprts TSQL queries only.</remarks>
Public Class clsSQLParser

    ''' <summary>The parsed query portions.</summary>
    Public tblParsed As New DataTable

    ''' <summary>The where clause of the query.</summary>
    Public WhereClause As String

    ''' <summary>The sort clause of the query.</summary>
    Public SortClause As String

    ''' <summary>Instantiate a SQLParser instance and optionally parse a SQL string.</summary>
    ''' <param name="sql">The SQL query to parse.</param>
    ''' <remarks>Should be a SELECT query.</remarks>
    Public Sub New(Optional ByVal sql As String = "")
        If sql.Trim <> "" Then Parse(sql)
    End Sub

    ''' <summary>Parse a SQL SELECT query.</summary>
    ''' <param name="sql">The SQL SELECT query to parse.</param>
    Public Sub Parse(ByVal sql As String)

        Dim sqlWork As String = sql

        SortClause = GetLastPart(sqlWork, "ORDER BY")
        sqlWork = Microsoft.VisualBasic.Left(sqlWork, sqlWork.Length - SortClause.Length).Trim

        WhereClause = GetLastPart(sqlWork, "WHERE")
        sqlWork = Microsoft.VisualBasic.Left(sqlWork, sqlWork.Length - WhereClause.Length).Trim

        Dim Tables As DataTable = GetTables(sqlWork)
        Dim Fields As DataTable = GetFields(sqlWork)

        Dim Table As String
        Dim rows As DataRow()
        For Each r As DataRow In Fields.Rows
            Table = r("Table")
            rows = Tables.Select("[Alias] = '" + Table + "'")
            If rows IsNot Nothing AndAlso rows.Length = 1 AndAlso rows(0)("Table").trim <> "" Then
                r("Table") = rows(0)("Table")
            End If
        Next

        tblParsed = Fields

    End Sub

    ''' <summary>Assign a set of query parameters from a table row.</summary>
    ''' <param name="dt">The source table.</param>
    ''' <param name="RowNumber">The row number to insert.</param>
    Public Sub AssignValues(ByVal dt As DataTable, ByVal RowNumber As Integer)

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Exit Sub
        If RowNumber > dt.Rows.Count - 1 OrElse RowNumber < 0 Then
            Throw New Exception("Invalid RowNumber " + RowNumber.ToString)
        End If

        ClearValues()

        Dim FieldName As String
        For Each r As DataRow In tblParsed.Rows
            If r("Alias") <> "" Then FieldName = r("Alias") Else FieldName = r("Field")
            If dt.Columns.Contains(FieldName) Then r("Value") = dt.Rows(RowNumber)(FieldName).ToString
        Next

    End Sub

    ''' <summary>Clear all assigned values.</summary>
    Private Sub ClearValues()

        If tblParsed Is Nothing OrElse tblParsed.Rows.Count = 0 Then Exit Sub

        For Each r As DataRow In tblParsed.Rows
            r("Value") = ""
        Next

    End Sub

    ''' <summary>Get an INSERT version of the query.</summary>
    Public Function GetInsert(ByVal TableName As String) As String

        Dim sql As String = "INSERT INTO " + TableName
        Dim rows As DataRow() = tblParsed.Select("[Table] = '" + TableName + "'")
        If rows Is Nothing OrElse rows.Length = 0 Then
            Throw New Exception("No fields for table '" + TableName + "'")
        End If

        sql += " ("
        For Each r As DataRow In rows
            sql += "[" + r("Field") + "], "
        Next
        sql = Microsoft.VisualBasic.Left(sql, sql.Length - 2) + ") VALUES ("

        For Each r As DataRow In rows
            sql += "'" + r("Value") + "', "
        Next
        sql = Microsoft.VisualBasic.Left(sql, sql.Length - 2) + ")"

        Return sql

    End Function

    ''' <summary>Get an UPDATE version of the query.</summary>
    Public Function GetUpdate(ByVal TableName As String) As String

        Dim sql As String = "UPDATE " + TableName + " SET "
        Dim rows As DataRow() = tblParsed.Select("[Table] = '" + TableName + "'")
        If rows Is Nothing OrElse rows.Length = 0 Then
            Throw New Exception("No fields for table '" + TableName + "'")
        End If

        For Each r As DataRow In rows
            sql += "[" + r("Field") + "] = '" + r("Value") + "', "
        Next
        sql = Microsoft.VisualBasic.Left(sql, sql.Length - 2)

        rows = tblParsed.Select("[Table] = '" + TableName + "' AND [IsKey] = 'True'")
        If rows Is Nothing OrElse rows.Length <> 1 Then
            Throw New Exception("No key field for table '" + TableName + "'")
        End If
        sql += " WHERE [" + rows(0)("Field") + "] = '" + rows(0)("Value")

        Return sql

    End Function

    ''' <summary>Get all tables references by a query.</summary>
    ''' <param name="sql">The query to parse.</param>
    ''' <returns>A table of table names.</returns>
    Private Function GetTables(ByVal sql As String) As DataTable

        ' Create a data table to return
        Dim tblTables As New DataTable
        tblTables.Columns.Add("Table", GetType(System.String))
        tblTables.Columns.Add("Alias", GetType(System.String))

        ' Replace all supported terminators with a standard value
        sql = sql.ToUpper
        sql = sql.Replace("INNER JOIN", "TABLEEND")
        sql = sql.Replace("OUTER JOIN", "TABLEEND")
        sql = sql.Replace("LEFT JOIN", "TABLEEND")
        sql = sql.Replace("JOIN", "TABLEEND")
        sql = sql.Replace("WHERE", "TABLEEND")
        sql = sql.Replace("ORDER BY", "TABLEEND")
        If InStr(sql, "TABLEEND") = 0 Then sql += " TABLEEND"
        sql = GetLastPart(sql, "FROM ")
        sql = "TABLEEND " + Microsoft.VisualBasic.Right(sql, sql.Length - 5).Trim

        ' Get all datatables
        Dim TableSpecs() As String = Split(sql, "TABLEEND")
        For Each TableSpec As String In TableSpecs
            GetTable(TableSpec, tblTables)
        Next

        Return tblTables

    End Function

    Private Sub GetTable(ByVal TableSpec As String, ByVal tblTables As DataTable)

        TableSpec = TableSpec.Trim
        If TableSpec = "" Then Exit Sub

        ' Remove any ON clause
        TableSpec = GetFirstPart(TableSpec, " ON ")

        ' Parse the tablespec
        Dim TableName As String = ""
        Dim TableAlias As String = ""
        Dim TableParts As String() = Split(TableSpec, " ", , CompareMethod.Text)
        If TableParts.Length = 1 Then
            TableName = TableSpec
            TableAlias = ""
        ElseIf TableParts.Length = 2 Then
            TableName = TableParts(0).Trim
            TableAlias = TableParts(1).Trim
        End If

        tblTables.Rows.Add(TableName, TableAlias)

    End Sub

    Private Function GetFields(ByVal sql As String) As DataTable

        ' Create a data table to return
        Dim tblFields As New DataTable
        tblFields.Columns.Add("Position", GetType(System.Int32))
        tblFields.Columns.Add("Table", GetType(System.String))
        tblFields.Columns.Add("Field", GetType(System.String))
        tblFields.Columns.Add("Alias", GetType(System.String))
        tblFields.Columns.Add("IsKey", GetType(System.String))
        tblFields.Columns.Add("Value", GetType(System.String))

        ' Parse the fields
        Dim Fields() As String = Split(fmt.GetTextBetween(sql, "SELECT", "FROM"), ",")
        Dim IsKey As String = ""
        Dim FieldParts() As String
        Dim TableName As String = ""
        Dim FieldName As String = ""
        Dim FieldAlias As String = ""
        For Position As Integer = 0 To Fields.Length - 1
            FieldParts = Split(Fields(Position), "AS", , CompareMethod.Text)
            If FieldParts.Length = 1 Then
                FieldName = FieldParts(0).Trim.ToUpper
                FieldAlias = ""
            ElseIf FieldParts.Length = 2 Then
                FieldName = FieldParts(0).Trim.ToUpper
                FieldAlias = FieldParts(1).Trim
            End If
            FieldParts = Split(FieldName, ".", , CompareMethod.Text)
            If FieldParts.Length = 1 Then
                TableName = ""
            ElseIf FieldParts.Length = 2 Then
                TableName = FieldParts(0).Trim
                FieldName = FieldParts(1).Trim
            End If
            If FieldName.StartsWith("[") Then FieldName = fmt.GetTextBetween(FieldName, "[", "]")
            If FieldName.StartsWith("'") Then FieldName = fmt.GetTextBetween(FieldName, "'", "'")
            If FieldAlias.StartsWith("[") Then FieldAlias = fmt.GetTextBetween(FieldAlias, "[", "]")
            If FieldAlias.StartsWith("'") Then FieldAlias = fmt.GetTextBetween(FieldAlias, "'", "'")
            If (Position = 0) Then IsKey = "True" Else IsKey = ""
            tblFields.Rows.Add(Position + 1, TableName, FieldName, FieldAlias, IsKey, "")

        Next
        Return tblFields

    End Function

    Private Function GetFirstPart(ByVal Full As String, ByVal ToPart As String) As String

        Dim endpos As Integer = InStrRev(Full, ToPart, , CompareMethod.Text)
        If endpos = 0 Then Return Full

        GetFirstPart = Microsoft.VisualBasic.Left(Full, endpos - 1).Trim

    End Function

    Private Function GetLastPart(ByVal Full As String, ByVal FromPart As String) As String

        Dim startpos As Integer = InStrRev(Full, FromPart, , CompareMethod.Text)
        If startpos = 0 Then Return ""

        GetLastPart = Microsoft.VisualBasic.Right(Full, Full.Length - startpos + 1).Trim

    End Function

End Class
