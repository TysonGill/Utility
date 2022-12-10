Imports Utility
Imports System.Data.SqlClient
Imports System.IO
Imports System.Windows.Forms

''' <summary>
''' Database Query Class
''' Features:
''' SQL Server, ODBC, and OLEDB Provider Types
''' Execute, GetValue, GetIdentity, and GetTable Query Types
''' All methods support Ad Hoc, Parameterized, and Stored Procedure calls
''' GetTable Queries with or without Schema Mapping
''' Synchronous or Asynchronous with Notification Event
''' Returns formatted sql for debugging parameterized queries
''' Sequential Transaction Support (not simultaneous)
''' Configurable Timeout
''' Automatic timing of all queries
''' Logs recent query history internally
''' Optional query logging to database
''' </summary>
Public Class clsDataQuery

#Region "Declarations"

    ''' <summary>Contains a history of recent queries</summary>
    Public Shared QueryHistory As DataTable

    ''' <summary>Used to lock the QueryHistory during updates</summary>
    Public Shared HistoryLock As New Object

    ''' <summary>The maximum number of QueryHistory records to retain</summary>
    Public QueryHistoryMax As Long = 100

    ''' <summary>Flag to disable parameterization
    ''' We have encountered SQL Server bugs where indices get misapplied with parameters are used
    ''' Dynamic queries do not suffer from this issue.</summary>
    Public EnableParameterization As Boolean = True

    ''' <summary>Enable translation of iSeries to SQL syntax</summary>
    Public EnableTranslateMHStoSQL As Boolean = False

    ''' <summary>A Tag property to identify the DataQuery</summary>
    Public Tag As Object = Nothing

    ''' <summary>The provider type, defaults to SQL Server</summary>
    Public ProviderType As ProviderTypes = ProviderTypes.SQLServer

    ''' <summary>The connection string used for all queries</summary>
    Public ConnectionString As String = ""

    ''' <summary>The sql statement prior to perparation</summary>
    Public sqlTemplate As String = ""

    ''' <summary>The query timeout in seconds, defaults to 15</summary>
    Public Timeout As Integer = 30

    ''' <summary>A collection of query paramters</summary>
    Private ParamCol As New SqlCommand

    ''' <summary>Stopwatch used to time the queries</summary>
    Private sw As New Stopwatch

    ''' <summary>The database connection object</summary>
    Private objCon As Object

    ''' <summary>The database transaction object</summary>
    Private objTrans As Object

    ''' <summary>The current provider</summary>
    Private CurProv As String = ""

    ''' <summary>The current database</summary>
    Private CurDB As String = ""

    ''' <summary>Supported database provider types</summary>
    Public Enum ProviderTypes
        SQLServer = 0
        ODBC = 1
        OLDEB = 2
    End Enum

    ''' <summary>Supported query types</summary>
    Public Enum QueryTypes
        None = 0
        Execute = 1
        GetValue = 2
        GetIdentity = 3
        GetTable = 4
    End Enum

    Private sqlPrepared As String = ""
    Private sqlDynamic As String = ""

    ''' <summary>Return an executable sql statement with any parameter values substituted</summary>
    Public ReadOnly Property sqlComplete() As String
        Get
            Return sqlPrepared
        End Get
    End Property

    Private LastQueryTimeInt As Integer = 0
    ''' <summary>Returns the timespan of the last query in milliseconds</summary>
    Public ReadOnly Property LastQueryTime() As Integer
        Get
            LastQueryTime = LastQueryTimeInt
        End Get
    End Property

    Private LastQueryValueInt As Object = Nothing
    ''' <summary>Returns the last value returned from GetValue</summary>
    Public ReadOnly Property LastQueryValue() As Object
        Get
            LastQueryValue = LastQueryValueInt
        End Get
    End Property

    Private LastQueryIdentityInt As Integer = 0
    ''' <summary>The last Identity returned from GetIdentity</summary>
    Public ReadOnly Property LastQueryIdentity() As Integer
        Get
            LastQueryIdentity = LastQueryIdentityInt
        End Get
    End Property

    Private LastQueryTableInt As DataTable = Nothing
    ''' <summary>The last datatable returned from GetTable</summary>
    Public ReadOnly Property LastQueryTable() As DataTable
        Get
            LastQueryTable = LastQueryTableInt
        End Get
    End Property

    Private LastQueryTypeInt As QueryTypes = QueryTypes.None
    ''' <summary>The current provider type, defaults to SQL Server</summary>
    Public ReadOnly Property LastQueryType() As QueryTypes
        Get
            LastQueryType = LastQueryTypeInt
        End Get
    End Property

    ''' <summary>Instantiate a new DataQuery object</summary>
    ''' <param name="Connection">The connection string</param>
    Public Sub New(Optional ByVal Connection As String = "", Optional ByVal Provider As ProviderTypes = ProviderTypes.SQLServer, Optional ByVal sql As String = "", Optional ByVal ConnectionName As String = "", Optional ByVal OptEnableParameterization As Boolean = True, Optional ByVal OptTranslateMHStoSQL As Boolean = False, Optional ByVal TimeoutSecs As Integer = 30)
        ConnectionString = Connection
        sqlTemplate = sql
        ProviderType = Provider
        CurProv = GetCurrentProvider()
        CurDB = GetCurrentDatabase()
        Timeout = TimeoutSecs
        EnableTranslateMHStoSQL = OptTranslateMHStoSQL
        EnableParameterization = OptEnableParameterization
        If QueryHistory Is Nothing Then
            SyncLock HistoryLock
                QueryHistory = New DataTable("Query History")
                QueryHistory.Columns.Add("Provider", GetType(String))
                QueryHistory.Columns.Add("Database", GetType(String))
                QueryHistory.Columns.Add("Query Date", GetType(String))
                QueryHistory.Columns.Add("Query Time in MS", GetType(Integer))
                QueryHistory.Columns.Add("Query Text", GetType(String))
            End SyncLock
        End If
    End Sub

#End Region

#Region "Execute"

    ''' <summary>Execute an action query</summary>
    ''' <param name="sql">The sql action query to execute</param>
    ''' <returns>The count of records affected</returns>
    Public Function Execute(ByVal sql As String) As Integer
        sqlTemplate = sql
        Return Execute()
    End Function

    ''' <summary>Execute an action query</summary>
    ''' <returns>The count of records affected</returns>
    Private Function Execute() As Integer
        Dim objCmd As Object = Nothing
        Try
            PreCheck()
            sw.Reset() : sw.Start()

            ' Determine which connection to use
            Dim sql As String = ""
            Dim CurCon As String = ConnectionString
            Dim CurProvType As ProviderTypes = ProviderType
            Dim Translate As Boolean = False
            sql = PrepareSQL()
            LogQueryStart(sql)

            ' Create the connection object and command objects
            Select Case CurProvType
                Case ProviderTypes.SQLServer

                    If objCon Is Nothing Then objCon = New SqlClient.SqlConnection(CurCon)
                    If objCon.State <> ConnectionState.Open Then objCon.Open()
                    If IsTransactional AndAlso objTrans Is Nothing Then
                        objTrans = CType(objTrans, SqlClient.SqlTransaction)
                        objTrans = objCon.BeginTransaction()
                    End If
                    objCmd = New SqlClient.SqlCommand(sql, objCon, objTrans)

                    ' Add parameters to command object
                    If EnableParameterization Then
                        For Each p As SqlParameter In ParamCol.Parameters
                            If ParamCol.Parameters(p.ParameterName).SqlDbType = SqlDbType.NVarChar Then
                                Dim sqp As SqlParameter = objCmd.Parameters.Add(p.ParameterName, SqlDbType.VarChar)
                                sqp.Value = p.Value
                                sqp.Size = GetStandardSize(sqp.Size)
                            Else
                                objCmd.Parameters.AddWithValue(p.ParameterName, p.Value)
                            End If
                        Next
                    End If

                Case ProviderTypes.OLDEB
                    If objCon Is Nothing Then objCon = New OleDb.OleDbConnection(CurCon)
                    If objCon.State <> ConnectionState.Open Then objCon.Open()
                    If IsTransactional AndAlso objTrans Is Nothing Then
                        objTrans = CType(objTrans, OleDb.OleDbTransaction)
                        objTrans = objCon.BeginTransaction()
                    End If
                    objCmd = New OleDb.OleDbCommand(sql, objCon, objTrans)
                Case Else

                    If objCon Is Nothing Then objCon = New Odbc.OdbcConnection(CurCon)
                    If objCon.State <> ConnectionState.Open Then objCon.Open()
                    If IsTransactional AndAlso objTrans Is Nothing Then
                        objTrans = CType(objTrans, Odbc.OdbcTransaction)
                        objTrans = objCon.BeginTransaction()
                    End If
                    objCmd = New Odbc.OdbcCommand(sql, objCon, objTrans)

                    ' Add parameters to command object
                    If EnableParameterization Then
                        For Each p As SqlParameter In ParamCol.Parameters
                            If ParamCol.Parameters(p.ParameterName).SqlDbType = SqlDbType.NVarChar Then
                                Dim sqp As Odbc.OdbcParameter = objCmd.Parameters.Add(p.ParameterName, SqlDbType.VarChar)
                                sqp.Value = p.Value
                                sqp.Size = GetStandardSize(sqp.Size)
                            Else
                                objCmd.Parameters.AddWithValue(p.ParameterName, p.Value)
                            End If
                        Next
                    End If

            End Select

            ' Get the data
            objCmd.CommandText = sql
            objCmd.CommandType = CommandType.Text
            objCmd.CommandTimeout = Timeout
            LastRowsAffected = objCmd.ExecuteNonQuery()
            LastQueryTypeInt = QueryTypes.Execute
            Return LastRowsAffected

        Catch Ex As Exception ' Return extended error information
            Dim NewMsg As String = "Database Query" + vbCrLf + "Target: " + CurDB + " on " + CurProv + vbCrLf + "Message: " + Ex.Message + vbCrLf + "SQL: " + sqlPrepared
            Dim exNew As New Exception(NewMsg, Ex)
            Throw exNew
        Finally
            If Not IsTransactional Then
                If objCon IsNot Nothing AndAlso objCon.State <> ConnectionState.Closed Then objCon.Close()
            End If
            ReleaseObj(objCmd)
            sw.Stop() : LastQueryTimeInt = sw.ElapsedMilliseconds
            LogQuery()
        End Try
    End Function

#End Region

#Region "Get Value"

    ''' <summary>Return a single value</summary>
    ''' <param name="sql">The sql string to return</param>
    ''' <returns>Returns the first row and field in the resultset</returns>
    Public Function GetValue(ByVal sql As String) As Object
        sqlTemplate = sql
        Return GetValue()
    End Function

    ''' <summary>Return a single value</summary>
    ''' <returns>Returns the first row and field in the resultset</returns>
    Private Function GetValue() As Object
        PreCheck()
        LastQueryTypeInt = QueryTypes.GetValue
        LastQueryValueInt = GetValueInternal()
        Return LastQueryValue
    End Function

    ''' <summary>Return an Identity value from SQL Server</summary>
    ''' <param name="sql">The sql string to insert or update</param>
    ''' <returns>Returns the row Identity after an Insert or Update</returns>
    ''' <remarks>SQL Server Only</remarks>
    Public Function GetIdentity(ByVal sql As String) As Integer
        sqlTemplate = sql
        Return GetIdentity()
    End Function

    ''' <summary>Return an Identity value from SQL Server</summary>
    ''' <returns>Returns the row Identity after an Insert or Update</returns>
    ''' <remarks>SQL Server Only</remarks>
    Public Function GetIdentity() As Integer
        PreCheck()
        LastQueryTypeInt = QueryTypes.GetIdentity
        LastQueryIdentityInt = Val(GetValueInternal(True))
        Return LastQueryIdentity
    End Function

    ''' <summary>Return a data value</summary>
    ''' <param name="GetSQLIdentity">Set to True to return the new Identity of an Insert action for SQL Server providers</param>
    ''' <returns>Returns the first row and field returned or the row Identity</returns>
    Private Function GetValueInternal(Optional ByVal GetSQLIdentity As Boolean = False) As Object
        Dim objCmd As Object = Nothing
        Try
            sw.Reset() : sw.Start()

            ' Determine which connection to use
            Dim sql As String = ""
            Dim CurCon As String = ConnectionString
            Dim CurProvType As ProviderTypes = ProviderType
            Dim Translate As Boolean = False
            sql = PrepareSQL()
            LogQueryStart(sql)

            ' Create the connection object and command objects
            Select Case CurProvType
                Case ProviderTypes.SQLServer

                    If objCon Is Nothing Then objCon = New SqlClient.SqlConnection(CurCon)
                    If objCon.State <> ConnectionState.Open Then objCon.Open()
                    If IsTransactional AndAlso objTrans Is Nothing Then
                        objTrans = CType(objTrans, SqlClient.SqlTransaction)
                        objTrans = objCon.BeginTransaction()
                    End If
                    If GetSQLIdentity Then sql += " SELECT scope_identity() OPTION (MAXDOP 1)"
                    objCmd = New SqlClient.SqlCommand(sql, objCon, objTrans)

                    ' Add parameters to command object
                    If EnableParameterization Then
                        For Each p As SqlParameter In ParamCol.Parameters
                            If ParamCol.Parameters(p.ParameterName).SqlDbType = SqlDbType.NVarChar Then
                                Dim sqp As SqlParameter = objCmd.Parameters.Add(p.ParameterName, SqlDbType.VarChar)
                                sqp.Value = p.Value
                                sqp.Size = GetStandardSize(sqp.Size)
                            Else
                                objCmd.Parameters.AddWithValue(p.ParameterName, p.Value)
                            End If
                        Next
                    End If

                Case ProviderTypes.OLDEB
                    If objCon Is Nothing Then objCon = New OleDb.OleDbConnection(CurCon)
                    If objCon.State <> ConnectionState.Open Then objCon.Open()
                    If IsTransactional AndAlso objTrans Is Nothing Then
                        objTrans = CType(objTrans, OleDb.OleDbTransaction)
                        objTrans = objCon.BeginTransaction()
                    End If
                    objCmd = New OleDb.OleDbCommand(sql, objCon, objTrans)
                Case Else

                    If objCon Is Nothing Then objCon = New Odbc.OdbcConnection(CurCon)
                    If objCon.State <> ConnectionState.Open Then objCon.Open()
                    If IsTransactional AndAlso objTrans Is Nothing Then
                        objTrans = CType(objTrans, Odbc.OdbcTransaction)
                        objTrans = objCon.BeginTransaction()
                    End If
                    objCmd = New Odbc.OdbcCommand(sql, objCon, objTrans)

                    ' Add parameters to command object
                    If EnableParameterization Then
                        For Each p As SqlParameter In ParamCol.Parameters
                            If ParamCol.Parameters(p.ParameterName).SqlDbType = SqlDbType.NVarChar Then
                                Dim sqp As Odbc.OdbcParameter = objCmd.Parameters.Add(p.ParameterName, SqlDbType.VarChar)
                                sqp.Value = p.Value
                                sqp.Size = GetStandardSize(sqp.Size)
                            Else
                                objCmd.Parameters.AddWithValue(p.ParameterName, p.Value)
                            End If
                        Next
                    End If

            End Select

            ' Get the data
            If Not GetSQLIdentity Then objCmd.CommandText = sql
            objCmd.CommandTimeout = Timeout
            GetValueInternal = objCmd.ExecuteScalar
            LastQueryTypeInt = QueryTypes.GetValue
            LastRowsAffected = 0

        Catch Ex As Exception ' Return extended error information
            Dim NewMsg As String = "Database Query" + vbCrLf + "Target: " + CurDB + " on " + CurProv + vbCrLf + "Message: " + Ex.Message + vbCrLf + "SQL: " + sqlPrepared
            Dim exNew As New Exception(NewMsg, Ex)
            Throw exNew
        Finally
            If Not IsTransactional Then
                If objCon IsNot Nothing AndAlso objCon.State <> ConnectionState.Closed Then objCon.Close()
            End If
            ReleaseObj(objCmd)
            sw.Stop() : LastQueryTimeInt = sw.ElapsedMilliseconds
            LogQuery()
        End Try
    End Function

#End Region

#Region "Get Table"

    ''' <summary>Return a data table (with schema mapping)</summary>
    ''' <returns>A datatable</returns>
    ''' <param name="TableName">The name to assign to the returned datatable</param>
    ''' <param name="sql">The sql string to return</param>
    ''' <param name="IncludeSchema">Set to False if you do not wish to return the table schema mapping</param>
    Public Function GetTable(ByVal TableName As String, ByVal sql As String, Optional ByVal IncludeSchema As Boolean = False) As DataTable
        Me.Tag = TableName
        sqlTemplate = sql
        If IncludeSchema Then Return GetTable() Else Return GetTableNoSchema()
    End Function

    ''' <summary>Return a data table (with schema mapping)</summary>
    ''' <returns>A datatable</returns>
    ''' <remarks>Table is given the name specified in the Tag property if a string</remarks>
    Private Function GetTable() As DataTable
        PreCheck()
        LastQueryTableInt = GetTableInternal()
        Return LastQueryTable
    End Function

    ''' <summary>Return a data table (without schema mapping)</summary>
    ''' <returns>A datatable</returns>
    ''' <remarks>Table is given the name specified in the Tag property if a string</remarks>
    Public Function GetTableNoSchema() As DataTable
        PreCheck()
        LastQueryTableInt = GetTableInternal(False)
        Return LastQueryTable
    End Function

    ''' <summary>Return a datatable</summary>
    ''' <remarks>Table is given the name specified in the Tag property if a string</remarks>
    Private Function GetTableInternal(Optional ByVal MapSchema As Boolean = True) As DataTable
        Dim objDA As Object = Nothing
        Try
            sw.Reset() : sw.Start()

            ' Determine which connection to use
            Dim sql As String = ""
            Dim CurCon As String = ConnectionString
            Dim CurProvType As ProviderTypes = ProviderType
            Dim Translate As Boolean = False
            sql = PrepareSQL()
            LogQueryStart(sql)

            ' Create the data adapter
            Select Case CurProvType
                Case ProviderTypes.SQLServer

                    objDA = New SqlClient.SqlDataAdapter(sql, CurCon)

                    ' Add parameters to command object
                    If EnableParameterization Then
                        For Each p As SqlParameter In ParamCol.Parameters
                            If ParamCol.Parameters(p.ParameterName).SqlDbType = SqlDbType.NVarChar Then
                                Dim sqp As SqlParameter = objDA.SelectCommand.Parameters.Add(p.ParameterName, SqlDbType.VarChar)
                                sqp.Value = p.Value
                                sqp.Size = GetStandardSize(sqp.Size)
                            Else
                                objDA.SelectCommand.Parameters.AddWithValue(p.ParameterName, p.Value)
                            End If
                        Next
                    End If

                Case ProviderTypes.OLDEB

                    objDA = New OleDb.OleDbDataAdapter(sql, CurCon)

                Case Else

                    objDA = New Odbc.OdbcDataAdapter(sql, CurCon)

                    ' Add parameters to command object
                    If EnableParameterization Then
                        For Each p As SqlParameter In ParamCol.Parameters
                            If ParamCol.Parameters(p.ParameterName).SqlDbType = SqlDbType.NVarChar Then
                                Dim sqp As Odbc.OdbcParameter = objDA.SelectCommand.Parameters.Add(p.ParameterName, SqlDbType.VarChar)
                                sqp.Value = p.Value
                                sqp.Size = GetStandardSize(sqp.Size)
                            Else
                                objDA.SelectCommand.Parameters.AddWithValue(p.ParameterName, p.Value)
                            End If
                        Next
                    End If

            End Select

            ' Get the data
            Dim TableName As String = "DataTable"
            If VarType(Tag) = VariantType.String Then TableName = Tag
            GetTableInternal = New DataTable(TableName)
            If MapSchema Then objDA.FillSchema(GetTableInternal, SchemaType.Mapped)
            objDA.SelectCommand.CommandText = sql
            objDA.SelectCommand.CommandTimeout = Timeout
            objDA.Fill(GetTableInternal)
            LastRowsAffected = GetTableInternal.Rows.Count

        Catch Ex As Exception ' Return extended error information
            Dim NewMsg As String = "Database Query" + vbCrLf + "Target: " + CurDB + " on " + CurProv + vbCrLf + "Message: " + Ex.Message + vbCrLf + "SQL: " + sqlPrepared
            Dim exNew As New Exception(NewMsg, Ex)
            Throw exNew
        Finally
            ReleaseObj(objDA)
            sw.Stop() : LastQueryTimeInt = sw.ElapsedMilliseconds
            LogQuery()
        End Try
    End Function

#End Region

#Region "General Support"

    ''' <summary>Verify the connection can be opened.</summary>
    ''' <returns>OK or an error message.</returns>
    Public Function IsConnectionValid() As String
        Dim con As String = ConnectionString.Trim
        If con = "" Then Return ""

        Dim objCmd As Object = Nothing
        Try
            If con.Trim = "" Then Return "Connection String not provided"
            Select Case ProviderType
                Case ProviderTypes.SQLServer
                    If objCon Is Nothing Then objCon = New SqlClient.SqlConnection(con)
                    If objCon.State <> ConnectionState.Open Then objCon.Open()
                Case ProviderTypes.OLDEB
                    If objCon Is Nothing Then objCon = New OleDb.OleDbConnection(con)
                    If objCon.State <> ConnectionState.Open Then objCon.Open()
                Case Else
                    If objCon Is Nothing Then objCon = New Odbc.OdbcConnection(con)
                    If objCon.State <> ConnectionState.Open Then objCon.Open()
            End Select
            Return "OK"

        Catch ex As Exception
            Return ex.Message
        Finally
            If objCon IsNot Nothing AndAlso objCon.State <> ConnectionState.Closed Then objCon.Close()
            ReleaseObj(objCmd)
        End Try
    End Function

    ''' <summary>Get default provider for current connection string.</summary>
    ''' <returns>The provider name.</returns>
    Public Function GetCurrentProvider() As String
        Dim con As String = ConnectionString.Trim
        If con = "" Then Return ""
        Select Case ProviderType
            Case ProviderTypes.SQLServer
                Dim cn As SqlClient.SqlConnection = New SqlClient.SqlConnection(con)
                Return cn.DataSource
            Case Else
                If con.StartsWith("DSN=") Then Return fmt.GetTextBetween(con, "DSN=", ";")
                Return fmt.GetTextBetween(con, "System=", ";")
        End Select
    End Function

    ''' <summary>Get default database for current connection string.</summary>
    ''' <returns>The database name.</returns>
    Public Function GetCurrentDatabase() As String
        Dim con As String = ConnectionString.Trim
        If con = "" Then Return ""
        Select Case ProviderType
            Case ProviderTypes.SQLServer
                Dim cn As SqlClient.SqlConnection = New SqlClient.SqlConnection(con)
                Return cn.Database
            Case Else
                If con.StartsWith("DSN=") Then Return ""
                Return fmt.GetTextBetween(con, "DBQ=", ";")
        End Select
    End Function

    ' Prepare the sqlTemplate for submission
    Private Function PrepareSQL() As String

        ' Get the sql template
        sqlPrepared = sqlTemplate

        ' Substitute parameters if parameterization is disabled or translating from iSeries
        If EnableTranslateMHStoSQL OrElse Not EnableParameterization Then sqlPrepared = SubstituteParameters(sqlPrepared)

        ' Translate to MHS if necessary
        If EnableTranslateMHStoSQL Then sqlPrepared = TranslateMHStoSQL(sqlPrepared)

        ' Return the prepared statement
        Return sqlPrepared

    End Function

    ' Substitute parameter values to produce a dynamic query
    Private Function SubstituteParameters(ByVal sql As String) As String

        ' Convert any parameters to dynamic sql
        For Each p As SqlParameter In ParamCol.Parameters
            If ProviderType = ProviderTypes.ODBC OrElse EnableTranslateMHStoSQL Then
                sql = Replace(sql, "?", fmt.q(p.Value), 1, 1, CompareMethod.Text)
            Else
                If p.DbType = DbType.String Or p.DbType = DbType.DateTime Then
                    sql = Replace(sql, p.ParameterName, fmt.q(p.Value), , , CompareMethod.Text)
                Else
                    sql = Replace(sql, p.ParameterName, p.Value.ToString, , , CompareMethod.Text)
                End If
            End If
        Next
        Return sql

    End Function

    ' Convert an MHS formatted query to SQL format
    Public Function TranslateMHStoSQL(ByVal sql As String) As String

        ' Replace quotes with brackets
        sql = fmt.ReplacePairs(sql, """", """", "[", "]")

        ' Remove libraries
        sql = System.Text.RegularExpressions.Regex.Replace(sql, "mhsfl.\.", "")
        sql = System.Text.RegularExpressions.Regex.Replace(sql, "finfl.\.", "")

        ' Replace FETCH with TOP
        Dim FetchIndex As Integer = sql.IndexOf("FETCH FIRST", CompareMethod.Text)
        If FetchIndex > -1 Then
            Dim RowLimit As String = fmt.GetTextBetween(sql, "FETCH FIRST ", " ROW")
            sql = sql.Substring(0, FetchIndex)
            If RowLimit <> "" Then
                sql = Replace(sql, "SELECT DISTINCT ", "SELECTX DISTINCT TOP " + RowLimit + " ", , , CompareMethod.Text)
                sql = Replace(sql, "SELECT ", "SELECT TOP " + RowLimit + " ", , , CompareMethod.Text)
                sql = Replace(sql, "SELECTX DISTINCT ", "SELECT DISTINCT ", , , CompareMethod.Text)
            End If
        End If

        ' Replace the concat function
        sql = TranslateConcat(sql)

        ' Replace other functions
        sql = Replace(sql, "TRIM(", "RTRIM(", , , CompareMethod.Text)
        sql = Replace(sql, "SUBSTR(", "SUBSTRING(", , , CompareMethod.Text)

        ' Bracket reserved words
        sql = Replace(sql, ".lineno", ".[lineno]", , , CompareMethod.Text)

        ' Handle hints
        sql = sql.Replace("FOR READ ONLY", "")
        sql = sql.Replace("WITH UR", "")

        ' Return the SQL format query
        Return sql

    End Function

    Private Function TranslateConcat(ByVal sql As String) As String
        Dim RightPos As Integer = 0
        Dim LeftPos As Integer = InStr(1, sql, "concat(", CompareMethod.Text)
        Do While LeftPos > 0
            RightPos = InStr(LeftPos + 1, sql, ")", CompareMethod.Text)
            If RightPos > 0 Then
                Dim fun As String = sql.Substring(LeftPos - 1, RightPos - LeftPos + 1)
                Dim params() As String = Split(fmt.GetTextBetween(fun, "(", ")"), ",")
                Dim p1 As String = params(0).Trim
                Dim p2 As String = params(1).Trim
                Dim NewConcat As String = p1 + " + " + p2
                sql = sql.Substring(0, LeftPos - 1) + NewConcat + sql.Substring(RightPos)
                LeftPos = InStr(LeftPos + 1, sql, "concat(", CompareMethod.Text)
            Else
                LeftPos = 0
            End If
        Loop
        Return sql
    End Function

    ''' <summary>Clean up an objects</summary>
    Protected Overrides Sub Finalize()
        If IsTransactional Then TransactionRollback()
        If objCon IsNot Nothing AndAlso objCon.State <> ConnectionState.Closed Then objCon.Close()
        ReleaseObj(objCon)
    End Sub

    ' Raise an error if the user has not completed prerequisites for a query
    Private Sub PreCheck()
        If ConnectionString.Trim = "" Then Throw New System.Exception("ConnectionString not provided")
        If sqlTemplate.Trim = "" Then Throw New System.Exception("sqlTemplate not provided")
    End Sub


#End Region

#Region "Parameter Support"

    ''' <summary>Create or assign a new parameter value</summary>
    ''' <param name="Name">The name of the parameter to add, e.g. @ID</param>
    ''' <param name="Value">The value of the specified parameter</param>
    Public Sub ParamSet(ByVal Name As String, ByVal Value As Object, Optional ByVal ClearPrevious As Boolean = False)

        Name = Name.Trim.ToUpper
        If ClearPrevious Then ParamsClear()
        If Value Is Nothing Then Value = DBNull.Value
        ParamCol.Parameters.AddWithValue(Name, Value)

    End Sub

    ''' <summary>Clear any defined parameters</summary>
    Public Sub ParamsClear()
        ParamCol.Parameters.Clear()
    End Sub

    ''' <summary>Get a standard string size so that string parameters will hit the cache properly</summary>
    Private Function GetStandardSize(ByVal ActualSize As Integer) As Integer
        If ActualSize <= 50 Then Return 50
        If ActualSize <= 100 Then Return 100
        If ActualSize <= 200 Then Return 200
        If ActualSize <= 500 Then Return 500
        If ActualSize <= 1000 Then Return 1000
        If ActualSize <= 10000 Then Return 10000
        If ActualSize <= 100000 Then Return 100000
        Return ActualSize
    End Function

#End Region

#Region "Transaction Support"

    ''' <summary>True if a transaction is currently in progress</summary>
    Private IsTransactional As Boolean = False

    ''' <summary>Initiate a new transaction</summary>
    ''' <remarks>Applies to Execute, GetValue, and GetIdentity methods</remarks>
    Public Sub TransactionBegin()
        IsTransactional = True
    End Sub

    ''' <summary>Abort the current transaction</summary>
    Public Sub TransactionRollback()
        If Not IsTransactional Then Exit Sub
        objTrans.Rollback()
        IsTransactional = False
        objTrans = Nothing
        If objCon IsNot Nothing AndAlso objCon.State <> ConnectionState.Closed Then objCon.Close()
    End Sub

    ''' <summary>Commit the current transaction</summary>
    Public Sub TransactionCommit()
        If Not IsTransactional Then Exit Sub
        objTrans.Commit()
        IsTransactional = False
        objTrans = Nothing
        If objCon IsNot Nothing AndAlso objCon.State <> ConnectionState.Closed Then objCon.Close()
    End Sub

#End Region

#Region "Multithreading Support"

    Event Finished(ByVal Sender As Object, ByVal QueryType As QueryTypes)

    ''' <summary>Execute an action query asynchronously</summary>
    ''' <remarks>Implement the Finished event if you wish to be notified of completion</remarks>
    Public Sub ExecuteAsync()
        PreCheck()
        Dim Exe As New System.Threading.Thread(AddressOf ExecuteDel)
        Exe.Start()
    End Sub

    ' Execute delegate
    Private Sub ExecuteDel()
        Execute()
        Raise(FinishedEvent, New Object() {Me, QueryTypes.Execute})
    End Sub

    ''' <summary>Execute a GetValue query asynchronously and populate LastQueryValue with the result</summary>
    ''' <remarks>Implement the Finished event if you wish to be notified of completion</remarks>
    Public Sub GetValueAsync()
        PreCheck()
        Dim Exe As New System.Threading.Thread(AddressOf GetValueDel)
        Exe.Start()
    End Sub

    ' GetValue delegate
    Private Sub GetValueDel()
        GetValue()
        Raise(FinishedEvent, New Object() {Me, QueryTypes.GetValue})
    End Sub

    ''' <summary>Execute a GetIdentity query asynchronously and populate LastQueryIdentity with the result</summary>
    ''' <remarks>Implement the Finished event if you wish to be notified of completion</remarks>
    Public Sub GetIdentityAsync()
        PreCheck()
        Dim Exe As New System.Threading.Thread(AddressOf GetIdentityDel)
        Exe.Start()
    End Sub

    ' GetIdentity delegate
    Private Sub GetIdentityDel()
        GetIdentity()
        Raise(FinishedEvent, New Object() {Me, QueryTypes.GetIdentity})
    End Sub

    ''' <summary>Execute a GetTable query asynchronously and populate LastQueryTable with the result</summary>
    ''' <remarks>Implement the Finished event if you wish to be notified of completion</remarks>
    Public Sub GetTableAsync()
        PreCheck()
        Dim Exe As New System.Threading.Thread(AddressOf GetTableDel)
        Exe.Start()
    End Sub

    ' GetTable delegate
    Private Sub GetTableDel()
        GetTable()
        Raise(FinishedEvent, New Object() {Me, QueryTypes.GetTable})
    End Sub

    ''' <summary>Execute a GetTableNoSchema query asynchronously and populate LastQueryTable with the result</summary>
    ''' <remarks>Implement the Finished event if you wish to be notified of completion</remarks>
    Public Sub GetTableNoSchemaAsync()
        PreCheck()
        Dim Exe As New System.Threading.Thread(AddressOf GetTableNoSchemaDel)
        Exe.Start()
    End Sub

    ' GetTable delegate
    Private Sub GetTableNoSchemaDel()
        GetTableNoSchema()
        Raise(FinishedEvent, New Object() {Me, QueryTypes.GetTable})
    End Sub

    ' Safely raise a cross-thread event
    Shared Sub Raise(ByVal [event] As [Delegate], ByVal data As Object())
        If [event] IsNot Nothing Then
            For Each C As Object In [event].GetInvocationList
                Dim T = CType(C.Target, System.ComponentModel.ISynchronizeInvoke)
                If T IsNot Nothing AndAlso T.InvokeRequired Then T.BeginInvoke(C, data) Else C.DynamicInvoke(data)
            Next
        End If
    End Sub

#End Region

#Region "History and File Logging"

    ' Log the query to the QueryHistory datatable
    Private Sub LogQuery()

        Dim sql As String = SubstituteParameters(sqlPrepared)

        ' Write to history datatable (shared across all DataQuery instances)
        SyncLock HistoryLock
            If QueryHistory.Rows.Count > QueryHistoryMax Then QueryHistory.Rows.RemoveAt(0)
            QueryHistory.Rows.Add(CurProv, CurDB, Now.ToShortDateString + " " + Now.ToShortTimeString, LastQueryTime, sql)
        End SyncLock

        ' Log long running queries to file
        If LogPath.Trim <> "" AndAlso LastQueryTime >= LogThreshold Then
            Dim LogEntry As String = fmt.Quote(Now.ToString) + ", " + fmt.Quote(Environment.MachineName) + ", " + fmt.Quote(CurProv) + ", " + fmt.Quote(CurDB) + ", " + fmt.Quote(LastQueryTime.ToString) + ", " + fmt.Quote(LastPing.ToString) + ", " + fmt.Quote(Replace(sql, """", "'")) + vbCrLf
            Dim LogFile As String = fmt.AppendIfNeeded(LogPath, "\") + CreateIdentifier() + ".log"
            AppendTextAsync(LogFile, LogEntry)
        End If

        ' Log to database
        LogQueryEnd(LastQueryTimeInt)

    End Sub

    Private Structure AppendTextAsyncArgs
        Dim FileName As String
        Dim FileText As String
    End Structure

    Private Sub AppendTextAsync(ByVal FileName As String, ByVal FileText As String)
        Dim r As New System.Threading.Thread(AddressOf AppendTextJob)
        r.Priority = System.Threading.ThreadPriority.BelowNormal
        r.IsBackground = False
        Dim Args As AppendTextAsyncArgs
        Args.FileName = FileName
        Args.FileText = FileText
        r.Start(Args)
    End Sub

    Private Sub AppendTextJob(ByVal Args As AppendTextAsyncArgs)
        If Not Directory.Exists(Path.GetPathRoot(Args.FileName)) Then Exit Sub
        Dim swf As New Stopwatch
        swf.Start()
        Do While swf.ElapsedMilliseconds < 5000
            Try
                File.AppendAllText(Args.FileName, Args.FileText)
                Return
            Catch
                Threading.Thread.Sleep(100)
            End Try
        Loop
        ' Retry time exceeded. No error. The line just doesn't get written.
    End Sub

    ' Create a unique date string for file naming
    Private Function CreateIdentifier() As String
        Dim SetDate As DateTime = Now
        CreateIdentifier = DateAndTime.Month(SetDate).ToString("D2") + "_" + DateAndTime.Day(SetDate).ToString("D2") + "_20" + DateAndTime.Year(SetDate).ToString.Substring(2, 2)
    End Function

#End Region

#Region "Database Logging"

    Private LastLogID As Integer = 0
    Private LastRowsAffected As Integer = 0

    ''' <summary>Log the query to the database.</summary>
    ''' <remarks>Note you must set the LogConnection connection to enable logging to the ApplicationShare database.</remarks>
    Public Sub LogQueryStart(ByVal sql As String)
        If LogConnection.Trim = "" Then Exit Sub

        ' Substitute any parameters
        sql = SubstituteParameters(sqlPrepared)

        ' Log the query
        Dim cn As New SqlClient.SqlConnection(LogConnection)
        Dim cmd As New SqlClient.SqlCommand("INSERT INTO Application_DB_Log (Machine_Name,Server_Name,User_Name,Query,Query_Date, App_Name) VALUES (@Mach,@Server,@User,@Query,@Dat," & fmt.q(Application.ProductName) & ");SELECT @@identity;", cn)
        Dim prmMach As New SqlClient.SqlParameter("@Mach", SqlDbType.VarChar)
        Dim prmServer As New SqlClient.SqlParameter("@Server", SqlDbType.VarChar)
        Dim prmUser As New SqlClient.SqlParameter("@User", SqlDbType.VarChar)
        Dim prmQuery As New SqlClient.SqlParameter("@Query", SqlDbType.VarChar)
        Dim prmDate As New SqlClient.SqlParameter("@Dat", SqlDbType.DateTime)
        Try
            cn.Open()
            prmMach.Value = Environment.MachineName
            prmServer.Value = GetCurrentProvider()
            prmUser.Value = nwk.GetUser
            prmQuery.Value = sql
            prmDate.Value = Now
            cmd.Parameters.Add(prmMach)
            cmd.Parameters.Add(prmServer)
            cmd.Parameters.Add(prmUser)
            cmd.Parameters.Add(prmQuery)
            cmd.Parameters.Add(prmDate)
            LastLogID = cmd.ExecuteScalar()
        Catch ex As Exception
            LastLogID = 0
            ReportError(ex, False)
        Finally
            If cn.State = ConnectionState.Open Then cn.Close()
        End Try
    End Sub

    ''' <summary>Update the query log with the elapsed query time.</summary>
    Public Sub LogQueryEnd(ByVal ElapsedMS As Integer)

        If LastLogID = 0 Then Exit Sub

        ' Update the query log entry
        Dim cn As New SqlClient.SqlConnection(LogConnection)
        Dim cmd As New SqlClient.SqlCommand("UPDATE Application_DB_Log SET query_Length=@qry, rows_affected=@rows WHERE DB_ID_Column=@id;", cn)
        Dim prmID As New SqlClient.SqlParameter("@ID", SqlDbType.Int)
        Dim prmQry As New SqlClient.SqlParameter("@qry", SqlDbType.Decimal, 18)
        Dim prmRows As New SqlClient.SqlParameter("@rows", SqlDbType.Int)
        Try
            cn.Open()
            prmID.Value = LastLogID
            prmQry.Value = ElapsedMS
            prmRows.Value = LastRowsAffected
            cmd.Parameters.Add(prmID)
            cmd.Parameters.Add(prmQry)
            cmd.Parameters.Add(prmRows)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            ReportError(ex, False)
        Finally
            LastLogID = 0
            If cn.State = ConnectionState.Open Then cn.Close()
        End Try
    End Sub

#End Region

End Class
