Imports System
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Data.OleDb
Imports MySql.Data.MySqlClient
Imports System.Data.Odbc

Public Class GeneralizedDAL
    Private _con As IDbConnection
    Private _conMySQL As IDbConnection
    Private _cmd As IDbCommand
    Private _adpt As IDbDataAdapter
    Private _reader As IDataReader
    Private _dbparameter As IDbDataParameter
    Private _connectionString As String
    Private _provider As EnumProvider

    '**** Constructor for initializing variables ****/
#Region "Constructor"
    Public Sub New()
        _provider = EnumProvider.SQL
        _con = GetConnectionObject()
        _cmd = GetCommandObject()
        _adpt = GetDataAdapterObject()
        'If Not ConfigurationSettings.AppSettings("FuelSecureConnectionString") = Nothing Then
        '    _connectionString = ConfigurationSettings.AppSettings("FuelSecureConnectionString").ToString()
        '    If (_connectionString.StartsWith(",")) Then
        '        Dim length As Integer = _connectionString.Length
        '        length = length - 1
        '        _connectionString = _connectionString.Substring(1, length)
        '    End If
        'End If
        If Not ConfigurationManager.ConnectionStrings("FluidSecureConnectionString").ConnectionString = Nothing Then
            _connectionString = ConfigurationManager.ConnectionStrings("FluidSecureConnectionString").ConnectionString.ToString()
            If (_connectionString.StartsWith(",")) Then
                Dim length As Integer = _connectionString.Length
                length = length - 1
                _connectionString = _connectionString.Substring(1, length)
            End If
        End If
    End Sub
#End Region


    Public Property ConnectionString() As String
        Get
            Return _connectionString
        End Get
        Set(ByVal Value As String)
            _connectionString = Value
        End Set
    End Property

    '*** Function which opens connection and return true if conenction 
    ' successful otherwise false ***/
    Public Function OpenConnection() As Boolean
        If _con Is Nothing Then
            _con = GetConnectionObject()
            _con.ConnectionString = _connectionString
            _con.Open()
        ElseIf _con.State = ConnectionState.Closed Then
            _con = GetConnectionObject()
            _con.ConnectionString = _connectionString
            _con.Open()
        End If

        Return True
    End Function

    Public Function GetConnection() As IDbConnection

        If _con Is Nothing Then
            _con = GetConnectionObject()
            _con.ConnectionString = _connectionString
            _con.Open()
        ElseIf _con.State = ConnectionState.Open Then
            _con.Close()
            _con.Open()
        Else
            _con.Open()
        End If
        Return _con
    End Function

#Region "Get DataSet"
    '*** Fuction will return a dataset for the query passed to it ***/
    Public Function GetDataSet(ByVal strQuery As String) As DataSet
        Try
            Dim ds = New DataSet()
            _cmd = GetCommandObject()
            _adpt = GetDataAdapterObject()
            If (OpenConnection()) Then
                _cmd.Connection = _con
                _cmd.CommandText = strQuery
                _adpt.SelectCommand = _cmd
                _adpt.Fill(ds)
                Return ds
            End If
            Return Nothing
        Catch ex As Exception
            If (_con.State.Equals(ConnectionState.Open)) Then
                _con.Close()
                Throw ex
            End If
        Finally
            _con.Close()
        End Try
    End Function

    Public Function GetDataTable(ByVal strQuery As String) As DataTable
        'Try
        '    Dim dt = New DataTable()
        '    _cmd = GetCommandObject()
        '    _adpt = GetDataAdapterObject()
        '    If (OpenConnection()) Then
        '        _cmd.Connection = _con
        '        _cmd.CommandText = strQuery
        '        _adpt.SelectCommand = _cmd
        '        _adpt.Fill(dt)
        '        Return dt
        '    End If
        '    Return Nothing
        'Catch ex As Exception
        '    If (_con.State.Equals(ConnectionState.Open)) Then
        '        _con.Close()
        '        Throw ex
        '    End If
        'Finally
        '    _con.Close()
        'End Try
    End Function
#End Region

    Public Function GetDataSetbool(ByVal strQuery As String) As Boolean
        Try
            Dim ds = New DataSet()
            _cmd = GetCommandObject()
            _adpt = GetDataAdapterObject()
            If (OpenConnection()) Then
                _cmd.Connection = _con
                _cmd.CommandText = strQuery
                _adpt.SelectCommand = _cmd
                _adpt.Fill(ds)
                If (ds.Tables(0).Rows.Count > 0) Then
                    Return True
                Else
                    Return False
                End If
            End If
            Return False
        Catch ex As Exception
            Throw ex
        Finally
            If (_con.State.Equals(ConnectionState.Open)) Then
                _con.Close()
            End If
        End Try
    End Function

#Region "Get DataReader"
    '*** Fuction will return a datareader for the query passed to it ***/
    Public Function GetDataReader(ByVal strQuery As String) As IDataReader
        _reader = Nothing
        Try
            If (OpenConnection()) Then
                _cmd = GetCommandObject()
                _cmd.Connection = _con
                _cmd.CommandText = strQuery
                _reader = _cmd.ExecuteReader(CommandBehavior.CloseConnection)
                Return _reader
            End If
            Return Nothing
        Catch ex As Exception
            Throw ex
        Finally
            If (_con.State.Equals(ConnectionState.Open)) Then
                _con.Close()
            End If
        End Try
    End Function
#End Region

#Region "ExecuteScalar"
    '*** Returns first row and first column value from the query result ***/

    Public Function ExecuteScalarGetInteger(ByVal strQuery As String) As Integer
        Try
            Dim obj As Object = Nothing
            If (OpenConnection()) Then
                _cmd = GetCommandObject()
                _cmd.CommandText = strQuery
                _cmd.Connection = _con
                If Convert.IsDBNull(_cmd.ExecuteScalar()) = False Then obj = _cmd.ExecuteScalar()
                If (Not obj = Nothing) Then
                    Dim result As Int32 = Convert.ToInt32(obj.ToString())
                    Return result
                End If
            End If
            Return -1
        Catch ex As Exception
            Throw ex
        Finally
            If (_con.State.Equals(ConnectionState.Open)) Then
                _con.Close()
            End If
        End Try
    End Function

    Public Function ExecuteScalarGetInteger(ByVal strQuery As String, ByVal parCollection() As IDataParameter) As Integer
        Try
            Dim obj As Object = Nothing
            If (OpenConnection()) Then
                _cmd = GetCommandObject()
                _cmd.CommandText = strQuery
                _cmd.Connection = _con
                Dim par As IDataParameter = GetDataParameterObject()
                Dim i As Integer = 0
                For i = 0 To parCollection.Length - 1
                    par = GetDataParameterObject()
                    par.ParameterName = parCollection(i).ParameterName
                    par.DbType = parCollection(i).DbType
                    par.Value = parCollection(i).Value
                    _cmd.Parameters.Add(par)
                Next
                If Convert.IsDBNull(_cmd.ExecuteScalar()) = False Then obj = _cmd.ExecuteScalar()
                If (Not obj = Nothing) Then
                    Dim result As Int32 = Convert.ToInt32(obj.ToString())
                    Return result
                End If
            End If
            Return -1
        Catch ex As Exception
            Throw ex
        Finally
            If (_con.State.Equals(ConnectionState.Open)) Then
                _con.Close()
            End If
        End Try
    End Function

    '*** Returns first row and first column value from the query result ***/

    Public Function ExecuteScalarGetString(ByVal strQuery As String) As String
        Try
            Dim obj As Object = Nothing
            If (OpenConnection()) Then
                _cmd = GetCommandObject()
                _cmd.CommandText = strQuery
                _cmd.Connection = _con
                obj = _cmd.ExecuteScalar()
                If Not obj Is DBNull.Value Then
                    Return obj.ToString()
                End If
            End If
            Return ""
        Catch ex As Exception
            Throw ex
        Finally
            If (_con.State.Equals(ConnectionState.Open)) Then
                _con.Close()
            End If
        End Try
    End Function

    Public Function ExecuteScalarGetString(ByVal strQuery As String, ByVal parCollection() As IDataParameter) As String
        Try
            Dim obj As Object = Nothing
            If (OpenConnection()) Then
                _cmd = GetCommandObject()
                _cmd.CommandText = strQuery
                _cmd.Connection = _con
                Dim par As IDataParameter = GetDataParameterObject()
                Dim i As Integer = 0
                For i = 0 To parCollection.Length - 1
                    par = GetDataParameterObject()
                    par.ParameterName = parCollection(i).ParameterName
                    par.DbType = parCollection(i).DbType
                    par.Value = parCollection(i).Value
                    _cmd.Parameters.Add(par)
                Next
                obj = _cmd.ExecuteScalar()
                If Not obj Is DBNull.Value Then
                    Return obj.ToString()
                End If
            End If
            Return ""
        Catch ex As Exception
            Throw ex
        Finally
            If (_con.State.Equals(ConnectionState.Open)) Then
                _con.Close()
            End If
        End Try
    End Function

    Public Function ExecuteScalarGetObject(ByVal strQuery As String) As Object
        Try
            Dim obj As Object = Nothing
            If (OpenConnection()) Then
                _cmd = GetCommandObject()
                _cmd.CommandText = strQuery
                _cmd.Connection = _con
                obj = _cmd.ExecuteScalar()
                If Not obj = Nothing Then
                    Return obj
                End If
            End If
            Return Nothing
        Catch ex As Exception
            Throw ex
        Finally
            If (_con.State.Equals(ConnectionState.Open)) Then
                _con.Close()
            End If
        End Try
    End Function
#End Region

#Region "ExecuteNonQuery Operations"
    Public Function ExecuteNonQuery(ByVal strQuery As String) As Integer
        Try
            If (OpenConnection()) Then
                _cmd = GetCommandObject()
                _cmd.Connection = _con
                _cmd.CommandText = strQuery
                Return _cmd.ExecuteNonQuery()
            End If
        Catch ex As Exception
            Throw ex
        Finally
            If (_con.State.Equals(ConnectionState.Open)) Then
                _con.Close()
            End If
        End Try
    End Function
#End Region

#Region "Excute StoreProcedure"
    '*** Initialize connection object depending upon the provider ***/
    Public Function ExecuteStoredProcedureGetString(ByVal strSPName As String, ByVal parCollection() As IDataParameter) As String
        Try
            If (OpenConnection()) Then
                _cmd = GetCommandObject()
                _cmd.Connection = _con
                _cmd.CommandText = strSPName
                _cmd.CommandType = CommandType.StoredProcedure
                Dim par As IDataParameter = GetDataParameterObject()
                Dim i As Integer = 0
                For i = 0 To parCollection.Length - 1
                    par = GetDataParameterObject()
                    par.ParameterName = parCollection(i).ParameterName
                    par.DbType = parCollection(i).DbType
                    par.Value = parCollection(i).Value
                    _cmd.Parameters.Add(par)
                Next
                Dim val As String = Nothing
                val = Convert.ToString(_cmd.ExecuteScalar())
                Return val
            End If
            Return Nothing
        Catch ex As Exception
            Throw ex
        Finally
            If (_con.State.Equals(ConnectionState.Open)) Then
                _con.Close()
            End If
        End Try
    End Function
#End Region

    Public Function ExecuteStoredProcedureGetBoolean(ByVal strSPName As String) As Boolean
        Try
            If (OpenConnection()) Then
                _cmd = GetCommandObject()
                _cmd.Connection = _con
                _cmd.CommandText = strSPName
                _cmd.CommandType = CommandType.StoredProcedure
                Dim result As Integer = _cmd.ExecuteNonQuery()
                If (result > 0) Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            Throw ex
        Finally
            If (_con.State.Equals(ConnectionState.Open)) Then
                _con.Close()
            End If
        End Try
    End Function

    Public Function ExecuteStoredProcedureGetInteger(ByVal strSPName As String) As Integer
        Try
            _cmd = GetCommandObject()
            _cmd.Connection = _con
            _cmd.CommandText = strSPName
            _cmd.CommandType = CommandType.StoredProcedure
            Return _cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        Finally
            If (_con.State.Equals(ConnectionState.Open)) Then
                _con.Close()
            End If
        End Try
    End Function


    Public Function ExecuteStoredProcedureGetBoolean(ByVal strSPName As String, ByVal parCollection() As IDataParameter) As Boolean
        Try
            If (OpenConnection()) Then
                _cmd = GetCommandObject()
                _cmd.Connection = _con
                _cmd.CommandText = strSPName
                _cmd.CommandType = CommandType.StoredProcedure
                Dim par As IDataParameter = GetDataParameterObject()
                Dim i As Integer = 0
                For i = 0 To parCollection.Length - 1
                    par = GetDataParameterObject()
                    par.ParameterName = parCollection(i).ParameterName
                    par.DbType = parCollection(i).DbType
                    par.Value = parCollection(i).Value
                    _cmd.Parameters.Add(par)
                Next
                Dim result As Object = _cmd.ExecuteScalar()
                _cmd.Parameters.Clear()
                If IsDBNull(result) And result > 0 Then
                    Return False
                ElseIf result = Nothing Then
                    Return False
                Else
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            Throw ex
        Finally
            If (_con.State.Equals(ConnectionState.Open)) Then
                _con.Close()
            End If
        End Try
    End Function

    Public Function ExecuteStoredProcedureGetInteger(ByVal strSPName As String, ByVal parCollection() As IDataParameter) As Integer
        Dim result As Object
        Try
            If (OpenConnection()) Then
                _cmd = GetCommandObject()
                _cmd.Connection = _con
                _cmd.CommandText = strSPName
                _cmd.CommandType = CommandType.StoredProcedure
                Dim par As IDataParameter = GetDataParameterObject()
                Dim i As Integer = 0
                For i = 0 To parCollection.Length - 1
                    par = GetDataParameterObject()
                    par.ParameterName = parCollection(i).ParameterName
                    par.DbType = parCollection(i).DbType
                    par.Value = parCollection(i).Value
                    _cmd.Parameters.Add(par)
                Next
                result = _cmd.ExecuteScalar()
                'result = _cmd.ExecuteNonQuery()
                _cmd.Parameters.Clear()
                Return Convert.ToInt32(result)
            End If
            Return 0
        Catch ex As Exception
            Throw ex
        Finally
            If (_con.State.Equals(ConnectionState.Open)) Then
                _con.Close()
            End If
        End Try
    End Function

    Public Function ExecuteStoredProcedureGetDataSet(ByVal strSPName As String, ByVal parCollection() As IDataParameter) As DataSet
        Try
            If (OpenConnection()) Then
                _cmd = GetCommandObject()
                _cmd.Connection = _con
                _cmd.CommandText = strSPName
                _cmd.CommandType = CommandType.StoredProcedure
                Dim par As IDataParameter = GetDataParameterObject()
                Dim i As Integer = 0
                For i = 0 To parCollection.Length - 1
                    par = GetDataParameterObject()
                    par.ParameterName = parCollection(i).ParameterName
                    par.DbType = parCollection(i).DbType
                    par.Value = parCollection(i).Value
                    _cmd.Parameters.Add(par)
                Next
                _adpt = GetDataAdapterObject()
                _adpt.SelectCommand = _cmd
                Dim ds As DataSet = New DataSet()
                _adpt.Fill(ds)
                _cmd.Parameters.Clear()
                Return ds
            End If
            Return Nothing
        Catch ex As Exception
            Throw ex
        Finally
            If (_con.State.Equals(ConnectionState.Open)) Then
                _con.Close()
            End If
        End Try
    End Function

    Public Function ExecuteStoredProcedureGetDataSet(ByVal strSPName As String) As DataSet
        Try
            If (OpenConnection()) Then
                _cmd = GetCommandObject()
                _cmd.Connection = _con
                _cmd.CommandText = strSPName
                _cmd.CommandType = CommandType.StoredProcedure
                _adpt = GetDataAdapterObject()
                _adpt.SelectCommand = _cmd
                Dim ds As DataSet = New DataSet()
                _adpt.Fill(ds)
                _cmd.Parameters.Clear()
                Return ds
            End If
            Return Nothing
        Catch ex As Exception
            Throw ex
        Finally
            If (_con.State.Equals(ConnectionState.Open)) Then
                _con.Close()
            End If
        End Try
    End Function

#Region "New Function"
    Public Function ExecuteSQLStoredProcedureGetBoolean(ByVal strSPName As String, ByVal parCollection() As SqlParameter) As Boolean
        Try
            If (OpenConnection()) Then
                _cmd = GetCommandObject()
                _cmd.Connection = _con
                _cmd.CommandText = strSPName
                _cmd.CommandType = CommandType.StoredProcedure
                Dim par As SqlParameter = New SqlParameter()
                Dim i As Integer = 0
                For i = 0 To parCollection.Length - 1
                    par = New SqlParameter()
                    par.ParameterName = parCollection(i).ParameterName
                    par.SqlDbType = parCollection(i).SqlDbType
                    par.Value = parCollection(i).Value
                    _cmd.Parameters.Add(par)
                Next
                Dim result As Integer = _cmd.ExecuteNonQuery()
                _cmd.Parameters.Clear()
                If (result > 0) Then
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            If (_con.State.Equals(ConnectionState.Open)) Then
                _con.Close()
            End If
        End Try
    End Function


    Public Function ExecuteStoredProcedureGetDataAdapter(ByVal strSPName As String) As IDataAdapter

        Try
            If (OpenConnection()) Then
                _cmd = GetCommandObject()
                _cmd.Connection = _con
                _cmd.CommandText = strSPName
                _cmd.CommandType = CommandType.StoredProcedure
                _adpt = New SqlDataAdapter(_cmd)
                Return _adpt
            End If
        Catch ex As Exception
            Throw ex
        Finally
            If (_con.State.Equals(ConnectionState.Open)) Then
                _con.Close()
            End If
        End Try
    End Function


    Public Function ExecuteStoredProcedureGetInteger(ByVal con As IDbConnection, ByVal trans As IDbTransaction, ByVal strSPName As String, ByVal parCollection() As IDataParameter) As Integer
        Try
            _cmd = GetCommandObject()
            _cmd.Connection = con
            _cmd.Transaction = trans
            _cmd.CommandText = strSPName
            _cmd.CommandType = CommandType.StoredProcedure
            Dim par As IDataParameter = GetDataParameterObject()
            Dim i As Integer = 0
            For i = 0 To parCollection.Length - 1
                par = GetDataParameterObject()
                par.ParameterName = parCollection(i).ParameterName
                par.DbType = parCollection(i).DbType
                par.Value = parCollection(i).Value
                _cmd.Parameters.Add(par)
            Next
            Dim result As Integer = _cmd.ExecuteNonQuery()
            _cmd.Parameters.Clear()
            Return result
        Catch ex As Exception
            Throw ex
        Finally
            If (_con.State.Equals(ConnectionState.Open)) Then
                _con.Close()
            End If
        End Try
    End Function

    Public Function ExecuteStoredProcedureTableValuePrameter(strSPName As String, parCollection As IDataParameter()) As Integer
        Try
            If OpenConnection() Then

                Using cmd As New SqlCommand(strSPName)
                    cmd.Connection = _con
                    cmd.CommandType = CommandType.StoredProcedure
                    For i = 0 To parCollection.Length - 1
                        cmd.Parameters.AddWithValue(parCollection(i).ParameterName, parCollection(i).Value)
                    Next
                    Dim temp As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                    Return temp
                End Using
            End If
            Return 0
        Catch ex As Exception
            _con.Close()
            Throw ex
        Finally
            _con.Close()
        End Try

    End Function
#End Region

#Region "Other Functions"

    Public Sub CloseReader()
        If _reader Is Nothing Then
            _reader.Close()
        End If
    End Sub

#End Region
    '*** Initialize command object depending upon the provider ***/
    Public Function GetConnectionObject() As IDbConnection
        Select Case _provider
            Case EnumProvider.ODBC
                _con = New OdbcConnection()
                Exit Select
            Case EnumProvider.OLEDB
                _con = New OleDbConnection()
                Exit Select
            Case EnumProvider.SQL
                _con = New SqlConnection()
                Exit Select
            Case Else
                Return Nothing
        End Select
        Return _con
    End Function

    Public Function GetCommandObject() As IDbCommand
        Select Case _provider
            Case EnumProvider.ODBC
                Dim _cmd = New OdbcCommand()
                Return _cmd
            Case EnumProvider.OLEDB
                Dim _cmd = New OleDbCommand()
                Return _cmd
            Case EnumProvider.SQL
                Dim _cmd = New SqlCommand()
                Return _cmd
            Case Else
                Return Nothing
        End Select
    End Function

    Public Function GetDataAdapterObject() As IDbDataAdapter
        Select Case _provider
            Case EnumProvider.ODBC
                Return New OdbcDataAdapter()
            Case EnumProvider.OLEDB
                Return New OleDbDataAdapter()
            Case EnumProvider.SQL
                Return New SqlDataAdapter()
            Case Else
                Return Nothing
        End Select
    End Function

    Public Function GetDataReaderObject() As IDataReader
        Select Case _provider
            Case EnumProvider.ODBC
                Dim _reader As OdbcDataReader()
                Exit Select
            Case EnumProvider.OLEDB
                Dim _reader As OleDbDataReader()
                Exit Select
            Case EnumProvider.SQL
                Dim _reader As SqlDataReader()
                Exit Select
            Case Else
                Return Nothing
        End Select
        Return _reader
    End Function


    Public Function GetDataParameterObject() As IDbDataParameter
        Select Case _provider
            Case EnumProvider.ODBC
                _dbparameter = New OdbcParameter()
                Exit Select
            Case EnumProvider.OLEDB
                _dbparameter = New OleDbParameter()
                Exit Select
            Case EnumProvider.SQL
                _dbparameter = New SqlParameter()
                Exit Select
            Case Else
                Return Nothing
        End Select
        Return _dbparameter
    End Function

    Public Function GetParameters(ByVal paramsCount As Integer) As IDbDataParameter()
        Dim idbParams(paramsCount) As IDbDataParameter
        Dim i As Integer = 0
        Select Case _provider
            Case EnumProvider.ODBC
                For i = 0 To paramsCount
                    idbParams(i) = New OdbcParameter()
                Next
                Exit Select
            Case EnumProvider.OLEDB
                For i = 0 To paramsCount
                    idbParams(i) = New OleDbParameter()
                Next
                Exit Select
            Case EnumProvider.SQL
                For i = 0 To paramsCount
                    idbParams(i) = New SqlParameter()
                Next
                Exit Select
            Case Else
                idbParams = Nothing
                Exit Select
        End Select
        Return idbParams
    End Function

    Public Function ExecuteStoredProcedureGetDataReader(ByVal strSPName As String, ByVal parCollection() As IDataParameter) As IDataReader
        Try
            If (OpenConnection()) Then
                _cmd = GetCommandObject()
                _cmd.Connection = _con
                _cmd.CommandText = strSPName
                _cmd.CommandType = CommandType.StoredProcedure
                Dim i As Integer = 0
                Dim par As IDataParameter = GetDataParameterObject()
                Int(i = 0)
                For i = 0 To parCollection.Length - 1
                    par = GetDataParameterObject()
                    par.ParameterName = parCollection(i).ParameterName
                    par.DbType = parCollection(i).DbType
                    par.Value = parCollection(i).Value
                    _cmd.Parameters.Add(par)
                Next
                Dim dr As IDataReader
                dr = _cmd.ExecuteReader(CommandBehavior.CloseConnection)
                Return dr
            End If
            Return Nothing
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function ExecuteStoredProcedureGetDataReader(ByVal strSPName As String) As IDataReader
        Try
            If (OpenConnection()) Then
                _cmd = GetCommandObject()
                _cmd.Connection = _con
                _cmd.CommandText = strSPName
                _cmd.CommandType = CommandType.StoredProcedure

                'Dim i As Integer = 0
                'Dim par As IDataParameter = GetDataParameterObject()
                'Int(i = 0)
                'For i = 0 To parCollection.Length - 1
                '    par = GetDataParameterObject()
                '    par.ParameterName = parCollection(i).ParameterName
                '    par.DbType = parCollection(i).DbType
                '    par.Value = parCollection(i).Value
                '    _cmd.Parameters.Add(par)
                'Next

                Dim dr As IDataReader
                dr = _cmd.ExecuteReader(CommandBehavior.CloseConnection)
                Return dr
            End If
            Return Nothing
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    'Public Function ExecuteStoredProcedureGetDataSetForMySQL(ByVal strSPName As String) As DataSet
    '    Try
    '        'If (OpenConnection()) Then
    '        _cmd = New MySqlCommand()
    '        _conMySQL = New MySqlConnection()
    '        _conMySQL.ConnectionString = ConfigurationManager.ConnectionStrings("FluidSecureMySqlConnectionString").ConnectionString
    '        _conMySQL.Open()
    '        _cmd.Connection = _conMySQL
    '        _cmd.CommandText = strSPName
    '        _cmd.CommandType = CommandType.Text
    '        Dim par As IDataParameter = New MySqlParameter()
    '        'Dim i As Integer = 0
    '        'For i = 0 To parCollection.Length - 1
    '        '    par = New MySqlParameter()
    '        '    par.ParameterName = parCollection(i).ParameterName
    '        '    par.DbType = parCollection(i).DbType
    '        '    par.Value = parCollection(i).Value
    '        '    _cmd.Parameters.Add(par)
    '        'Next
    '        _adpt = New MySqlDataAdapter()
    '        _adpt.SelectCommand = _cmd
    '        Dim ds As DataSet = New DataSet()
    '        _adpt.Fill(ds)
    '        _cmd.Parameters.Clear()
    '        Return ds
    '        'End If
    '        Return Nothing
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        If (_con.State.Equals(ConnectionState.Open)) Then
    '            _con.Close()
    '        End If
    '    End Try
    'End Function


    'Added By varun to Test Conn
    Public Function GetsqlConn() As SqlConnection
        Dim cn As SqlConnection = New SqlConnection(_connectionString)
        OpenConnection()
        Return cn
    End Function
    Private Property DataProvoider() As EnumProvider
        Get
            Return _provider
        End Get
        Set(ByVal Value As EnumProvider)
            _provider = Value
        End Set
    End Property
End Class

'**** Enumerator initialized for initializing Provider ****/
Public Enum EnumProvider
    ODBC = 1
    OLEDB = 2
    ORACLE = 3
    SQL = 4
End Enum
