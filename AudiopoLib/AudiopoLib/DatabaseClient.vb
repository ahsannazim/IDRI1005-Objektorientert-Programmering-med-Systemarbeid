Option Strict On
Option Explicit On
Option Infer Off

Imports MySql.Data.MySqlClient
Public Class DatabaseClient
    Implements IDisposable
    Private ServerString As String
    Private DatabaseString As String
    Private UIDString As String
    Private PasswordString As String
    Private LatestDT As DataTable
    Private SQLq As String = ""
    Private Tag As Integer = 0
    Private HandledThread As ThreadStarter
    Public Event ListLoaded(Sender As Object, e As DatabaseListEventArgs)
    Public Event ExecutionFailed(ByVal ClientTag As Integer)
    Public Event ExistsCheckCompleted(ByVal Exists As Boolean, ByVal RowCount As Integer, ByVal Tag As Integer, ByVal ErrorOccurred As Boolean, ErrorMessage As String)
    Public Event ValidationCompleted(ByVal Success As Boolean)
    Private DBCSB As MySqlConnectionStringBuilder
    Public Property UID As String
        Get
            Return UIDString
        End Get
        Set(value As String)
            UIDString = value
        End Set
    End Property
    Public Property SQLQuery As String
        Get
            Return SQLq
        End Get
        Set(query As String)
            SQLq = query
        End Set
    End Property
    Public WriteOnly Property Password As String
        Set(Value As String)
            DBCSB.Password = Value
        End Set
    End Property
    Public Sub New(ByVal Server As String, ByVal Database As String, ByVal UID As String, ByVal Password As String)
        DBCSB = New MySqlConnectionStringBuilder
        DBCSB.Server = Server
        DBCSB.Database = Database
        DBCSB.UserID = UID
        DBCSB.Password = Password
    End Sub
    Public Sub ValidateConnection()
        HandledThread = New ThreadStarter(AddressOf AsyncCheckConnection)
        HandledThread.WhenFinished = AddressOf ValidateFinished
        HandledThread.Start(Me.Connection.ConnectionString)
    End Sub
    Private Function AsyncCheckConnection(Data As Object) As Boolean
        Dim CS As String = DirectCast(Data, String)
        Dim DBConnection As New MySqlConnection(CS)
        Dim Valid As Boolean = False
        With DBConnection
            Try
                .Open()
                If .State = ConnectionState.Open Then
                    Valid = True
                Else
                    Valid = False
                End If
            Catch
                Valid = False
            Finally
                .Close()
                .Dispose()
            End Try
            Return Valid
        End With
    End Function
    Private Sub ValidateFinished(Result As Object, e As ThreadStarterEventArgs)
        Dim Valid As Boolean = DirectCast(Result, Boolean)
        RaiseEvent ValidationCompleted(Valid)
    End Sub
    Public Overloads Sub Execute(Optional CheckIfExistsOnly As Boolean = False, Optional Tag As Integer = 0)
        Dim RegularQueryInstance As New RegularQuery(DBCSB.ConnectionString, SQLq, Tag)
        HandledThread = New ThreadStarter(AddressOf ExecuteQuery)
        If CheckIfExistsOnly = True Then
            RegularQueryInstance.CheckIfExists = True
            HandledThread.WhenFinished = AddressOf ExistCheck
        Else
            HandledThread.WhenFinished = AddressOf GetListFinished
        End If
        HandledThread.Start(RegularQueryInstance)
    End Sub
    Public Overloads Sub Execute(Parameters() As String, Values() As String, Optional CheckIfExistsOnly As Boolean = False, Optional Tag As Integer = 0)
        Try
            If Parameters.Length <> Values.Length Then
                Throw New Exception("Parameters() and Values() arguments must be of equal length.")
            Else
                Dim DBParams As New ParameterizedQuery(DBCSB.GetConnectionString(True), SQLq, Tag)
                For i As Integer = 0 To Parameters.Length - 1
                    DBParams.AddPair(Parameters(i), Values(i))
                Next
                HandledThread = New ThreadStarter(AddressOf ExecuteParamQuery, Tag)
                If CheckIfExistsOnly = True Then
                    DBParams.CheckIfExists = True
                    HandledThread.WhenFinished = AddressOf ExistCheck
                Else
                    HandledThread.WhenFinished = AddressOf GetListFinished
                End If
                HandledThread.Start(DBParams)
            End If
        Catch ex As Exception
            Throw ex
            RaiseEvent ExecutionFailed(Me.Tag)
        End Try
    End Sub
    Private Sub ExistCheck(Result As Object, e As ThreadStarterEventArgs)
        Dim ResultArr() As Object = DirectCast(Result, Object())
        RaiseEvent ExistsCheckCompleted(DirectCast(ResultArr(0), Boolean), DirectCast(ResultArr(1), Integer), DirectCast(ResultArr(2), Integer), DirectCast(ResultArr(3), Boolean), DirectCast(ResultArr(4), String))
    End Sub
    ' FIKS
    Private Function ExecuteParamQuery(Params As Object) As Object
        Dim ParamInstance As ParameterizedQuery = DirectCast(Params, ParameterizedQuery)
        Dim DBConnection As New MySqlConnection(ParamInstance.ConnectionString)
        Dim ErrorOccurred As Boolean
        Dim ErrorMessage As String = "Nothing to show"
        Dim Ret(2) As Object
        Dim RetTable As New DataTable
        Try
            DBConnection.Open()
            Dim SQLcmd As New MySqlCommand(ParamInstance.Query, DBConnection)
            Dim iLast As Integer = ParamInstance.Count - 1
            With SQLcmd.Parameters
                For i As Integer = 0 To iLast
                    .AddWithValue(ParamInstance.Pair(i)(0), ParamInstance.Pair(i)(1))
                Next
            End With
            Dim SqlDA As New MySqlDataAdapter(SQLcmd)
            With SqlDA
                .Fill(RetTable)
                .Dispose()
            End With
            SQLcmd.Dispose()
        Catch ex As Exception
            ErrorOccurred = True
            ErrorMessage = ex.Message
        Finally
            With DBConnection
                .Close()
                .Dispose()
            End With
        End Try
        If Not ParamInstance.CheckIfExists Then
            Ret(0) = RetTable
            Ret(1) = ErrorOccurred
            Ret(2) = ErrorMessage
            Return Ret
        Else
                Dim Result(4) As Object
            Dim RetCount As Integer
            With RetTable
                RetCount = .Rows.Count
                If RetCount > 0 Then
                    Result(0) = True
                Else
                    Result(0) = False
                End If
                Result(1) = RetCount
                Result(2) = ParamInstance.Tag
                Result(3) = ErrorOccurred
                Result(4) = ErrorMessage
                .Dispose()
            End With
            Return Result
        End If
    End Function
    Private Function ExecuteQuery(Params As Object) As Object
        Dim QueryInstance As RegularQuery = DirectCast(Params, RegularQuery)
        Dim DBConnection As New MySqlConnection(QueryInstance.ConnectionString)
        Dim Ret(2) As Object
        Dim RetTable As New DataTable
        Dim ErrorOccurred As Boolean = False
        Dim ErrorMessage As String = ""
        Try
            DBConnection.Open()
            Dim SQLcmd As New MySqlCommand(QueryInstance.Query, DBConnection)
            Dim SqlDA As New MySqlDataAdapter
            With SqlDA
                .SelectCommand = SQLcmd
                .Fill(RetTable)
                SQLcmd.Dispose()
                .Dispose()
            End With
        Catch ex As Exception
            ErrorOccurred = True
            ErrorMessage = ex.Message
        Finally
            With DBConnection
                .Close()
                .Dispose()
            End With
        End Try
        Ret(0) = RetTable
        Ret(1) = ErrorOccurred
        Ret(2) = ErrorMessage
        Return Ret
    End Function
    Public ReadOnly Property Connection As MySqlConnectionStringBuilder
        Get
            Return DBCSB
        End Get
    End Property
    Private Sub GetListFinished(State As Object, e As ThreadStarterEventArgs)
        Dim RetArr() As Object = DirectCast(State, Object())
        Dim DT As DataTable = DirectCast(RetArr(0), DataTable)
        Dim ErrorOccurred As Boolean = DirectCast(RetArr(1), Boolean)
        Dim ErrorMessage As String = DirectCast(RetArr(2), String)
        If DT IsNot Nothing Then
            RaiseEvent ListLoaded(Me, New DatabaseListEventArgs(DT, ErrorOccurred, ErrorMessage))
        Else
            RaiseEvent ExecutionFailed(Me.Tag)
        End If
    End Sub
#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ServerString = ""
                DatabaseString = ""
                UIDString = ""
                PasswordString = ""
                SQLq = ""
                If Not LatestDT Is Nothing Then
                    LatestDT.Clear()
                    LatestDT.Dispose()
                End If
                If Not HandledThread Is Nothing Then
                    HandledThread.Dispose()
                End If
            End If
            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
    Private Class RegularQuery
        Private Q As String
        Private TagInt As Integer = 0
        Private ConString As String
        Private ReturnBoolean As Boolean = False
        Public ReadOnly Property Tag As Integer
            Get
                Return TagInt
            End Get
        End Property
        Public Property CheckIfExists As Boolean
            Get
                Return ReturnBoolean
            End Get
            Set(value As Boolean)
                ReturnBoolean = value
            End Set
        End Property
        Public Sub New(ByVal ConnectionString As String, ByVal Query As String, ByVal Tag As Integer)
            TagInt = Tag
            Q = Query
            ConString = ConnectionString
        End Sub
        Public ReadOnly Property Query As String
            Get
                Return Q
            End Get
        End Property
        Public ReadOnly Property ConnectionString As String
            Get
                Return ConString
            End Get
        End Property
    End Class
    Private Class ParameterizedQuery
        Private Q As String
        Private ParameterPairs As List(Of String())
        Private ConString As String
        Private TagInt As Integer = 0
        Private ReturnBoolean As Boolean = False
        Public ReadOnly Property Tag As Integer
            Get
                Return TagInt
            End Get
        End Property
        Public Property CheckIfExists As Boolean
            Get
                Return ReturnBoolean
            End Get
            Set(value As Boolean)
                ReturnBoolean = value
            End Set
        End Property
        Public ReadOnly Property Count As Integer
            Get
                Return ParameterPairs.Count
            End Get
        End Property
        Public ReadOnly Property Pair(ByVal index As Integer) As String()
            Get
                Return ParameterPairs(index)
            End Get
        End Property
        Public Sub New(ByVal ConnectionString As String, ByVal Query As String, ByVal Tag As Integer)
            Q = Query
            TagInt = Tag
            ParameterPairs = New List(Of String())
            ConString = ConnectionString
        End Sub
        Public ReadOnly Property Query As String
            Get
                Return Q
            End Get
        End Property
        Public ReadOnly Property ConnectionString As String
            Get
                Return ConString
            End Get
        End Property
        Public Sub AddPair(ByVal ParameterName As String, ByVal ParameterValue As String)
            ParameterPairs.Add({ParameterName, ParameterValue})
        End Sub
    End Class
End Class