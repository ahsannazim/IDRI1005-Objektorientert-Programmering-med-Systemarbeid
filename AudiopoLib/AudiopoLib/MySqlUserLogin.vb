Option Strict On
Option Explicit On
Option Infer Off

Public Class MySqlUserLogin
    Private WithEvents DBC As DatabaseClient
    Private IfValidAction As Action
    Private InIfInvalidAction As Action(Of Boolean, String)
    Private ExecuteIfMultipleMatches As Boolean = True
    Public Event MultipleFound()
    Public Property ExecuteIfMultiple As Boolean
        Get
            Return ExecuteIfMultipleMatches
        End Get
        Set(value As Boolean)
            ExecuteIfMultipleMatches = value
        End Set
    End Property
    Public Sub New(ByVal Server As String, ByVal Database As String, ByVal UserID As String, ByVal Password As String)
        DBC = New DatabaseClient(Server, Database, UserID, Password)
    End Sub
    Public Overloads Sub LoginAsync(ByVal Username As String, ByVal Password As String, ByVal TableName As String, ByVal UidColumnName As String, ByVal PwdColumnName As String)
        ' Parameteriser UidColumnName og TableName også
        DBC.SQLQuery = "SELECT " & UidColumnName & " FROM " & TableName & " WHERE " & UidColumnName & "=@username AND " & PwdColumnName & "=@password"
        DBC.Execute({"@username", "@password"}, {Username, Password}, True)
    End Sub
    Public Overloads Sub LoginAsync(ByVal Username As String, ByVal Password As String, ByVal TableName As String, ByVal UidColumnName As String, ByVal PwdColumnName As String, ByVal Limit As Integer)
        ' Parameteriser UidColumnName og TableName også
        DBC.SQLQuery = "SELECT " & UidColumnName & " FROM " & TableName & " WHERE " & UidColumnName & "=@username AND " & PwdColumnName & "=@password LIMIT " & CStr(Limit)
        DBC.Execute({"@username", "@password"}, {Username, Password}, True)
    End Sub
    Public Overloads Sub LoginAsync(ByVal Username As String, ByVal Password As String, ByVal Query As String, ByVal Parameters() As String, ByVal Values() As String)
        DBC.SQLQuery = Query
        DBC.Execute(Parameters, Values, True)
    End Sub
    Private Sub Fin(ByVal Exists As Boolean, ByVal RowCount As Integer, Tag As Integer, ByVal ErrorOccurred As Boolean, ErrorMessage As String) Handles DBC.ExistsCheckCompleted
        If Exists Then
            If RowCount = 1 OrElse ExecuteIfMultipleMatches Then
                If IfValidAction IsNot Nothing Then
                    IfValidAction.Invoke
                End If
            Else
                RaiseEvent MultipleFound()
            End If
        Else
            If InIfInvalidAction IsNot Nothing Then
                InIfInvalidAction.Invoke(ErrorOccurred, ErrorMessage)
            End If
        End If
    End Sub
    Public Overloads WriteOnly Property IfValid As Action
        Set(Action As Action)
            IfValidAction = Action
        End Set
    End Property
    ''' <summary>
    ''' Address of a method that accepts a boolean value indicating whether or not an error occurred.
    ''' </summary>
    Public Overloads WriteOnly Property IfInvalid As Action(Of Boolean, String)
        Set(Action As Action(Of Boolean, String))
            InIfInvalidAction = Action
        End Set
    End Property
End Class