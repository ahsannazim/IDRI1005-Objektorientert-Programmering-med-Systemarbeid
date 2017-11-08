Option Strict On
Option Explicit On
Option Infer Off

Public Class MySqlAdminLogin
    Implements IDisposable
    Private ServerString, DatabaseString As String
    Private boolAutoDispose As Boolean
    Private WhenCompletedAction As Action(Of Boolean)
    Private DBC As DatabaseClient
    Public Event ValidationCompleted(ByVal Valid As Boolean)
    Public Property Server As String
        Get
            Return ServerString
        End Get
        Set(value As String)
            ServerString = value
        End Set
    End Property
    Public Property Database As String
        Get
            Return DatabaseString
        End Get
        Set(value As String)
            DatabaseString = value
        End Set
    End Property
    Public WriteOnly Property WhenFinished As Action(Of Boolean)
        Set(Action As Action(Of Boolean))
            WhenCompletedAction = Action
        End Set
    End Property
    Public Property AutoDispose As Boolean
        Get
            Return boolAutoDispose
        End Get
        Set(value As Boolean)
            boolAutoDispose = value
        End Set
    End Property
    Public Sub New(ByVal Server As String, ByVal Database As String, Optional ByVal WhenFinished As Action(Of Boolean) = Nothing)
        ServerString = Server
        DatabaseString = Database
        WhenCompletedAction = WhenFinished
    End Sub
    Public Sub LoginAsync(ByVal UID As String, ByVal Password As String)
        DBC = New DatabaseClient(ServerString, DatabaseString, UID, Password)
        AddHandler DBC.ValidationCompleted, AddressOf ValidationCompelted
        DBC.ValidateConnection()
    End Sub
    Private Sub ValidationCompelted(ByVal Valid As Boolean)
        RaiseEvent ValidationCompleted(Valid)
        If WhenCompletedAction IsNot Nothing Then
            WhenCompletedAction.Invoke(Valid)
        End If
        RemoveHandler DBC.ValidationCompleted, AddressOf ValidationCompelted
        DBC.Dispose()
        If AutoDispose Then
            Dispose()
        End If
    End Sub
#Region "IDisposable Support"
    Private disposedValue As Boolean
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                If DBC IsNot Nothing Then
                    DBC.Dispose()
                End If
            End If
                WhenCompletedAction = Nothing
        End If
        disposedValue = True
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
    End Sub
#End Region
End Class