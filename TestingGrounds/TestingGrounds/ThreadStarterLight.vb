Option Strict On
Option Explicit On
Option Infer Off

Imports System.Threading

Public NotInheritable Class ThreadStarterLight
    Implements IDisposable
#Region "Fields"
    Private SC As SynchronizationContext = SynchronizationContext.Current
    Private MethodToRun As Action(Of Object)
    Private FuncToRun As Func(Of Object, Object)
    Private MethodType As Integer
    Private Running As Boolean
    Private DisposeWhenFinished As Boolean = True
    Private IsBG As Boolean = True
    Private SubWhenFinished As Action(Of Object)
    Private SubMethod As Task, FuncMethod As Task(Of Object)
#End Region
#Region "Events"
    Public Event WorkCompleted(Result As Object)
#End Region
    Public WriteOnly Property WhenFinished As Action(Of Object)
        Set(Action As Action(Of Object))
            SubWhenFinished = Action
        End Set
    End Property
    Public Property IsBackground As Boolean
        Get
            Return IsBG
        End Get
        Set(value As Boolean)
            IsBG = value
        End Set
    End Property
    Public ReadOnly Property IsRunning As Boolean
        Get
            Return Running
        End Get
    End Property
    Public Sub New(Method As Action(Of Object))
        MethodToRun = Method
        MethodType = 1
    End Sub
    Public Sub New(Func As Func(Of Object, Object))
        FuncToRun = Func
        MethodType = 2
    End Sub
    Private Sub RaiseWorkCompleted(Result As Object)
        Running = False
        If Not SubWhenFinished Is Nothing Then
            SubWhenFinished.Invoke(Result)
        End If
        RaiseEvent WorkCompleted(Result)
        If DisposeWhenFinished Then
            Dispose()
        End If
    End Sub
    Public Overloads Sub Start(Optional Parameters As Object = Nothing)
        If Running = False Then
            Running = True
            If IsBG Then
                Select Case MethodType
                    Case 0
                        Running = False
                        Throw New Exception("Cannot start this thread before it has been assigned a method to run (use the New constructor).")
                    Case 1
                        SubMethod = Task.Factory.StartNew(AddressOf Execute, Parameters)
                    Case 2
                        FuncMethod = Task.Factory.StartNew(AddressOf ExecuteFunction, Parameters)
                End Select
            Else
                Select Case MethodType
                    Case 0
                        Running = False
                        Throw New Exception("Cannot start this thread before it has been assigned a method to run (use the New constructor).")
                    Case 1
                        Dim SubMethodThread As New Thread(AddressOf Execute)
                        SubMethodThread.IsBackground = False
                        SubMethodThread.Start(Parameters)
                    Case 2
                        Dim test As ParameterizedThreadStart = AddressOf ExecuteFunction
                        Dim FuncMethodThread As New Thread(test)
                        FuncMethodThread.IsBackground = False
                        FuncMethodThread.Start(Parameters)
                End Select
            End If
        End If
    End Sub
    Private Sub Execute(Parameters As Object)
        MethodToRun.Invoke(Parameters)
        SC.Post(AddressOf RaiseWorkCompleted, Nothing)
    End Sub
    Private Overloads Function ExecuteFunction(P As Object) As Object
        Dim Ret As Object = FuncToRun.Invoke(P)
        SC.Post(AddressOf RaiseWorkCompleted, Ret)
        Return Ret
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                SubMethod.Dispose()
                FuncMethod.Dispose()
            End If
            MethodToRun = Nothing
            FuncToRun = Nothing
            SubWhenFinished = Nothing
            MethodType = 0
            Running = False
            SC = Nothing
        End If
        disposedValue = True
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
    End Sub
#End Region
End Class

Public Class ThreadStarterEventArgs
    Inherits EventArgs
    Private STime As Date
    Private ETime As Date
    Private ThreadID As Object
    Public ReadOnly Property ID As Object
        Get
            Return ThreadID
        End Get
    End Property
    Public ReadOnly Property StartTime As Date
        Get
            Return STime
        End Get
    End Property
    Public ReadOnly Property EndTime As Date
        Get
            Return ETime
        End Get
    End Property
    Public Function TimeToFinish() As TimeSpan
        Return ETime.Subtract(STime)
    End Function
    Public Sub New(ByVal ID As Object, StartTime As Date, EndTime As Date)
        ThreadID = ID
        STime = StartTime
        ETime = EndTime
    End Sub
End Class
