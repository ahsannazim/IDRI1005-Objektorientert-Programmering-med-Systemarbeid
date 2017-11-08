Option Strict On
Option Explicit On
Option Infer Off

Imports System.Threading

Public Class ThreadStarter
    Implements IDisposable
#Region "Fields"
    Private SC As SynchronizationContext = SynchronizationContext.Current
    Private MethodToRun As Action
    Private MethodToRunIn As Action(Of Object)
    Private FuncToRunOut As Func(Of Object)
    Private FuncToRunInOut As Func(Of Object, Object)
    Private MethodType As Integer
    Private StartTime, EndTime As Date
    Private Running As Boolean
    Private ThreadID As Object = Nothing
    Private DisposeWhenFinished As Boolean = True
    Private IsBG As Boolean = True
    Private SubWhenFinished As Action(Of Object, ThreadStarterEventArgs)
    Private NoParamMethodTask, ParamMethodTask As Task
    Private NoParamFuncThread, ParamFuncThread As Task(Of Object)
#End Region
#Region "Events"
    Public Event WorkCompleted(Result As Object, e As ThreadStarterEventArgs)
#End Region
    Public WriteOnly Property WhenFinished As Action(Of Object, ThreadStarterEventArgs)
        Set(Action As Action(Of Object, ThreadStarterEventArgs))
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
    Public Property ID As Object
        Get
            Return ThreadID
        End Get
        Set(value As Object)
            ThreadID = value
        End Set
    End Property

    Public Sub New(Method As Action, Optional ID As Object = 0)
        MethodToRun = Method
        MethodType = 1
        ThreadID = ID
    End Sub
    Public Sub New(Method As Action(Of Object), Optional ID As Object = 0)
        MethodToRunIn = Method
        MethodType = 2
        ThreadID = ID
    End Sub
    Public Sub New(Func As Func(Of Object), Optional ID As Object = 0)
        FuncToRunOut = Func
        MethodType = 3
        ThreadID = ID
    End Sub
    Public Sub New(Func As Func(Of Object, Object), Optional ID As Object = 0)
        FuncToRunInOut = Func
        MethodType = 4
        ThreadID = ID
    End Sub
    Public Function RunTime() As TimeSpan
        If Running = True Then
            Return Date.Now.Subtract(StartTime)
        Else
            Return EndTime.Subtract(StartTime)
        End If
    End Function
    Private Sub RaiseWorkCompleted(Result As Object)
        EndTime = Date.Now
        Running = False
        If Not SubWhenFinished Is Nothing Then
            SubWhenFinished.Invoke(Result, New ThreadStarterEventArgs(ThreadID, StartTime, EndTime))
        End If
        RaiseEvent WorkCompleted(Result, New ThreadStarterEventArgs(ThreadID, StartTime, EndTime))
        If DisposeWhenFinished Then
            Dispose()
        End If
    End Sub
    Public Overloads Sub Start(Optional Parameters As Object = Nothing)
        If Running = False Then
            StartTime = Date.Now
            Running = True
            If IsBG = True Then
                Select Case MethodType
                    Case 0
                        Running = False
                        Throw New Exception("Cannot start this thread before it has been assigned a method to run (use the New constructor).")
                    Case 1
                        NoParamMethodTask = Task.Factory.StartNew(AddressOf ExecuteNoParams)
                    Case 2
                        ParamMethodTask = Task.Factory.StartNew(AddressOf Execute, Parameters)
                    Case 3
                        NoParamFuncThread = Task.Factory.StartNew(AddressOf ExecuteFunction)
                    Case 4
                        ParamFuncThread = Task.Factory.StartNew(AddressOf ExecuteFunction, Parameters)
                End Select
            Else
                Select Case MethodType
                    Case 0
                        Running = False
                        Throw New Exception("Cannot start this thread before it has been assigned a method to run (use the New constructor).")
                    Case 1
                        Dim NoParamMethodTask As New Thread(AddressOf ExecuteNoParams)
                        With NoParamMethodTask
                            .IsBackground = False
                            .Start()
                        End With
                    Case 2
                            Dim ParamMethodTask As New Thread(AddressOf Execute)
                        With ParamMethodTask
                            .IsBackground = False
                            .Start(Parameters)
                        End With
                    Case 3
                        Dim test As ThreadStart = AddressOf ExecuteFunction
                        Dim NoParamFuncThread As New Thread(test)
                        With NoParamFuncThread
                            .IsBackground = False
                            .Start()
                        End With
                    Case 4
                        Dim test As ParameterizedThreadStart = AddressOf ExecuteFunction
                        Dim ParamFuncThread As New Thread(test)
                        With ParamFuncThread
                            .IsBackground = False
                            .Start(Parameters)
                        End With
                End Select
            End If
        End If
    End Sub
    Private Sub ExecuteNoParams()
        MethodToRun.Invoke
        SC.Post(AddressOf RaiseWorkCompleted, Nothing)
    End Sub
    Private Sub Execute(Parameters As Object)
        MethodToRunIn.Invoke(Parameters)
        SC.Post(AddressOf RaiseWorkCompleted, Nothing)
    End Sub
    Private Overloads Function ExecuteFunction(P As Object) As Object
        Dim Ret As Object = FuncToRunInOut.Invoke(P)
        SC.Post(AddressOf RaiseWorkCompleted, Ret)
        Return Ret
    End Function
    Private Overloads Function ExecuteFunction() As Object
        Dim Ret As Object = FuncToRunOut.Invoke
        SC.Post(AddressOf RaiseWorkCompleted, Ret)
        Return Ret
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                If NoParamMethodTask IsNot Nothing AndAlso (NoParamMethodTask.IsCanceled OrElse NoParamMethodTask.IsCompleted OrElse NoParamMethodTask.IsFaulted) Then
                    NoParamMethodTask.Dispose()
                End If
                If ParamMethodTask IsNot Nothing AndAlso (ParamMethodTask.IsCanceled OrElse ParamMethodTask.IsCompleted OrElse ParamMethodTask.IsFaulted) Then
                    ParamMethodTask.Dispose()
                End If
                If NoParamFuncThread IsNot Nothing AndAlso (NoParamFuncThread.IsCanceled OrElse NoParamFuncThread.IsCompleted OrElse NoParamFuncThread.IsFaulted) Then
                    NoParamFuncThread.Dispose()
                End If
                If ParamFuncThread IsNot Nothing AndAlso (ParamFuncThread.IsCanceled OrElse ParamFuncThread.IsCompleted OrElse ParamFuncThread.IsFaulted) Then
                    ParamFuncThread.Dispose()
                End If
            End If
            MethodToRun = Nothing
            MethodToRunIn = Nothing
            FuncToRunOut = Nothing
            FuncToRunInOut = Nothing
            SubWhenFinished = Nothing

            ThreadID = Nothing
            StartTime = Nothing
            EndTime = Nothing
            MethodType = 0
            Running = False
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
