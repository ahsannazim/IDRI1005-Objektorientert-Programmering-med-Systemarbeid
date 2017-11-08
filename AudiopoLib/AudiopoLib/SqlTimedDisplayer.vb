Option Strict On
Option Explicit On
Option Infer Off

Imports System.Threading
Imports System.Windows.Forms

Public Class SqlTimedDisplayer
    Implements IDisposable

    Private SC As SynchronizationContext = SynchronizationContext.Current
    Private WithEvents DBC As DatabaseClient
    Private WithEvents NotifManager As NotificationManager
    Private Q As String
    Private TickDelay As Double
    Private CurrentInt As Integer
    Private RandomMinMax(1) As Integer
    Private IncrementStep As Integer = 1
    Private WithEvents DelayTimer As Timers.Timer
    Private ParentControl As Control
    Private ParamArr() As String
    Private Rnd As Random
    Private DisplayerDuration As Double
    Private NotifAppearance As NotificationAppearance
    Private IsRandom As Boolean
    Private IsFinished As Boolean
    Private DoRepeat As Boolean
    Private CounterMax As Integer
    Private Previous As Integer
    Private FloatX As FloatX
    Private FloatY As FloatY
    Public Event ConnectionErrorOccurred()
    Public Sub New(ByVal Server As String, ByVal Database As String, ByVal UID As String, ByVal Password As String, ByVal Parent As Control, ByVal Query As String, ByVal Parameter As String, ByVal RandomMin As Integer, ByVal RandomMax As Integer, ByVal Appearance As NotificationAppearance, Optional ByVal Duration As Double = 10, Optional ByVal Delay As Double = 4, Optional ByVal AlignmentX As FloatX = FloatX.Left, Optional ByVal AlignmentY As FloatY = FloatY.Top)
        IsRandom = True
        FloatX = AlignmentX
        FloatY = AlignmentY
        ParentControl = Parent
        DisplayerDuration = Duration
        TickDelay = Delay
        NotifAppearance = Appearance
        Rnd = New Random
        DBC = New DatabaseClient(Server, Database, UID, Password)
        NotifManager = New NotificationManager(Parent)
        DelayTimer = New Timers.Timer(TickDelay)
        RandomMinMax(0) = RandomMin
        RandomMinMax(1) = RandomMax
        DBC.SQLQuery = Query
        ParamArr = {Parameter}
    End Sub
    Public Sub New(ByVal Server As String, ByVal Database As String, ByVal UID As String, ByVal Password As String, ByVal Parent As Control, ByVal Query As String, ByVal Parameter As String, ByVal Delay As Double, ByVal Duration As Double, ByVal Appearance As NotificationAppearance, ByVal Maximum As Integer, Optional ByVal Repeat As Boolean = True, Optional ByVal Increment As Integer = 1)
        IsRandom = False
        DoRepeat = Repeat
        CounterMax = Maximum
        ParentControl = Parent
        DisplayerDuration = Duration
        TickDelay = Delay
        NotifAppearance = Appearance
        DBC = New DatabaseClient(Server, Database, UID, Password)
        NotifManager = New NotificationManager(Parent)
        DelayTimer = New Timers.Timer(TickDelay)
        DBC.SQLQuery = Query
        ParamArr = {Parameter}
    End Sub
    Private Sub Fetch(State As Object)
        DelayTimer.Stop()
        If Not IsFinished Then
            If IsRandom = True Then
                Dim ValArr() As String = {CStr(Rnd.Next(RandomMinMax(0), RandomMinMax(1) + 1))}
                If RandomMinMax(1) + 1 - RandomMinMax(0) > 1 Then
                    Do While CInt(ValArr(0)) = Previous
                        ValArr = {CStr(Rnd.Next(RandomMinMax(0), RandomMinMax(1) + 1))}
                    Loop
                End If
                Previous = CInt(ValArr(0))
                DBC.Execute(ParamArr, ValArr)
            Else
                CurrentInt += IncrementStep
                If CurrentInt <= CounterMax Then
                    Dim ValArr() As String = {CStr(CurrentInt)}
                    DBC.Execute(ParamArr, ValArr)
                ElseIf DoRepeat Then
                    CurrentInt = 0
                    Dim ValArr() As String = {CStr(CurrentInt)}
                    DBC.Execute(ParamArr, ValArr)
                Else
                    Finish()
                End If
            End If
        End If
    End Sub
    Private Sub FetchFinished(Sender As Object, e As DatabaseListEventArgs) Handles DBC.ListLoaded
        With e
            If Not IsFinished Then
                If Not .ErrorOccurred Then
                    Dim Result As String = CStr(.Data.Rows.Item(0)(0))
                    NotifManager.Display(Result, NotifAppearance, DisplayerDuration, FloatX, FloatY)
                Else
                    Finish()
                    RaiseEvent ConnectionErrorOccurred()
                End If
            End If
        End With
    End Sub
    Private Sub DisplayFinished(sender As Notification) Handles NotifManager.NotificationClosed
        If Not IsFinished Then
            DelayTimer.Start()
        Else
            Dispose()
        End If
    End Sub
    Public Sub Start()
        If Not IsFinished Then
            DelayTimer.Start()
        End If
    End Sub
    Public Sub Finish()
        IsFinished = True
    End Sub
    Private Sub DelayTimer_Elapsed() Handles DelayTimer.Elapsed
        SC.Post(AddressOf Fetch, Nothing)
    End Sub
#Region "IDisposable Support"
    Private disposedValue As Boolean
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                DBC.Dispose()
                NotifManager.Dispose()
                DelayTimer.Dispose()
            End If
            ParentControl = Nothing
            Rnd = Nothing
        End If
        disposedValue = True
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
    End Sub
#End Region
End Class
