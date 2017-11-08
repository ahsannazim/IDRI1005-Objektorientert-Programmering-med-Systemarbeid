Option Strict On
Option Explicit On
Option Infer Off

Imports System.Threading
Imports System.Windows.Forms

Public Class ProgressReporter(Of T As {New, Control})
    Implements IDisposable
    Private ProgRes As ProgressResolution = ProgressResolution.Smooth
    Private iLast As Integer
    Private IncrementThreshold As Integer
    Private SC As SynchronizationContext = SynchronizationContext.Current
    Private LinkedGraphics As LoadingGraphics(Of T)
    Private CurrentIteration As Integer
    Private TotalDegrees As Integer = 360
    Private IncreaseCounter As Integer
    Public Property Resolution As ProgressResolution
        Get
            Return ProgRes
        End Get
        Set(ByVal Smoothness As ProgressResolution)
            ProgRes = Smoothness
            IncrementThreshold = CalculateIncrement()
        End Set
    End Property
    Private Property Current As Integer
        Get
            Return CurrentIteration
        End Get
        Set(ByVal value As Integer)
            IncreaseCounter += (value - CurrentIteration)
            CurrentIteration = value
            If IncreaseCounter >= IncrementThreshold Then
                IncreaseCounter = 0
                ReportProgress(CalculateProgress(CurrentIteration, iLast, TotalDegrees))
            ElseIf CurrentIteration >= iLast Then
                CurrentIteration = iLast
                ReportProgress(TotalDegrees)
                SC.Send(AddressOf LinkedGraphics.Complete, Nothing)
            End If
        End Set
    End Property
    Public Overloads Sub Tick()
        Current += 1
    End Sub
    Public Overloads Sub Tick(ByVal Total As Integer)
        Current = Total
    End Sub
    Private Function CalculateIncrement() As Integer
        Return CInt(Math.Floor((iLast / TotalDegrees) * ProgRes.Smoothness)) + 1
    End Function
    Public Sub New(ByRef LinkedLoadingGraphics As LoadingGraphics(Of T), ByVal Last As Integer)
        iLast = Last
        IncrementThreshold = CalculateIncrement()
        LinkedGraphics = LinkedLoadingGraphics
    End Sub
    Private Function CalculateProgress(ByVal ThisIteration As Integer, ByVal LastIteration As Integer, ByVal TotalSteps As Integer, Optional ByVal Decimals As Integer = -1) As Single
        If Decimals < 0 Then
            Return CSng((ThisIteration / LastIteration) * TotalSteps)
        Else
            Return CSng(Math.Round((ThisIteration / LastIteration) * TotalSteps, Decimals))
        End If
    End Function
    Private Sub ReportProgress(ByVal TotalProgress As Single)
        SC.Post(AddressOf LinkedGraphics.UpdateProgress, TotalProgress)
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If
            LinkedGraphics = Nothing


            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class