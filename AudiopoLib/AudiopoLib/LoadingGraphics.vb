Option Strict On
Option Explicit On
Option Infer Off
Imports System.Drawing
Imports System.Threading
Imports System.Windows.Forms

Public Class LoadingGraphics(Of T As {New, Control})
    Implements IDisposable
    Private SC As SynchronizationContext = SynchronizationContext.Current
    Private WithEvents ParentSurface As T
    Private Rect As Rectangle
    Private CirclePen, ShadowPen As Pen
    Private WithEvents OpacityTimer, SpinTimer As Timers.Timer
    Private SpinDegrees, CurrentDegree, PreviousAngle, ArcAngle As Single
    Private IsSpinning As Boolean
    Public Event Completed(Sender As Object, SurfaceControl As Object)
    Public Property Pen As Pen
        Get
            Return CirclePen
        End Get
        Set(Pen As Pen)
            CirclePen.Dispose()
            CirclePen = Pen
        End Set
    End Property
    Public Sub Spin(ByVal FramesPerSecond As Double, ByVal DegreesPerFrame As Single)
        SpinDegrees = DegreesPerFrame
        IsSpinning = True
        SpinTimer.Interval = 1000 / FramesPerSecond
        ParentSurface.BringToFront()
        ParentSurface.Show()
        SpinTimer.Start()
    End Sub
    Public Sub StopSpin()
        IsSpinning = False
        SpinTimer.Stop()
        CurrentDegree = 0
        ParentSurface.Hide()
    End Sub
    Private Sub IncrementSpin(ByVal Data As Object)
        CurrentDegree += SpinDegrees
        ParentSurface.Invalidate()
    End Sub
    Private Sub SpinTimerElapsed() Handles SpinTimer.Elapsed
        SC.Send(AddressOf IncrementSpin, Nothing)
    End Sub
    Public Property Stroke As Integer
        Get
            If Not CirclePen Is Nothing Then
                Return CInt(CirclePen.Width)
            Else
                Return 8
            End If
        End Get
        Set(value As Integer)
            Rect = New Rectangle(ParentSurface.ClientRectangle.X, ParentSurface.ClientRectangle.Y, ParentSurface.ClientRectangle.Width, ParentSurface.ClientRectangle.Height)
            Rect.Inflate(-(value + 1), -(value + 1))
            CirclePen.Width = value
            ShadowPen.Width = value + 3
        End Set
    End Property
    Public Sub New(ByRef Surface As T)
        OpacityTimer = New Timers.Timer(33.333)
        SpinTimer = New Timers.Timer
        ParentSurface = Surface
        CirclePen = New Pen(Color.GreenYellow)
        ShadowPen = New Pen(Color.FromArgb(150, CInt(CirclePen.Color.R / 2), CInt(CirclePen.Color.G / 2), CInt(CirclePen.Color.B / 2)))
        Stroke = 8
        'testbrush.SetSigmaBellShape(1)
    End Sub
    Public Sub UpdateProgress(ByVal DegreesObj As Object)
        Dim Degrees As Single = DirectCast(DegreesObj, Single)
        If Degrees > ArcAngle Then
            ArcAngle = Degrees
            ParentSurface.Invalidate()
        End If
    End Sub
    Public Sub Complete(State As Object)
        OpacityTimer.Start()
    End Sub
    Public Sub OpacityTimerTick() Handles OpacityTimer.Elapsed
        SC.Send(AddressOf ReduceOpacity, 15)
    End Sub
    Public Sub ReduceOpacity(ByVal AmountObj As Object)
        Dim Amount As Integer = DirectCast(AmountObj, Integer)
        If CirclePen.Color.A > Amount Then
            CirclePen.Color = Color.FromArgb(CirclePen.Color.A - Amount, CirclePen.Color)
            ParentSurface.Invalidate()
        Else
            OpacityTimer.Stop()
            ParentSurface.Hide()
            CirclePen.Color = Color.FromArgb(255, CirclePen.Color)
            RaiseEvent Completed(Me, ParentSurface)
        End If
    End Sub
    'Dim GP As New GraphicsPath()
    'Dim testbrush As New PathGradientBrush(GP)
    Public Sub ControlPaint(Sender As Object, e As PaintEventArgs) Handles ParentSurface.Paint
        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        'GP.Reset()
        'GP.AddArc(Rect, -90, ArcAngle)
        'e.Graphics.FillPath(testbrush, GP)
        If IsSpinning = False Then
            e.Graphics.DrawArc(CirclePen, Rect, -90, ArcAngle)
        Else
            e.Graphics.DrawArc(CirclePen, Rect, CurrentDegree, 160)
            e.Graphics.DrawArc(CirclePen, Rect, CurrentDegree + 180, 160)
        End If
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls
    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ParentSurface.Dispose()
                CirclePen.Dispose()
                ShadowPen.Dispose()
                OpacityTimer.Dispose()
                SpinTimer.Dispose()
            End If
            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
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
End Class