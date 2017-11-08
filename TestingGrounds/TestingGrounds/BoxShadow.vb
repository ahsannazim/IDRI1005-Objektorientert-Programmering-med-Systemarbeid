Option Strict On
Option Explicit On
Option Infer Off

Public Class BoxShadow
    Inherits PictureBox
    Private WithEvents TargetControl As Control
    Private ShadWidth As Integer
    Private ShadColor As Color
    Private ShadIntensity As Double = 0.4
    Public Sub New(Target As Control, Optional Width As Integer = 4, Optional Intensity As Double = 0.5)
        ShadWidth = Width
        ShadColor = Color.Black
        ShadIntensity = Intensity
        Me.Target = Target
    End Sub
    Public Sub New(Target As Control, Color As Color, Optional Width As Integer = 5, Optional Intensity As Double = 0.5)
        ShadColor = Color
        ShadIntensity = Intensity
        ShadWidth = Width
        Me.Target = Target
    End Sub
    Private Sub RepositionShadow(Sender As Object, e As EventArgs) Handles TargetControl.LocationChanged, TargetControl.SizeChanged
        Dim SenderX As Control = DirectCast(Sender, Control)
        With Me
            .SuspendLayout()
            .Left = SenderX.Left - ShadWidth - 1
            .Top = SenderX.Top - ShadWidth - 1
            .Width = SenderX.Width + ShadWidth * 2 + 1
            .Height = SenderX.Height + ShadWidth * 2 + 1
            .ResumeLayout()
        End With
    End Sub
    Public Property Target As Control
        Get
            Return TargetControl
        End Get
        Set(value As Control)
            TargetControl = value
            With Me
                .Hide()
                .SuspendLayout()
                .Parent = value.Parent
                .Left = value.Left - ShadWidth - 1
                .Top = value.Top - ShadWidth - 1
                .Width = value.Width + ShadWidth * 2 + 1
                .Height = value.Height + ShadWidth * 2 + 1
                .SendToBack()
                .ResumeLayout()
                .Show()
            End With
        End Set
    End Property
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        Dim Rect As New Rectangle(DisplayRectangle.Location, DisplayRectangle.Size)
        'Dim StepInt As Integer = ((255 * ShadIntensity) \ ShadWidth) - 2
        For i As Integer = 0 To ShadWidth
            Rect.Inflate(-1, -1)
            Dim Alpha As Integer = CInt(2 + 10 * i ^ (1 + ShadIntensity))
            If Alpha > 255 Then
                Alpha = 255
            End If
            Dim ShadPen As New Pen(Color.FromArgb(Alpha, ShadColor), 1)
            e.Graphics.DrawRectangle(ShadPen, Rect)
            ShadPen.Dispose()
        Next
    End Sub
End Class
