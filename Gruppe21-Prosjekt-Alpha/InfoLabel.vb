Option Strict On
Option Explicit On
Option Infer Off

Imports System.Threading
Public Enum Direction
    Left
    Up
    Right
    Down
End Enum
Public Class InfoLabel
    Inherits Control
    Private WithEvents BarControl As New Control
    Private SC As SynchronizationContext = SynchronizationContext.Current
    Private WithEvents PanTimer As New System.Timers.Timer(1000 \ 50)
    Private WithEvents DrawLineTimer As New System.Timers.Timer(1000 \ 50)
    Private BarBrush As New SolidBrush(Color.FromArgb(60, 60, 60))
    Private BarRect As Rectangle
    Private WithEvents TextLab As New Label
    Private varMaxWidth As Integer = 1000
    Private PanFromDirection As Direction = Direction.Left
    Private PanToSelection As Direction = Direction.Left
    Private varLineHeight As Integer = 0
    Public ReadOnly Property Label As Label
        Get
            Return TextLab
        End Get
    End Property
    Public Shadows Property Text As String
        Get
            Return TextLab.Text
        End Get
        Set(value As String)
            TextLab.Text = value
        End Set
    End Property
    Private Sub PanTimer_Tick() Handles PanTimer.Elapsed
        SC.Send(AddressOf SendTick, Nothing)
    End Sub
    Private Sub LineTimer_Tick() Handles DrawLineTimer.Elapsed
        SC.Send(AddressOf SendLineTick, Nothing)
    End Sub
    Private Sub SendLineTick(Data As Object)
        varLineHeight += 5
        If varLineHeight >= Height Then
            varLineHeight = Height
            BarRect = New Rectangle(New Point(0, Height \ 2 - varLineHeight \ 2), New Size(4, varLineHeight))
            Refresh()
            PanTimer.Start()
        Else
            BarRect = New Rectangle(New Point(BarControl.Width - 5, Height \ 2 - varLineHeight \ 2), New Size(4, varLineHeight))
            Refresh()
            DrawLineTimer.Start()
        End If
    End Sub
    Public Shadows Property ForeColor As Color
        Get
            Return TextLab.ForeColor
        End Get
        Set(value As Color)
            MyBase.ForeColor = value
            TextLab.ForeColor = value
            With BarBrush
                .Color = value
            End With
            BarControl.Invalidate()
        End Set
    End Property
    Private Sub SendTick(Data As Object)
        Select Case PanFromDirection
            Case Direction.Left
                If TextLab.Left < 8 Then
                    TextLab.Left += 50
                End If
                If TextLab.Left >= 8 Then
                    TextLab.Left = 8
                Else
                    PanTimer.Start()
                End If
            Case Direction.Right
                If TextLab.Left > 0 Then
                    TextLab.Left -= 50
                End If
                If TextLab.Left <= 0 Then
                    TextLab.Left = 0
                Else
                    PanTimer.Start()
                End If
            Case Direction.Up
            Case Direction.Down
        End Select

    End Sub
    Public Sub PanIn()
        Select Case PanFromDirection
            Case Direction.Left
                With TextLab
                    .Left = -Width
                    .Show()
                End With
            Case Direction.Up
            Case Direction.Right
                With BarControl
                    .Left = Width - .Width
                End With
                With TextLab
                    .TextAlign = ContentAlignment.MiddleRight
                    .Width = Width - 8
                    .Left = Width
                    .Show()
                End With
            Case Direction.Down
        End Select
        DrawLineTimer.Start()
    End Sub
    Public Sub New(Optional Hide As Boolean = True, Optional PanInFromDirection As Direction = Direction.Left)
        PanFromDirection = PanInFromDirection
        Select Case PanFromDirection
            Case Direction.Left
                With TextLab
                    .Parent = Me
                    .Location = New Point(8, 0)
                    .AutoSize = False
                    .Size = New Size(Width - 8, Height)
                    .ForeColor = Color.FromArgb(60, 60, 60)
                    .TextAlign = ContentAlignment.MiddleLeft
                    If Hide Then
                        .Hide()
                    End If
                End With
            Case Direction.Right
                With TextLab
                    .Parent = Me
                    .Location = New Point(50, 0)
                    .AutoSize = False
                    .Size = New Size(Width - 12, Height)
                    .ForeColor = Color.FromArgb(60, 60, 60)
                    .TextAlign = ContentAlignment.MiddleRight
                    If Hide Then
                        .Hide()
                    End If
                End With
        End Select
        With DrawLineTimer
            .AutoReset = False
        End With
        With PanTimer
            .AutoReset = False
        End With
        With BarControl
            .Parent = Me
            .Size = New Size(8, Height)
            .BringToFront()
        End With
    End Sub
    Protected Overrides Sub OnBackColorChanged(e As EventArgs)
        MyBase.OnBackColorChanged(e)
        BarControl.BackColor = BackColor
    End Sub
    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        Select Case PanFromDirection
            Case Direction.Left
                'BarRect = New Rectangle(Point.Empty, New Size(4, 0))
                TextLab.Size = New Size(Width - 8, Height)
                With BarControl
                    .Height = Height
                    .Left = 0
                End With
            Case Direction.Up
            Case Direction.Right
                With BarControl
                    .Height = Height
                    .Left = Width - .Width
                    .BringToFront()
                End With
                'BarRect = New Rectangle(New Point(BarControl.Width - 5, 0), New Size(4, 0))
                With TextLab
                    .Size = New Size(Width - 12, Height)
                End With
            Case Direction.Down
        End Select
    End Sub
    Public ReadOnly Property Brush As SolidBrush
        Get
            Return BarBrush
        End Get
    End Property
    Private Sub BarPaint(Sender As Object, e As PaintEventArgs) Handles BarControl.Paint
        With e.Graphics
            .FillRectangle(BarBrush, BarRect)
        End With
    End Sub
End Class
