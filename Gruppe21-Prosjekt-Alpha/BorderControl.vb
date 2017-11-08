Imports System.Drawing.Drawing2D

Public Class BorderControl
    Inherits Control
    Private varBorderPen As Pen
    Private varDashed As Boolean
    Private varDrawBorders() As Boolean = {True, True, True, True}
    Private DashFG, DashBG As Color
    Public Property DrawBorder(ByVal Side As AudiopoLib.FormField.ElementSide) As Boolean
        Get
            Return varDrawBorders(CInt(Side))
        End Get
        Set(value As Boolean)
            varDrawBorders(CInt(Side)) = value
            Invalidate()
        End Set
    End Property
    Public Sub MakeDashed(ForeColor As Color, BackColor As Color)
        DashFG = ForeColor
        DashBG = BackColor
        varDashed = True
        Invalidate()
    End Sub
    Public Sub MakeSolid()
        varDashed = False
    End Sub
    Public Sub New(BorderColor As Color)
        DoubleBuffered = True
        ResizeRedraw = True
        varBorderPen = New Pen(BorderColor)
    End Sub
    Public Property BorderPen As Pen
        Get
            Return varBorderPen
        End Get
        Set(value As Pen)
            varBorderPen.Dispose()
            varBorderPen = value
        End Set
    End Property
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        Dim Height As Integer = ClientSize.Height
        Dim Width As Integer = ClientSize.Width
        With e.Graphics
            If Not varDashed Then
                If varDrawBorders(0) Then
                    .DrawLine(varBorderPen, Point.Empty, New Point(0, Height - 1))
                End If
                If varDrawBorders(1) Then
                    .DrawLine(varBorderPen, Point.Empty, New Point(Width - 1, 0))
                End If
                If varDrawBorders(2) Then
                    .DrawLine(varBorderPen, New Point(Width - 1, 0), New Point(Width - 1, Height - 1))
                End If
                If varDrawBorders(3) Then
                    .DrawLine(varBorderPen, New Point(0, Height - 1), New Point(Width - 1, Height - 1))
                End If
            Else
                Using HB As New HatchBrush(HatchStyle.LargeCheckerBoard, DashFG, DashBG)
                    If varDrawBorders(0) Then
                        Using P As New Pen(HB)
                            .DrawLine(P, Point.Empty, New Point(0, Height - 1))
                        End Using
                    End If
                    If varDrawBorders(1) Then
                        Using P As New Pen(HB)
                            .DrawLine(P, Point.Empty, New Point(Width - 1, 0))
                        End Using
                    End If
                    If varDrawBorders(2) Then
                        Using P As New Pen(HB)
                            .DrawLine(P, New Point(Width - 1, 0), New Point(Width - 1, Height - 1))
                        End Using
                    End If
                    If varDrawBorders(3) Then
                        Using P As New Pen(HB)
                            .DrawLine(P, New Point(0, Height - 1), New Point(Width - 1, Height - 1))
                        End Using
                    End If
                End Using
            End If
        End With
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            varBorderPen.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class
