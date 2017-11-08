Option Strict On
Option Explicit On
Option Infer Off
Imports System.Drawing
Imports System.Threading
Imports System.Windows.Forms

Public Class FormLayoutTools
    Implements IDisposable

    Private SC As SynchronizationContext = SynchronizationContext.Current
    Private Reserved As Integer
    Private ReservedGoal As Integer
    Private ParentForm As Control
    Private IncludeTitle As Boolean = False
    Private ReservedTimer As Timers.Timer
    Public Event Refreshed()
    Public Sub New(ByRef TargetParent As Control)
        ReservedTimer = New Timers.Timer(1000 / 60)
        ParentForm = TargetParent
        AddHandler ReservedTimer.Elapsed, AddressOf SlideReservation
    End Sub
    Public Property IncludeFormTitle As Boolean
        Get
            Return IncludeTitle
        End Get
        Set(value As Boolean)
            IncludeTitle = value
            SlideToDefault(True)
        End Set
    End Property
    Public Sub Refresh()
        RaiseEvent Refreshed()
    End Sub
    Public Sub SlideToDefault(Optional Immediate As Boolean = False)
        If IncludeFormTitle = True Then
            If Immediate = True Then
                ReservedSpaceTop = -GetTitleHeight()
            Else
                SlideToHeight(-GetTitleHeight())
            End If
        Else
            If Immediate = True Then
                ReservedSpaceTop = 0
            Else
                SlideToHeight(0)
            End If
        End If
    End Sub
    Private Function GetTitleHeight() As Integer
        Dim TitlebarHeight As Integer
        Dim BorderWidth As Integer
        With ParentForm
            BorderWidth = (.Width - .ClientSize.Width) \ 2
            TitlebarHeight = .Height - .ClientSize.Height - 2 * BorderWidth
        End With
        Return TitlebarHeight
    End Function
    Private Sub SlideReservation(sender As Object, e As Timers.ElapsedEventArgs)
        SC.Send(AddressOf SlideReservationIncrement, Nothing)
    End Sub
    Private Sub SlideReservationIncrement(State As Object)
        If Reserved > ReservedGoal Then
            Reserved -= 5
            If Reserved < ReservedGoal Then
                Reserved = ReservedGoal
            End If
        ElseIf Reserved < ReservedGoal Then
            Reserved += 5
            If Reserved > ReservedGoal Then
                Reserved = ReservedGoal
            End If
        Else
            ReservedTimer.Stop()
        End If
        Refresh()
    End Sub
    Public Sub SlideToHeight(ByVal value As Integer)
        ReservedGoal = value
        ReservedTimer.Start()
    End Sub
    Public Property ReservedSpaceTop() As Integer
        Get
            Return Reserved
        End Get
        Set(Pixels As Integer)
            If Pixels <> Reserved Then
                Reserved = Pixels
                Refresh()
            End If
        End Set
    End Property
    Public Overloads Sub CenterSurfaceH(Target As Control, AlignWith As Control, Optional ByVal OffsetX As Integer = 0)
        Dim AlignWithX As Integer
        If Target.Parent Is AlignWith Then
            AlignWithX = 0
        Else
            AlignWithX = AlignWith.Left
        End If
        With Target
            .Left = AlignWithX + AlignWith.Width \ 2 - .Width \ 2 + OffsetX
        End With
    End Sub
    Public Overloads Sub CenterSurfaceH(Target As Control, AlignWith As Rectangle, Optional ByVal OffsetX As Integer = 0)
        With Target
            .Left = AlignWith.Left + AlignWith.Width \ 2 - .Width \ 2 + OffsetX
        End With
    End Sub
    Public Overloads Sub CenterSurfaceV(Target As Control, AlignWith As Control, Optional ByVal OffsetY As Integer = 0)
        Dim AlignWithY As Integer
        If ReferenceEquals(ParentForm, AlignWith) Then
            AlignWithY = AlignWith.Top
        End If
        With Target
            .Top = AlignWithY + AlignWith.Height \ 2 - .Height \ 2 + OffsetY
        End With
    End Sub
    Public Overloads Sub CenterSurfaceV(Target As Control, AlignWith As Rectangle, Optional ByVal OffsetY As Integer = 0)
        With Target
            .Top = AlignWith.Top + AlignWith.Height \ 2 - .Height \ 2 + OffsetY
        End With
    End Sub
    Public Overloads Sub CenterSurface(Target As Control, AlignWith As Control, Optional ByVal OffsetX As Integer = 0, Optional ByVal OffsetY As Integer = 0)
        Dim AlignWithXY As Point
        If Not ReferenceEquals(ParentForm, AlignWith) Then
            AlignWithXY = AlignWith.Location
        End If
        With Target
            .Top = AlignWithXY.Y + AlignWith.Height \ 2 - .Height \ 2 + OffsetY
            .Left = AlignWithXY.X + AlignWith.Width \ 2 - .Width \ 2 + OffsetX
        End With
        'CenterSurfaceH(Target, AlignWith, OffsetX)
        'CenterSurfaceV(Target, AlignWith, OffsetY)
    End Sub
    Public Overloads Sub CenterSurface(Target As Control, AlignWith As Rectangle, Optional ByVal OffsetX As Integer = 0, Optional ByVal OffsetY As Integer = 0)
        'CenterSurfaceH(Target, AlignWith, OffsetX)
        'CenterSurfaceV(Target, AlignWith, OffsetY)
        With Target
            .Top = AlignWith.Top + AlignWith.Height \ 2 - .Height \ 2 + OffsetY
            .Left = AlignWith.Left + AlignWith.Width \ 2 - .Width \ 2 + OffsetX
        End With

    End Sub
    Public Sub CenterOnForm(Target As Control, Optional ByVal OffsetX As Integer = 0, Optional ByVal OffsetY As Integer = 0)
        With Target
            .Top = Reserved + (ParentForm.ClientSize.Height - Reserved) \ 2 - .Height \ 2 + OffsetY
            .Left = ParentForm.ClientRectangle.Width \ 2 - .Width \ 2 + OffsetX
        End With
    End Sub
    Public Sub CopySpacingLeftToRight(LeftmostControl As Control, RightmostControl As Control)
        With ParentForm
            .ClientSize = New Size(LeftmostControl.Left + RightmostControl.Left + RightmostControl.Width - 1, .ClientSize.Height)
        End With
    End Sub
    'Public Overloads Sub CopySpacingTopToBottom(TopmostControl As Control, BottomControl As Control, Optional ByVal IncludeBorders As Boolean = True, Optional ByVal OffsetY As Integer = 0)
    '    If IncludeBorders = True Then
    '        ParentForm.ClientSize = New Size(ParentForm.ClientSize.Width, TopmostControl.Top + BottomControl.Bottom)
    '        ParentForm.ClientSize = New Size(ParentForm.ClientSize.Width, ParentForm.Height + OffsetY)
    '    Else
    '        ParentForm.ClientSize = New Size(ParentForm.ClientSize.Width, TopmostControl.Top + BottomControl.Bottom + OffsetY)
    '    End If
    'End Sub
    Public Sub CopySpacingTopToBottom(TopmostControl As Control, BottomControl As Control, ContainerControl As Control)
        With ContainerControl
            .ClientSize = New Size(.ClientSize.Width, TopmostControl.Top + BottomControl.Bottom - 1)
        End With
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            'RemoveHandler ParentForm.Resize, AddressOf onFormResize
            RemoveHandler ReservedTimer.Elapsed, AddressOf SlideReservation
            If disposing Then
                ReservedTimer.Dispose()

            End If
        End If
        disposedValue = True
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
    End Sub
#End Region

End Class
