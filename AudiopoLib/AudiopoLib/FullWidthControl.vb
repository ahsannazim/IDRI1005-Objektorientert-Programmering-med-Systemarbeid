Option Strict On
Option Explicit On
Option Infer Off

Imports System.Windows.Forms
Imports System.Drawing

Public Class FullWidthControl
    Inherits Label
    Private NormBG As Color = Color.FromArgb(0, 72, 83)
    Private SelectedBG As Color = Color.FromArgb(0, 42, 50)
    Private DisabledBG As Color = Color.FromArgb(50, 50, 50)
    Private MouseDownBG As Color = Color.FromArgb(0, 25, 40)
    Private Snap As SnapType = SnapType.Top
    Private Offset As Integer
    Private IsEnabled As Boolean = True
    Public Shadows Event Click(sender As Object, e As EventArgs)
    Public Shadows Event EnabledChanged(value As Boolean)
    Public Shadows Property Enabled() As Boolean
        Get
            Return IsEnabled
        End Get
        Set(value As Boolean)
            RaiseEvent EnabledChanged(value)
        End Set
    End Property
    Public Property BackColorSelected As Color
        Get
            Return SelectedBG
        End Get
        Set(value As Color)
            SelectedBG = value
            If Me.Focused Then
                BackColor = value
            End If
        End Set
    End Property
    Public Property BackColorNormal As Color
        Get
            Return NormBG
        End Get
        Set(value As Color)
            NormBG = value
            If Not Me.Focused AndAlso Me.Enabled Then
                BackColor = value
            End If
        End Set
    End Property
    Public Property BackColorDisabled As Color
        Get
            Return DisabledBG
        End Get
        Set(value As Color)
            DisabledBG = value
            If Not Me.Enabled Then
                BackColor = value
            End If
        End Set
    End Property
    Public Sub BehaveLikeButton()
        Try
            BackColor = NormBG
            AddHandler Me.MouseUp, AddressOf Me_MouseUp
            AddHandler Me.MouseDown, AddressOf Me_MouseDown
            AddHandler Me.MouseClick, AddressOf Me_MouseClick
            AddHandler Me.MouseEnter, AddressOf Me_Selected
            AddHandler Me.GotFocus, AddressOf Me_Selected
            AddHandler Me.MouseLeave, AddressOf Me_Deselected
            AddHandler Me.LostFocus, AddressOf Me_Deselected
            TabStop = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub BehaveLikeLabel()
        Try
            RemoveHandler Me.MouseUp, AddressOf Me_MouseUp
            RemoveHandler Me.MouseDown, AddressOf Me_MouseDown
            RemoveHandler Me.MouseClick, AddressOf Me_MouseClick
            RemoveHandler Me.MouseEnter, AddressOf Me_Selected
            RemoveHandler Me.GotFocus, AddressOf Me_Selected
            RemoveHandler Me.MouseLeave, AddressOf Me_Deselected
            RemoveHandler Me.LostFocus, AddressOf Me_Deselected
            TabStop = False
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub New(Parent As Control, Optional IsButton As Boolean = False, Optional SnapTo As SnapType = SnapType.Top, Optional OffsetY As Integer = 0)
        Visible = False
        Me.Parent = Parent
        DoubleBuffered = True
        If IsButton = True Then
            SetStyle(ControlStyles.Selectable, True)
        Else
            SetStyle(ControlStyles.Selectable, False)
        End If
        InitializeLayout(SnapTo, OffsetY, IsButton)
    End Sub
    Private Sub Me_Selected(sender As Object, e As EventArgs)
        If Enabled Then
            BackColor = SelectedBG
        End If
    End Sub
    Private Sub Me_MouseUp(sender As Object, e As MouseEventArgs)
        If Enabled Then
            BackColor = SelectedBG
            RaiseEvent Click(Me, Nothing)
            Debug.Print("True")
        Else
            Debug.Print("False")
        End If
    End Sub
    Private Sub Me_Deselected(sender As Object, e As EventArgs)
        If Enabled Then
            If Not Focused Then
                BackColor = NormBG
            End If
        End If
    End Sub
    Private Sub Me_EnabledChanged(value As Boolean) Handles Me.EnabledChanged
        If value = True Then
            If Not Focused Then
                BackColor = NormBG
            Else
                BackColor = SelectedBG
            End If
        Else
            BackColor = DisabledBG
        End If
        IsEnabled = value
    End Sub
    Private Sub Me_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If Enabled Then
            Select Case e.KeyCode
                Case Keys.Space, Keys.Enter
                    RaiseEvent Click(Me, Nothing)
            End Select
        End If
    End Sub
    Private Sub Me_MouseClick(sender As Object, e As MouseEventArgs)
        'If Enabled Then
        '    If e.Button = MouseButtons.Left Then
        '        RaiseEvent Click(Me, Nothing)
        '    End If
        'End If
    End Sub
    Private Sub Me_MouseDown(sender As Object, e As MouseEventArgs)
        If Enabled Then
            BackColor = MouseDownBG
        End If
    End Sub
    Public Enum SnapType
        Top
        Bottom
    End Enum
    Private Sub InitializeLayout(Optional SnapTo As SnapType = SnapType.Top, Optional OffsetY As Integer = 0, Optional IsButton As Boolean = False)
        Snap = SnapTo
        Offset = OffsetY
        ForeColor = Color.White
        BackColor = Color.FromArgb(0, 72, 83)
        AutoSize = False
        Text = "Test"
        Height = 30
        TextAlign = ContentAlignment.MiddleCenter
        Width = Parent.ClientSize.Width
        Left = 0
        If SnapTo = SnapType.Top Then
            Top = OffsetY
        Else
            Top = Parent.ClientSize.Height - Height + OffsetY
        End If
        If IsButton = True Then
            BehaveLikeButton()
        Else
            BehaveLikeLabel()
        End If
        Show()
    End Sub
End Class
