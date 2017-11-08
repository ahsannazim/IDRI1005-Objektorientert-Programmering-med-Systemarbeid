Option Strict On
Option Explicit On
Option Infer Off

Imports System.Drawing
Imports System.Windows.Forms

Public Class CooldownTimer
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    Private WithEvents KeyCooler As Timers.Timer
    Public Sub New(Optional ByVal milliseconds As Double = 80)
        KeyCooler = New Timers.Timer With {.Interval = milliseconds, .AutoReset = True}
    End Sub

    Public ReadOnly Property Cool As Boolean
        Get
            Return Not KeyCooler.Enabled
        End Get
    End Property

    Public Property Cooldown As Double
        Get
            Return KeyCooler.Interval
        End Get
        Set(milliseconds As Double)
            KeyCooler.Interval = milliseconds
        End Set
    End Property

    Public Sub Tick() Handles KeyCooler.Elapsed
        KeyCooler.Enabled = False
    End Sub

    Public Overloads Sub Switch()
        KeyCooler.Enabled = Not KeyCooler.Enabled
    End Sub
    Public Overloads Sub Switch(ByVal SwitchOn As Boolean)
        KeyCooler.Enabled = SwitchOn
    End Sub
End Class

Public Class ListMenu(Of T As {New, Control})
    Private Class ItemProperties
        Private IsEnabled As Boolean = True
        Public Property Enabled As Boolean
            Get
                Return IsEnabled
            End Get
            Set(value As Boolean)
                IsEnabled = value
            End Set
        End Property
    End Class

    Private PropertiesList As New List(Of ItemProperties)
    Private SelectionTimer As CooldownTimer
    Private ItemList As List(Of T)
    Private ParentContainer As Control
    Private SelectedListIndex As Integer = -1
    Private LoopSelectionSetting As Boolean = True
    Private IsToggled As Boolean = True
    Private ParentManager As MenuManager(Of T) = Nothing
    Private MenuID As Integer = -1
    Public Event SelectionChanged(SenderMenu As Object, SenderItem As Object, Selected As Boolean)
    Public Event EnabledChanged(SenderMenu As Object, SenderItem As Object, Enabled As Boolean)
    Public Event ToggledChanged(ByVal MenuID As Integer, ByVal Toggled As Boolean, SenderList As Object)

    Public Property ID As Integer
        Get
            Return MenuID
        End Get
        Set(value As Integer)
            MenuID = value
        End Set
    End Property
    Public Property Toggled As Boolean
        Get
            Return IsToggled
        End Get
        Set(value As Boolean)
            IsToggled = value
            RaiseEvent ToggledChanged(MenuID, value, Me)
        End Set
    End Property

    Public Property LoopSelection As Boolean
        Get
            Return LoopSelectionSetting
        End Get
        Set(value As Boolean)
            LoopSelectionSetting = value
        End Set
    End Property

    Public Property SelectedIndex As Integer
        Get
            Return SelectedListIndex
        End Get
        Set(value As Integer)
            If value >= 0 Then
                If SelectionTimer.Cool = True AndAlso PropertiesList(value).Enabled = True Then
                    SelectionTimer.Switch()
                    If SelectedListIndex >= 0 Then
                        RaiseEvent SelectionChanged(Me, ItemList(SelectedListIndex), False)
                    End If
                    SelectedListIndex = value
                    RaiseEvent SelectionChanged(Me, ItemList(SelectedListIndex), True)
                End If
            Else
                RaiseEvent SelectionChanged(Me, ItemList(SelectedListIndex), False)
                SelectedListIndex = -1
            End If
        End Set
    End Property

    Public Property SelectionCooldown As Double
        Get
            Return SelectionTimer.Cooldown
        End Get
        Set(value As Double)
            SelectionTimer.Cooldown = value
        End Set
    End Property

    Public ReadOnly Property List As List(Of T)
        Get
            Return ItemList
        End Get
    End Property

    Public Property Enabled(ByVal Index As Integer) As Boolean
        Get
            Return PropertiesList(Index).Enabled
        End Get
        Set(value As Boolean)
            PropertiesList(Index).Enabled = value
            RaiseEvent EnabledChanged(Me, ItemList(Index), value)
        End Set
    End Property

    Public Sub CreateMenu(ByVal Amount As Integer, ByVal Spacing As Integer, Optional ByVal Width As Integer = 100, Optional ByVal Height As Integer = 40, Optional Left As Integer = 10, Optional Top As Integer = 10)
        For Each Item As T In ItemList
            Item.Hide()
            Item.Dispose()
        Next
        ItemList.Clear()
        For i As Integer = 0 To Amount - 1
            Dim Item As New T
            With Item
                .Hide()
                .BackColor = Color.Red
                .Width = Width
                .Height = Height
                .Left = Left
                .Top = Top + (Height + Spacing) * i
                .Tag = i
            End With
            If GetType(T) Is GetType(Label) Then
                Dim ItemObject As Object = DirectCast(Item, Object)
                Dim ItemLabel As Label = DirectCast(ItemObject, Label)
                ItemLabel.TextAlign = ContentAlignment.MiddleCenter
            End If
            ItemList.Add(Item)
            ParentContainer.Controls.Add(Item)
            Dim ItemProperties As New ItemProperties
            PropertiesList.Add(ItemProperties)
            AddHandler Item.MouseEnter, AddressOf Item_MouseEnter
        Next
        For Each Item As T In ItemList
            Item.Show()
        Next
    End Sub
    Public Sub Item_MouseEnter(Sender As Object, e As EventArgs)
        Dim SenderItem As T = DirectCast(Sender, T)
        If Not SelectedIndex = CInt(SenderItem.Tag) Then
            SelectionTimer.Tick()
            SelectedIndex = CInt(SenderItem.Tag)
        End If
    End Sub
    Public Sub SelectNext()
        If SelectedListIndex < ItemList.Count - 1 Then
            Dim i As Integer = 1
            Dim EnabledFound As Boolean = False
            Do Until EnabledFound = True
                If PropertiesList(SelectedListIndex + i).Enabled = True Then
                    EnabledFound = True
                    SelectedIndex += i
                Else
                    If SelectedListIndex + i < ItemList.Count - 1 Then
                        i += 1
                    ElseIf LoopSelection = True Then
                        i = 0 - SelectedListIndex
                    Else
                        EnabledFound = True
                    End If
                End If
            Loop
        ElseIf LoopSelectionSetting = True Then
            SelectedIndex = 0
            If PropertiesList(0).Enabled = False Then
                SelectNext()
            End If
        End If
    End Sub
    Public Overloads Sub PopulateLabels(ByVal Labels() As String)
        If GetType(T) Is GetType(Label) Then
            For i As Integer = 0 To Labels.Length - 1
                Dim ObjectItem As Object = DirectCast(ItemList(i), Object)
                Dim LabelItem As Label = DirectCast(ObjectItem, Label)
                LabelItem.Text = Labels(i)
            Next
        End If
    End Sub
    Public Overloads Sub PopulateLabels(ByVal Index As Integer, ByVal Text As String)
        If GetType(T) Is GetType(Label) Then
            Dim ObjectItem As Object = DirectCast(ItemList(Index), Object)
            Dim LabelItem As Label = DirectCast(ObjectItem, Label)
            LabelItem.Text = Text
        End If
    End Sub
    Public Sub SelectPrevious()
        If SelectedListIndex > 0 Then
            Dim i As Integer = 1
            Dim EnabledFound As Boolean = False
            Do Until EnabledFound = True
                If PropertiesList(SelectedListIndex - i).Enabled = True Then
                    EnabledFound = True
                    SelectedIndex -= i
                Else
                    If SelectedListIndex - i >= 0 Then
                        i += 1
                    ElseIf LoopSelection = True Then
                        i = SelectedListIndex - ItemList.Count + 1
                    Else
                        EnabledFound = True
                    End If
                End If
            Loop
        ElseIf LoopSelectionSetting = True Then
            SelectedIndex = ItemList.Count - 1
        End If
    End Sub

    Public Sub CenterMenu()
        For Each Item As T In ItemList
            Item.Left = CInt((ParentContainer.ClientRectangle.Width / 2) - (Item.Width / 2))
        Next
    End Sub

    Public Sub New(ByRef Parent As Control, Optional ByRef Manager As MenuManager(Of T) = Nothing)
        ParentManager = Manager
        ParentContainer = Parent
        SelectionTimer = New CooldownTimer
        ItemList = New List(Of T)
    End Sub
End Class
