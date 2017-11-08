Option Strict On
Option Explicit On
Option Infer Off

Imports System.Windows.Forms

Public Class MenuManager(Of T As {New, Control})
    Inherits Control
    ' Local variables
    Private MenuList As List(Of ListMenu(Of T))
    Private CurrentlyToggled As ListMenu(Of T)
    Private ToggledID As Integer = 0
    Private AutoLayoutOn As Boolean = False
    Private NameDictionary As Dictionary(Of String, ListMenu(Of T))

    ' Events
    Public Event ToggledChanged(SenderMenu As Object, Toggled As Boolean)
    Public Event SelectionChanged(SenderMenu As Object, SenderItem As Object, Selected As Boolean)
    Public Shadows Event EnabledChanged(SenderMenu As Object, SenderItem As Object, Enabled As Boolean)

    ' Properties
    Public ReadOnly Property Count As Integer
        Get
            Return MenuList.Count
        End Get
    End Property

    Public Property AutoLayout As Boolean
        Get
            Return AutoLayoutOn
        End Get
        Set(value As Boolean)
            AutoLayoutOn = value
        End Set
    End Property

    Public ReadOnly Property List As List(Of ListMenu(Of T))
        Get
            Return MenuList
        End Get
    End Property

    Public Property ToggledMenu As ListMenu(Of T)
        Get
            Return CurrentlyToggled
        End Get
        Set(value As ListMenu(Of T))
            CurrentlyToggled = value
        End Set
    End Property
    ''' <summary>
    ''' Returns the ListMenu associated with the specified key (string) in the MenuManager's dictionary.
    ''' </summary>
    ''' <param name="Key">If specified, the string or name associated with the ListMenu.</param>
    ''' <returns>The ListMenu associated with the specified key.</returns>
    Public Overloads ReadOnly Property MenuAtIndex(ByVal Key As String) As ListMenu(Of T)
        Get
            Return NameDictionary(Key)
        End Get
    End Property
    ''' <summary>
    ''' Returns the ListMenu associated at the specified index in the MenuManager's dictionary, usually following the order in which ListMenus were added to or created within the manager.
    ''' </summary>
    ''' <param name="Index">The index of the ListMenu in the dictionary.</param>
    ''' <returns>The ListMenu associated with the specified key.</returns>
    Public Overloads ReadOnly Property MenuAtIndex(ByVal Index As Integer) As ListMenu(Of T)
        Get
            Return MenuList(Index)
        End Get
    End Property

    Public ReadOnly Property Key(ByVal Index As Integer) As String
        Get
            Return NameDictionary.Keys.ElementAt(Index)
        End Get
    End Property

    ' Methods

    Public Sub Clear()
        MenuList.Clear()
        NameDictionary.Clear()
        ToggledID = 0
    End Sub

    Public Sub CreateMenu(ByVal Amount As Integer, Optional ByVal Width As Integer = 100, Optional ByVal Height As Integer = 40, Optional ByVal Left As Integer = 10, Optional ByVal Top As Integer = 10, Optional ByVal Spacing As Integer = 10)
        Dim LM As New ListMenu(Of T)(Me, Me)
        LM.CreateMenu(Amount, Spacing, Width, Height, Left, Top)
        AddMenu(LM)
    End Sub

    Private Sub LM_SelectionChanged(SenderMenu As Object, SenderItem As Object, Selected As Boolean)
        RaiseEvent SelectionChanged(SenderMenu, SenderItem, Selected)
    End Sub
    Private Sub LM_EnabledChanged(SenderMenu As Object, SenderItem As Object, Enabled As Boolean)
        RaiseEvent EnabledChanged(SenderMenu, SenderItem, Enabled)
    End Sub
    Private Sub LM_ToggledChanged(ByVal MenuID As Integer, ByVal Toggled As Boolean, SenderList As Object)
        RaiseEvent ToggledChanged(SenderList, Toggled)
    End Sub

    Public Sub AddMenu(ByRef ListMenu As ListMenu(Of T), Optional ByVal DictionaryKey As String = Nothing)
        If DictionaryKey Is Nothing Then
            NameDictionary.Add(CStr(NameDictionary.Count), ListMenu)
        Else
            NameDictionary.Add(DictionaryKey, ListMenu)
        End If
        ListMenu.ID = MenuList.Count
        MenuList.Add(ListMenu)
        AddHandler ListMenu.SelectionChanged, AddressOf LM_SelectionChanged
        AddHandler ListMenu.EnabledChanged, AddressOf LM_EnabledChanged
        AddHandler ListMenu.ToggledChanged, AddressOf LM_ToggledChanged
        If ListMenu.ID <> ToggledID Then
            ListMenu.Toggled = False
        Else
            ListMenu.Toggled = True
        End If
    End Sub

    Public Sub SelectNext()
        MenuList(ToggledID).SelectNext()
    End Sub
    Public Sub SelectPrevious()
        MenuList(ToggledID).SelectPrevious()
    End Sub
    Public Sub Toggle(ByRef Menu As ListMenu(Of T))

    End Sub
    Public Sub ToggleNext()
        MenuList(ToggledID).Toggled = False
        If ToggledID < MenuList.Count - 1 Then
            ToggledID += 1
        End If
        MenuList(ToggledID).Toggled = True
    End Sub
    Public Sub TogglePrevious()
        MenuList(ToggledID).Toggled = False
        If ToggledID > 0 Then
            ToggledID -= 1
        End If
        MenuList(ToggledID).Toggled = True
    End Sub

    Public Sub New()
        MenuList = New List(Of ListMenu(Of T))
        NameDictionary = New Dictionary(Of String, ListMenu(Of T))
    End Sub

    Public Sub ToggleMenu(ByVal MenuID As Integer)
        MenuList(ToggledID).Toggled = False
        MenuList(MenuID).Toggled = True
        ToggledID = MenuID
    End Sub
End Class