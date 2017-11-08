Option Strict On
Option Explicit On
Option Infer Off

Public NotInheritable Class MultiTabWindow
    Inherits Panel
#Region "Fields"
    Private TabList As List(Of Tab)
    Private Shared ZeroPoint As New Point(0, 0)
    Private SelectedIndex As Integer = 0
    Private ScaleXY() As Boolean = {True, False}
#End Region
#Region "Events"
    Public Event TabChanged(Sender As MultiTabWindow)
    Public Event TabCountChanged(Sender As MultiTabWindow)
    Public Event TabResized(Sender As MultiTabWindow, Tab As Tab)
#End Region
#Region "Public properties"
    Public Overloads Property ScaleTabsToSize As Boolean()
        Get
            Return ScaleXY
        End Get
        Set(value As Boolean())
            ScaleXY(0) = value(0)
            ScaleXY(1) = value(1)
        End Set
    End Property
    Public Enum TabSize
        Horizontal
        Vertical
    End Enum
    Public Overloads Property ScaleTabsToSize(ByVal Component As TabSize) As Boolean
        Get
            Select Case Component
                Case TabSize.Horizontal
                    Return ScaleXY(0)
                Case Else
                    Return ScaleXY(1)
            End Select
        End Get
        Set(value As Boolean)
            Select Case Component
                Case TabSize.Horizontal
                    ScaleXY(0) = value
                Case Else
                    ScaleXY(1) = value
            End Select
        End Set
    End Property
    Public ReadOnly Property Tab(ByVal Index As Integer) As Tab
        Get
            Return TabList(Index)
        End Get
    End Property
    Public ReadOnly Property Tabs As List(Of Tab)
        Get
            Return TabList
        End Get
    End Property
    Public Property Index As Integer
        Get
            Return SelectedIndex
        End Get
        Set(value As Integer)
            If SelectedIndex >= 0 AndAlso SelectedIndex < TabList.Count Then
                TabList(SelectedIndex).Hide()
            End If
            If value >= 0 AndAlso value < TabList.Count Then
                SelectedIndex = value
                With TabList(SelectedIndex)
                    .Show()
                    .BringToFront()
                End With
            Else
                SelectedIndex = -1
            End If
        End Set
    End Property
    Private Enum WindowProperty
        Tab
        ScaleTabs
        TabCount
    End Enum
    Public ReadOnly Property CurrentTab As Tab
        Get
            If SelectedIndex >= 0 AndAlso SelectedIndex < TabList.Count Then
                Return TabList(SelectedIndex)
            Else
                Return Nothing
            End If
        End Get
    End Property
#End Region
#Region "Public methods"
    Public Sub New(Parent As Control)
        TabList = New List(Of Tab)
        With Me
            .Hide()
            .DoubleBuffered = True
            .Parent = Parent
            .Location = ZeroPoint
            .ClientSize = Parent.ClientSize
            .Show()
        End With
        AddHandler Parent.Resize, AddressOf OnParentResize
    End Sub
    Public Sub ResetAll(Optional ArgumentAll As Object = Nothing)
        For Each T As Tab In TabList
            T.ResetTab(ArgumentAll)
        Next
    End Sub
    Public Overloads Sub AddTab(Tab As Tab)
        AddTab(Tab, TabList.Count)
    End Sub
    Public Overloads Sub AddTab(Controls As List(Of Control))
        Dim iLast As Integer = Controls.Count - 1
        Dim NT As New Tab(Me)
        With NT
            For i As Integer = 0 To iLast
                .Controls.Add(Controls(i))
            Next
            .ListIndex = TabList.Count
        End With
        AddTab(NT)
    End Sub
    Public Overloads Sub AddTab(Controls() As Control)
        Dim iLast As Integer = Controls.Count - 1
        Dim NT As New Tab(Me)
        With NT
            For i As Integer = 0 To iLast
                .Controls.Add(Controls(i))
            Next
            .ListIndex = TabList.Count
        End With
        AddTab(NT)
    End Sub
    Public Overloads Sub AddTab(Tab As Tab, AtIndex As Integer)
        With Tab
            .Hide()
            .Location = ZeroPoint
            .Size = ClientSize
            .ListIndex = AtIndex
        End With
        Dim iLast As Integer
        With TabList
            .Insert(AtIndex, Tab)
            iLast = .Count - 1
        End With
        If AtIndex < iLast Then
            For i As Integer = AtIndex + 1 To iLast
                TabList(i).ListIndex = i
            Next
        End If
    End Sub
    Public Overloads Sub RemoveTab(Index As Integer)
        If Index = SelectedIndex Then
            Me.Index -= 1
        End If
        TabList.RemoveAt(Index)
        Dim iLast As Integer = TabList.Count - 1
        If iLast >= Index Then
            For i As Integer = Index To TabList.Count - 1
                TabList(i).ListIndex = i
            Next
        End If
    End Sub
    Public Sub ShowTab(ByVal Index As Integer)
        Me.Index = Index
    End Sub
#End Region
#Region "Private methods"
    Private Sub OnParentResize(Sender As Object, e As EventArgs)
        With Me
            .SuspendLayout()
            .ClientSize = Parent.ClientSize
            If .CurrentTab IsNot Nothing Then
                .CurrentTab.RefreshLayout()
            End If
            .ResumeLayout()
        End With
    End Sub
    Protected Overrides Sub OnParentChanged(e As EventArgs)
        Try
            MyBase.OnParentChanged(e)
            With Me
                .Location = ZeroPoint
                .ClientSize = .Parent.ClientSize
            End With
            If CurrentTab IsNot Nothing Then
                CurrentTab.RefreshLayout()
            End If
        Catch
            MsgBox("Feil")
        End Try
    End Sub
#End Region
End Class


Public Class TabCancelEventArgs
    Inherits EventArgs
    Public Cancel As Boolean
End Class

''' <summary>
''' A view or screen to be displayed in a MultiTabWindow, used for keeping several views within a single window.
''' </summary>
Public Class Tab
    Inherits Control
    Private varScaleToParent As Boolean = True
    Private varRefreshLayoutOnShow As Boolean = True
    Protected Friend ListIndex As Integer = -1
    Public Event LayoutRefreshed(Sender As Object, e As EventArgs)
    Public Event TabClosing(Sender As Object, e As TabClosingEventArgs)
    Public Event TabClosed(Sender As Object, e As EventArgs)
    ''' <summary>
    ''' Gets or sets a value indicating whether or not this Tab should handle its parent's Resize event and resize itself accordingly.
    ''' Default = True.
    ''' </summary>
    Public Property ScaleToWindow As Boolean
        Get
            Return varScaleToParent
        End Get
        Set(value As Boolean)
            varScaleToParent = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value specifying whether or not to automatically call the RefreshLayout method when the tab is shown (default = true).
    ''' </summary>
    Public Property RefreshLayoutOnShow As Boolean
        Get
            Return varRefreshLayoutOnShow
        End Get
        Set(value As Boolean)
            varRefreshLayoutOnShow = value
        End Set
    End Property
    ''' <summary>
    ''' If the ScaleToWindow property is set to true, but no new Resize event is raised by the parent, sets this tab's size to the ClientSize of its parent MultiTabWindow.
    ''' </summary>
    Public Overridable Sub RefreshLayout()
        OnLayoutRefreshed(New TabCancelEventArgs)
    End Sub
    ''' <summary>
    ''' Raises the LayoutRefreshed event.
    ''' </summary>
    ''' <param name="e">Contains information about this event, if any.</param>
    Protected Overridable Sub OnLayoutRefreshed(e As TabCancelEventArgs)
        If Not e.Cancel Then
            If varScaleToParent AndAlso Parent IsNot Nothing Then
                ClientSize = Parent.ClientSize
            End If
            RaiseEvent LayoutRefreshed(Me, EventArgs.Empty)
        End If
    End Sub
    ''' <summary>
    ''' Shows this tab. This action should only be performed internally 
    ''' </summary>
    Protected Friend Overridable Shadows Sub Show()
        If varRefreshLayoutOnShow Then
            RefreshLayout()
        End If
        MyBase.Show()
    End Sub
    ''' <summary>
    ''' If properly overridden in a class that inherits Tab, sets all variables and states to their initial values, erasing changes made to the Tab.
    ''' </summary>
    ''' <param name="Arguments">Optional parameter that may or may not be used by classes derived from Tab.</param>
    Public Overridable Sub ResetTab(Optional Arguments As Object = Nothing)
        ' Override and call in Globals.Logout method.
    End Sub
    ''' <summary>
    ''' Gets or sets the MultiTabWindow instance that contains this Tab.
    ''' </summary>
    Public Shadows Property Parent As MultiTabWindow
        Get
            Return DirectCast(MyBase.Parent, MultiTabWindow)
        End Get
        Set(value As MultiTabWindow)
            If Parent IsNot Nothing AndAlso ListIndex >= 0 Then
                Parent.RemoveTab(ListIndex)
            End If
            MyBase.Parent = value
            If Parent IsNot Nothing Then
                Parent.AddTab(Me)
            End If
        End Set
    End Property
    ''' <summary>
    ''' Creates a new Tab instance and inserts it into the specified MultiTabWindow (at the last index). For complex (form-like) tabs, inherit this class instead.
    ''' </summary>
    ''' <param name="Parent">The MultiTabWindow instance to which this tab should be added.</param>
    Public Sub New(Parent As MultiTabWindow)
        Hide()
        DoubleBuffered = True
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        UpdateStyles()
        Me.Parent = Parent
    End Sub
    ''' <summary>
    ''' Creates a new Tab instance. See the other overload. For complex (form-like) tabs, inherit this class instead.
    ''' </summary>
    Public Sub New()
        Hide()
        DoubleBuffered = True
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        UpdateStyles()
    End Sub
    ''' <summary>
    ''' Closes the Tab, which may or may not result in disposal, depending on the state of the TabClosingEventArgs after it has been passed to the overridable method OnClosing(e as TabClosingEventArgs), in which the Tab's closing may be canceled.
    ''' </summary>
    ''' <param name="Dispose">Specifies whether or not to automatically release the resources used by this Tab after closing.</param>
    Public Sub Close(Optional Dispose As Boolean = True)
        Dim Args As New TabClosingEventArgs(False, Dispose)
        OnClosing(New TabClosingEventArgs)
    End Sub
    ''' <summary>
    ''' Raises the TabClosing event and calls the OnClosed method unless canceled.
    ''' </summary>
    ''' <param name="e">Contains information about how to proceed.</param>
    Protected Overridable Sub OnClosing(e As TabClosingEventArgs)
        RaiseEvent TabClosing(Me, e)
        If Not e.Cancel Then
            OnClosed(e)
        End If
    End Sub
    ''' <summary>
    ''' Raises the TabClosed event and calls Dispose(True), depending on the Dispose property of the TabClosingEventArgs passed to it.
    ''' </summary>
    ''' <param name="e">Contains information about whether or not to proceed. e.Cancel is not checked at this point.</param>
    Protected Overridable Sub OnClosed(e As TabClosingEventArgs)
        RaiseEvent TabClosed(Me, EventArgs.Empty)
        If e.Dispose Then
            Dispose(True)
        End If
    End Sub
    ''' <summary>
    ''' Raises the ParentChanged event and sets this Tab's size to the ClientSize of its parent.
    ''' </summary>
    Protected Overrides Sub OnParentChanged(e As EventArgs)
        MyBase.OnParentChanged(e)
        ClientSize = Parent.ClientSize
    End Sub
    Protected Overrides Sub Dispose(disposing As Boolean)
        MyBase.Dispose(disposing)
    End Sub
End Class
Public Class TabClosingEventArgs
    Inherits EventArgs
    Private DoCancel As Boolean
    Private DisposeOnClose As Boolean
    Public Property Cancel As Boolean
        Get
            Return DoCancel
        End Get
        Set(value As Boolean)
            DoCancel = value
        End Set
    End Property
    Public Property Dispose As Boolean
        Get
            Return DisposeOnClose
        End Get
        Set(value As Boolean)
            DisposeOnClose = value
        End Set
    End Property
    Public Sub New(Optional ByVal Cancel As Boolean = False, Optional ByVal Dispose As Boolean = True)
        DoCancel = Cancel
        DisposeOnClose = Dispose
    End Sub
End Class