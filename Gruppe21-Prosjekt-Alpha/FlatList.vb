'Option Strict On
'Option Explicit On
'Option Infer Off

'Imports AudiopoLib

'Public Class Donasjoner
'    Inherits ContainerControl
'    Private Wrapper As New Control
'    Private WithEvents HeaderControl As New FullWidthControl(Me)
'    Private FList As New FlatList
'    Public ReadOnly Property Header As FullWidthControl
'        Get
'            Return HeaderControl
'        End Get
'    End Property
'    Public ReadOnly Property List As FlatList
'        Get
'            Return FList
'        End Get
'    End Property
'    Public Sub New(Parent As Control)
'        Hide()
'        ResizeRedraw = True
'        Me.Parent = Parent
'        With Wrapper
'            .Parent = Me
'        End With
'        BackColor = Color.FromArgb(200, 200, 200)
'        With Me
'            .Size = New Size(400, 500)
'        End With
'        With HeaderControl
'            .BackColor = Color.FromArgb(200, 200, 200)
'            .TextAlign = ContentAlignment.MiddleLeft
'            .ForeColor = Color.FromArgb(60, 60, 60)
'            .Text = "Tidligere blodtappinger"
'        End With
'        With FList
'            .Left = 1
'            .Top = HeaderControl.Bottom
'            .Parent = Wrapper
'        End With
'        Show()
'    End Sub
'    Private Sub OnHeaderResize() Handles HeaderControl.Resize
'        With FList
'            .Top = HeaderControl.Bottom
'        End With
'    End Sub
'    Public Shadows Property BackColor As Color
'        Get
'            Return Wrapper.BackColor
'        End Get
'        Set(value As Color)
'            Wrapper.BackColor = value
'            MyBase.BackColor = value
'            FList.BackColor = value
'        End Set
'    End Property
'    Protected Overrides Sub OnResize(e As EventArgs)
'        MyBase.OnResize(e)
'        With HeaderControl
'            .Width = Width - 1
'            .Left = 1
'        End With
'        With Wrapper
'            .Width = Width - 1
'            .Height = Height
'        End With
'        With FList
'            .Width = Wrapper.Width + 1
'            .Height = Wrapper.Height - HeaderControl.Bottom - 1
'        End With
'    End Sub
'End Class

'Public Class FlatList
'    Inherits Panel
'    Private varItemHeight As Integer = 40
'    Private varItemSpacing As Integer = 1
'    Private varNewItemBG As Color = Color.White
'    Private RowList As New List(Of FlatListRow)
'    Public Property DefaultColor As Color
'        Get
'            Return varNewItemBG
'        End Get
'        Set(value As Color)
'            varNewItemBG = value
'        End Set
'    End Property
'    Public ReadOnly Property Item(ByVal Index As Integer) As FlatListRow
'        Get
'            Return RowList(Index)
'        End Get
'    End Property
'    Public ReadOnly Property Items As List(Of FlatListRow)
'        Get
'            Return RowList
'        End Get
'    End Property
'    Public Property ItemSpacing As Integer
'        Get
'            Return varItemSpacing
'        End Get
'        Set(value As Integer)
'            varItemSpacing = value
'        End Set
'    End Property
'    Public Property ItemHeight As Integer
'        Get
'            Return varItemHeight
'        End Get
'        Set(value As Integer)
'            varItemHeight = value
'        End Set
'    End Property
'    Public Sub AddItem()
'        Dim NewItem As New FlatListRow(Me, 3)
'        With NewItem
'            .Height = varItemHeight
'            If RowList.Count > 0 Then
'                .Top = RowList(RowList.Count - 1).Bottom + varItemSpacing
'            Else
'                .Top = 0
'            End If
'            .BackColor = varNewItemBG
'            .Show()
'        End With
'        RowList.Add(NewItem)
'    End Sub
'    Public Sub New()
'        With Me
'            '.HorizontalScroll.Enabled = False
'            .HorizontalScroll.Maximum = 0
'            .AutoScroll = False
'            .VerticalScroll.Visible = False
'            .AutoScroll = True
'            .Size = New Size(400, 500)
'        End With
'    End Sub
'    'Public Function GetValues(ByVal Row As Integer) As Object()

'    'End Function
'    Protected Overrides Sub OnSizeChanged(e As EventArgs)
'        MyBase.OnSizeChanged(e)
'        Dim iLast As Integer = RowList.Count - 1
'        For i As Integer = 0 To iLast
'            RowList(i).Width = Width
'        Next
'    End Sub
'End Class
'Public Class FlatListRow
'    Inherits Control
'    'Private ItemArr() As FlatListitem
'    'Public Function GetValues() As Object()
'    '    Return ItemArr
'    'End Function
'    Public Function GetValue(Of T)(ByVal Item As Integer) As T
'        'Return DirectCast(ItemArr(Item).Value, T)
'    End Function
'    Protected Friend Sub New(ParentList As FlatList, ByVal Slots As Integer)
'        'ItemArr = New FlatListitem(Slots) {}
'        Hide()
'        Parent = ParentList
'        With ParentList
'            Width = .Width
'        End With
'        Show()
'    End Sub
'    Protected Overrides Sub OnPaint(e As PaintEventArgs)
'        MyBase.OnPaint(e)

'    End Sub
'End Class

''Public Class FlatListitem
''    Inherits Control
''    Private TextBrush As SolidBrush
''    Private varValue As Object = Nothing
''    Public Shadows Property Parent As FlatListRow
''        Get
''            Return DirectCast(MyBase.Parent, FlatListRow)
''        End Get
''        Set(value As FlatListRow)
''            MyBase.Parent = value
''        End Set
''    End Property
''    Public Property Value As Object
''        Get
''            Return varValue
''        End Get
''        Set(value As Object)
''            varValue = value
''            Text = CStr(value)
''        End Set
''    End Property
''    Protected Friend Sub New(ParentRow As FlatListRow)
''        TextBrush = New SolidBrush(ForeColor)
''        Parent = ParentRow
''    End Sub
''End Class
