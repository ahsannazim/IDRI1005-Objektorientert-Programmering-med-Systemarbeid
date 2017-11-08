Imports AudiopoLib
Imports System.Threading

Public NotInheritable Class UserNotificationContainer
    Inherits BorderControl
    Private SC As SynchronizationContext = SynchronizationContext.Current
    Private varID As Integer
    Private WithEvents CloseTimer, TopTimer As New Timers.Timer(16)
    Private NotificationList As New List(Of UserNotification)
    Private WithEvents NotificationCounter As New Label
    Private LoadingSurface As New PictureBox
    Private LG As LoadingGraphics(Of PictureBox)
    Private Header As FullWidthControl
    Private EmptyLabel As New Label
    Private varDisplayEmptyLabel As Boolean
    Private ClosingNotifications As New List(Of UserNotification)
    Private NotificationContainer As New Control
    Private WithEvents RefreshButton As New PictureBox
    Public Event Reloaded(Sender As Object, e As EventArgs)
    Public Sub Reload(Optional Args As Object = Nothing)
        Clear()
        Spin()
        RaiseEvent Reloaded(Me, EventArgs.Empty)
    End Sub
    Public Sub Clear()
        For Each N As UserNotification In NotificationList
            N.Dispose()
        Next
        NotificationList.Clear()
        NotificationCounter.Text = "0"
    End Sub
    Public Property ID As Integer
        Get
            Return varID
        End Get
        Set(value As Integer)
            varID = value
        End Set
    End Property
    Private Sub Refresh_Enter() Handles RefreshButton.MouseEnter
        With RefreshButton
            .BackgroundImage = My.Resources.RefreshHover
            .BackColor = Color.FromArgb(180, 180, 180)
        End With
    End Sub
    Private Sub Refresh_Leave() Handles RefreshButton.MouseLeave
        With RefreshButton
            .BackgroundImage = My.Resources.Refresh
            .BackColor = Header.BackColor
        End With
    End Sub
    Private Sub Refresh_Click() Handles RefreshButton.Click
        If Not LoadingSurface.Visible Then
            Reload()
        End If
    End Sub
    Private Sub CloseTimer_Tick() Handles CloseTimer.Elapsed
        SC.Post(AddressOf AdjustWidth, Nothing)
    End Sub
    Private Sub TopTimer_Tick() Handles TopTimer.Elapsed
        SC.Post(AddressOf AdjustTop, Nothing)
    End Sub
    Private Sub NotificationFinished(Sender As Notification)
        NotificationContainer.Hide()
        Sender.Dispose()
    End Sub
    Public Sub ShowMessage(Message As String, Style As NotificationAppearance)
        LG.StopSpin()
        Dim Notification As New Notification(NotificationContainer, Style, Message, 5, AddressOf NotificationFinished, FloatX.FillWidth, FloatY.FillHeight)
    End Sub
    Private Sub RemoveAt(Index As Integer)
        With NotificationList
            .RemoveAt(Index)
            NotificationCounter.Text = CStr(.Count)
        End With
    End Sub
    Private Sub AdjustTop(State As Object)
        SuspendLayout()
        Dim iLast As Integer = NotificationList.Count - 1
        Dim ProblemFound As Boolean
        With NotificationList(0)
            If .Top >= Header.Bottom + 30 Then
                ProblemFound = True
                .Top -= 10
            Else
                .Top = Header.Bottom + 20
            End If
        End With
        If iLast > 0 Then
            For i As Integer = 1 To iLast
                With NotificationList(i)
                    If .Top >= NotificationList(i - 1).Bottom + 30 Then
                        .Top -= 10
                        ProblemFound = True
                    Else
                        .Top = NotificationList(i - 1).Bottom + 20
                    End If
                End With
            Next
        End If
        ResumeLayout(True)
        If ProblemFound Then
            TopTimer.Start()
        End If
    End Sub
    Private Sub AdjustWidth(State As Object)
        If ClosingNotifications.Count > 0 Then
            Dim iLast As Integer = ClosingNotifications.Count - 1
            For i As Integer = iLast To 0 Step -1
                Dim N As UserNotification = ClosingNotifications(i)
                With N
                    If .Width > 40 Then
                        .Width -= 40
                    Else
                        .IsClosed = True
                        ClosingNotifications.RemoveAt(i)
                    End If
                End With
            Next
            Dim nLast As Integer = NotificationList.Count - 1
            Dim StartTimer As Boolean
            For n As Integer = nLast To 0 Step -1
                Dim CurrentNotification As UserNotification = NotificationList(n)
                With CurrentNotification
                    If .IsClosed Then
                        RemoveAt(n)
                        .Dispose()
                        StartTimer = True
                    End If
                End With
            Next
            If ClosingNotifications.Count > 0 Then
                CloseTimer.Start()
            End If
            If NotificationList.Count > 0 AndAlso StartTimer Then
                TopTimer.Start()
            End If
        End If
    End Sub
    Public Sub Spin(Optional SwitchOn As Boolean = True)
        If SwitchOn Then
            LG.Spin(30, 10)
            NotificationCounter.Text = NotificationList.Count.ToString
            DisplayEmptyLabel = False
        Else
            LG.StopSpin()
            NotificationCounter.Text = NotificationList.Count.ToString
            DisplayEmptyLabel = True
        End If
    End Sub
    Public Property DisplayEmptyLabel As Boolean
        Get
            Return varDisplayEmptyLabel
        End Get
        Set(value As Boolean)
            varDisplayEmptyLabel = value
            RefreshLabel()
        End Set
    End Property
    Private Sub RefreshLabel()
        If Not varDisplayEmptyLabel Then
            EmptyLabel.Hide()
        End If
        If CInt(NotificationCounter.Text) > 0 Then
            NotificationCounter.BackColor = Color.FromArgb(0, 99, 157)
            EmptyLabel.Hide()
        Else
            NotificationCounter.BackColor = Color.FromArgb(180, 180, 180)
            If varDisplayEmptyLabel Then
                EmptyLabel.Show()
            End If
        End If
    End Sub
    Private Sub CounterChanged() Handles NotificationCounter.TextChanged
        RefreshLabel()
    End Sub
    Public Sub New(BorderColor As Color)
        MyBase.New(BorderColor)
        Hide()
        SuspendLayout()
        Size = New Size(500, 500)
        Header = New FullWidthControl(Me)
        With Header
            .Width -= 2
            .Location = New Point(1, 1)
            .BackColor = Color.FromArgb(220, 220, 220)
            .ForeColor = Color.FromArgb(60, 60, 60)
            .Text = "Notifikasjoner og gjøremål"
        End With
        With NotificationContainer
            .Parent = Me
            .Location = New Point(1, Header.Bottom)
            .Size = New Size(500 - 2, 500 - .Top - 1)
            .Hide()
        End With
        With RefreshButton
            .Size = New Size(30, 30)
            .BackgroundImage = My.Resources.Refresh
            .Parent = Header
            .Left = Header.Width - 30
            .BackColor = Header.BackColor
        End With
        With NotificationCounter
            .Parent = Header
            .Size = New Size(Header.Height, Header.Height)
            .BackColor = Color.FromArgb(180, 180, 180)
            .ForeColor = Color.White
            .Font = New Font(.Font.FontFamily, 15)
            .TextAlign = ContentAlignment.MiddleCenter
            .Text = "0"
        End With
        With EmptyLabel
            .Hide()
            .Parent = Me
            .AutoSize = False
            .Size = New Size(100, 40)
            .BackColor = Color.White
            .ForeColor = Color.FromArgb(80, 80, 80)
            .Text = "Ingenting å vise"
            .TextAlign = ContentAlignment.MiddleCenter
            .Location = New Point(Width \ 2 - .Width \ 2, (Height - .Height + Header.Bottom) \ 2)
        End With
        With LoadingSurface
            .Hide()
            .Size = New Size(50, 50)
            .Parent = Me
            .Location = New Point(Width \ 2 - .Width \ 2, (Height - .Height + Header.Bottom) \ 2)
        End With
        CloseTimer.AutoReset = False
        TopTimer.AutoReset = False
        LG = New LoadingGraphics(Of PictureBox)(LoadingSurface)
        With LG
            .Stroke = 3
            .Pen.Color = Color.FromArgb(230, 50, 80)
        End With
        BackColor = Color.White
        Show()
        ResumeLayout(True)
    End Sub
    Public Sub AddNotification(Text As String, ID As Object, ClickAction As Action(Of UserNotification, UserNotificationEventArgs), Color As Color, RelatedElement As Object, ElementType As DatabaseElementType)
        Dim NewNotification As New UserNotification(Me, Text, ID, ClickAction, Color, RelatedElement, ElementType)
        With NotificationList
            'If .Count > 0 Then
            '    NewNotification.Top = .Last.Bottom + 20
            'Else
            '    NewNotification.Top = Header.Bottom + 20
            'End If
            .Insert(0, NewNotification)
            NotificationList(0).Top = Header.Bottom + 20
            Dim iLast As Integer = .Count - 1
            If iLast > 0 Then
                For i As Integer = 1 To iLast
                    NotificationList(i).Top = NotificationList(i - 1).Bottom + 20
                Next
            End If
            NotificationCounter.Text = CStr(.Count)
        End With
        varDisplayEmptyLabel = True
        LG.StopSpin()
    End Sub
    Public Sub AddNotification(NewNotification As UserNotification)
        With NotificationList
            'If .Count > 0 Then
            '    NewNotification.Top = .Last.Bottom + 20
            'Else
            '    NewNotification.Top = Header.Bottom + 20
            'End If
            .Insert(0, NewNotification)
            NotificationList(0).Top = Header.Bottom + 20
            Dim iLast As Integer = .Count - 1
            If iLast > 0 Then
                For i As Integer = 1 To iLast
                    NotificationList(i).Top = NotificationList(i - 1).Bottom + 20
                Next
            End If
            NotificationCounter.Text = CStr(.Count)
        End With
        varDisplayEmptyLabel = True
        LG.StopSpin()
    End Sub
    Protected Friend Sub CloseNotification(Sender As UserNotification)
        With NotificationList
            Dim iLast As Integer = .Count - 1
            Dim MatchFound As Boolean
            For i As Integer = 0 To iLast
                If NotificationList(i) Is Sender Then
                    ClosingNotifications.Add(NotificationList(i))
                    MatchFound = True
                    Exit For
                End If
            Next
            If MatchFound Then
                CloseTimer.Start()
            End If
        End With
    End Sub
    Public Sub RemoveNotification(ID As Object)
        With NotificationList
            Dim iLast As Integer = .Count - 1
            Dim MatchFound As Boolean
            For i As Integer = 0 To iLast
                Dim Notification As UserNotification = NotificationList(i)
                If Notification.ID.Equals(ID) Then
                    Notification.Dispose()
                    .RemoveAt(i)
                    MatchFound = True
                    Exit For
                End If
            Next
            If MatchFound AndAlso .Count > 0 Then
                iLast = .Count - 1
                NotificationList(0).Top = Header.Bottom + 20
                If iLast > 0 Then
                    For i As Integer = 1 To iLast
                        NotificationList(i).Top = NotificationList(i - 1).Bottom + 20
                    Next
                End If
            Else

            End If
            NotificationCounter.Text = CStr(.Count)
        End With
    End Sub
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing Then
                LG.Dispose()
                TopTimer.Dispose()
                CloseTimer.Dispose()
                For Each N As UserNotification In NotificationList
                    N.Dispose()
                Next
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub
End Class


Public Class UserNotification
    Inherits Control
    Private WithEvents Textlab As New Label
    Private varDefaultColor, varHoverColor, varPressColor As Color
    Private varID, varRelatedElement As Object
    Private varIsClosed, varCloseOnClick, varIsSelected, varInfoLoaded As Boolean
    Private varClickAction As Action(Of UserNotification, UserNotificationEventArgs)
    Private varRelatedDonor As Donor
    Private varElementType As DatabaseElementType

    Public Property IsSelected As Boolean
        Get
            Return varIsSelected
        End Get
        Set(value As Boolean)
            varIsSelected = value
        End Set
    End Property
    Public Property IsClosed As Boolean
        Get
            Return varIsClosed
        End Get
        Set(value As Boolean)
            varIsClosed = value
            If value Then
                Hide()
            End If
        End Set
    End Property
    Public ReadOnly Property RelatedDonor As Donor
        Get
            Return varRelatedDonor
        End Get
    End Property
    Public ReadOnly Property InfoLoaded As Boolean
        Get
            Return varInfoLoaded
        End Get
    End Property
    Public ReadOnly Property ElementType As DatabaseElementType
        Get
            Return varElementType
        End Get
    End Property
    Public Property RelatedElement As Object
        Get
            Return varRelatedElement
        End Get
        Set(value As Object)
            varRelatedElement = value
        End Set
    End Property
    Public Property CloseOnClick As Boolean
        Get
            Return varCloseOnClick
        End Get
        Set(value As Boolean)
            varCloseOnClick = value
        End Set
    End Property
    Private Sub TextLab_MouseDown(Sender As Object, e As MouseEventArgs) Handles Textlab.MouseDown
        OnMouseDown(e)
    End Sub
    Private Sub TextLab_MouseUp(Sender As Object, e As MouseEventArgs) Handles Textlab.MouseUp
        OnMouseUp(e)
    End Sub
    Public Shadows Property Parent As UserNotificationContainer
        Get
            Return DirectCast(MyBase.Parent, UserNotificationContainer)
        End Get
        Set(value As UserNotificationContainer)
            MyBase.Parent = value
        End Set
    End Property
    Public Property ID As Object
        Get
            Return varID
        End Get
        Set(value As Object)
            varID = value
        End Set
    End Property
    Public Sub New(ParentContainer As UserNotificationContainer, Message As String, ID As Object, ClickAction As Action(Of UserNotification, UserNotificationEventArgs), Color As Color, RelatedElement As Object, ElementType As DatabaseElementType)
        Hide()
        varRelatedElement = RelatedElement
        varElementType = ElementType

        DoubleBuffered = True
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        UpdateStyles()
        MyBase.Parent = ParentContainer
        varDefaultColor = Color
        varHoverColor = ColorHelper.Multiply(varDefaultColor, 0.9)
        varPressColor = ColorHelper.Multiply(varDefaultColor, 0.8)
        varClickAction = ClickAction
        varID = ID
        Size = New Size(Parent.Width - 40, 60)
        With Textlab
            .Parent = Me
            .ForeColor = Color.White
            .Font = New Font(Font.FontFamily, 10)
            .TextAlign = ContentAlignment.MiddleCenter
            .AutoSize = True
            .MaximumSize = New Size(Width - 20, Height)
            .Text = Message
            .BackColor = Color.Transparent
        End With
        Left = 20
        BackColor = varDefaultColor
        Show()
    End Sub
    Private Sub TextLab_TextChanged() Handles Textlab.SizeChanged, Textlab.TextChanged
        With Textlab
            .Location = New Point(Width \ 2 - .Width \ 2, Height \ 2 - .Height \ 2)
        End With
    End Sub
    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        BackColor = varDefaultColor
        If varClickAction IsNot Nothing Then
            varClickAction.Invoke(Me, New UserNotificationEventArgs(varID))
        End If
    End Sub
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        BackColor = varPressColor
    End Sub
    Protected Overrides Sub OnMouseEnter(e As EventArgs)
        MyBase.OnMouseEnter(e)
        BackColor = varHoverColor
    End Sub
    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        If Not Textlab.ClientRectangle.Contains(Textlab.PointToClient(MousePosition)) Then
            BackColor = varDefaultColor
        End If
    End Sub
    Public Sub Close()
        Parent.CloseNotification(Me)
    End Sub
End Class

Public Class UserNotificationEventArgs
    Inherits EventArgs
    Private varID As Object
    Public ReadOnly Property ID As Object
        Get
            Return varID
        End Get
    End Property
    Public Sub New(ID As Object)
        varID = ID
    End Sub
End Class

Public Class MessageNotification
    Inherits Control
    Private SC As SynchronizationContext = SynchronizationContext.Current
    Private WithEvents JumpTimer As New Timers.Timer(16)
    Private varMessageCount As Integer = -1
    Private varLabelIn As Boolean

    Private Const NumberOfFramesIn As Integer = 20
    Private Const NumberOfFramesJump As Integer = 13
    Private CurrentFrame As Integer
    Private LoadingSurface As New PictureBox
    Private LG As LoadingGraphics(Of PictureBox)
    Private WithEvents ActualLabel As New PictureBox

    Public Sub Start()
        JumpTimer.Start()
    End Sub
    ' Function in: 4(x-0.75x^2)
    ' Function jump: 1+x-x^2
    Private Sub JumpTimer_Tick() Handles JumpTimer.Elapsed
        SC.Post(AddressOf AdjustHeight, Nothing)
    End Sub
    Private Sub AdjustHeight(State As Object)
        If Not varLabelIn Then
            CurrentFrame += 1
            Dim NewCurrentFrame As Double = CurrentFrame / NumberOfFramesIn
            Dim Multiplier As Double = 4 * (NewCurrentFrame - 0.75 * NewCurrentFrame * NewCurrentFrame)
            With ActualLabel
                .Top = Height - CInt(Multiplier * .Height)
            End With
            If CurrentFrame = NumberOfFramesIn Then
                CurrentFrame = 0
                varLabelIn = True

                LG.Spin(30, 10)
                LoadingSurface.SendToBack()
                Height += 10
                Top -= 5

                JumpTimer.Interval = 10000
            End If
        Else
            CurrentFrame += 1
            Dim NewCurrentFrame As Double = CurrentFrame / NumberOfFramesJump
            Dim Multiplier As Double = 1 + 0.8 * (NewCurrentFrame - NewCurrentFrame * NewCurrentFrame)
            With ActualLabel
                .Top = Height - CInt(Multiplier * .Height)
            End With
            If CurrentFrame = NumberOfFramesJump Then
                CurrentFrame = 0
                JumpTimer.Interval = 10000
            ElseIf CurrentFrame = 1 Then
                JumpTimer.Interval = 16
            End If
        End If
        JumpTimer.Start()
    End Sub
    Public Property MessageCount As Integer
        Get
            Return varMessageCount
        End Get
        Set(value As Integer)
            varMessageCount = value
            ActualLabel.Text = CStr(value)
        End Set
    End Property
    Public Sub New(ParentControl As Control)
        Parent = ParentControl
        DoubleBuffered = True
        Dim Bmp As Bitmap = My.Resources.Meldingsboks
        Size = New Size(Bmp.Width + 20, 40)
        Top = (Parent.Height - Height) \ 2
        With ActualLabel
            .AutoSize = False
            .BackgroundImage = Bmp
            .BackgroundImageLayout = ImageLayout.None
            .Size = Bmp.Size
            .Left = (Width - .Width) \ 2
            '.TextAlign = ContentAlignment.MiddleCenter
            .Text = " 1"
            '.Padding = Padding.Empty
            .Parent = Me
            .Top = Height
        End With
        With JumpTimer
            .AutoReset = False
        End With
        With LoadingSurface
            .Hide()
            .Parent = ActualLabel
            .Size = New Size(ActualLabel.Width, ActualLabel.Width \ 4)
            .Top = Height - .Height
            .BackColor = Color.Transparent
        End With
        LG = New LoadingGraphics(Of PictureBox)(LoadingSurface)
        LG.Stroke = 3
        LoadingSurface.SendToBack()
    End Sub
    Private Sub OnLabelResize() Handles ActualLabel.Resize
        LoadingSurface.Top = ActualLabel.Height - LoadingSurface.Height
    End Sub
End Class



Public NotInheritable Class StaffNotificationContainer
    Inherits Tab
    Private SC As SynchronizationContext = SynchronizationContext.Current
    Private varID As Integer
    Private WithEvents CloseTimer, TopTimer As New Timers.Timer(16)
    Private NotificationList As New List(Of StaffNotification)
    Private WithEvents NotificationCounter As New Label
    Private LoadingSurface As New PictureBox
    Private LG As LoadingGraphics(Of PictureBox)
    Private Header As FullWidthControl
    Private HeaderButtons(3) As PictureBox
    Private EmptyLabel As New Label
    Private varDisplayEmptyLabel As Boolean
    Private ClosingNotifications As New List(Of StaffNotification)
    Private NotificationContainer As New Control
    Private WithEvents RefreshButton As New PictureBox
    Public Event Reloaded(Sender As Object, e As EventArgs)
    Public Property HeaderText As String
        Get
            Return Header.Text
        End Get
        Set(value As String)
            Header.Text = value
        End Set
    End Property
    Public Sub Reload(Optional Args As Object = Nothing)
        Clear()
        Spin()
        RaiseEvent Reloaded(Me, EventArgs.Empty)
    End Sub
    Public Sub Clear()
        For Each N As StaffNotification In NotificationList
            N.Dispose()
        Next
        NotificationList.Clear()
        NotificationCounter.Text = "0"
    End Sub
    Public Property ID As Integer
        Get
            Return varID
        End Get
        Set(value As Integer)
            varID = value
        End Set
    End Property
    Private Sub Refresh_Enter() Handles RefreshButton.MouseEnter
        If Not LoadingSurface.Visible Then
            With RefreshButton
                .BackgroundImage = My.Resources.RefreshHover
                .BackColor = Color.FromArgb(180, 180, 180)
            End With
        End If
    End Sub
    Private Sub Refresh_Leave() Handles RefreshButton.MouseLeave
        With RefreshButton
            .BackgroundImage = My.Resources.Refresh
            .BackColor = Header.BackColor
        End With
    End Sub
    Private Sub Refresh_Click() Handles RefreshButton.Click
        If Not LoadingSurface.Visible Then
            Reload()
        End If
    End Sub
    Private Sub CloseTimer_Tick() Handles CloseTimer.Elapsed
        SC.Post(AddressOf AdjustWidth, Nothing)
    End Sub
    Private Sub TopTimer_Tick() Handles TopTimer.Elapsed
        SC.Post(AddressOf AdjustTop, Nothing)
    End Sub
    Private Sub NotificationFinished(Sender As Notification)
        NotificationContainer.Hide()
        Sender.Dispose()
    End Sub
    Public Sub ShowMessage(Message As String, Style As NotificationAppearance)
        LG.StopSpin()
        Dim Notification As New Notification(Me, Style, Message, 5, AddressOf NotificationFinished, FloatX.FillWidth, FloatY.FillHeight)
        NotificationContainer.Show()
        Notification.Display()
    End Sub
    Private Sub RemoveAt(Index As Integer)
        With NotificationList
            .RemoveAt(Index)
            NotificationCounter.Text = CStr(.Count)
        End With
    End Sub
    Private Sub AdjustTop(State As Object)
        SuspendLayout()
        Dim iLast As Integer = NotificationList.Count - 1
        Dim ProblemFound As Boolean
        With NotificationList(0)
            If .Top >= Header.Bottom + 30 Then
                ProblemFound = True
                .Top -= 10
            Else
                .Top = Header.Bottom + 20
            End If
        End With
        If iLast > 0 Then
            For i As Integer = 1 To iLast
                With NotificationList(i)
                    If .Top >= NotificationList(i - 1).Bottom + 30 Then
                        .Top -= 10
                        ProblemFound = True
                    Else
                        .Top = NotificationList(i - 1).Bottom + 20
                    End If
                End With
            Next
        End If
        ResumeLayout(True)
        If ProblemFound Then
            TopTimer.Start()
        End If
    End Sub
    Private Sub AdjustWidth(State As Object)
        If ClosingNotifications.Count > 0 Then
            Dim iLast As Integer = ClosingNotifications.Count - 1
            For i As Integer = iLast To 0 Step -1
                Dim N As StaffNotification = ClosingNotifications(i)
                With N
                    If .Width > 40 Then
                        .Width -= 40
                    Else
                        .IsClosed = True
                        ClosingNotifications.RemoveAt(i)
                    End If
                End With
            Next
            Dim nLast As Integer = NotificationList.Count - 1
            Dim StartTimer As Boolean
            For n As Integer = nLast To 0 Step -1
                Dim CurrentNotification As StaffNotification = NotificationList(n)
                With CurrentNotification
                    If .IsClosed Then
                        RemoveAt(n)
                        .Dispose()
                        StartTimer = True
                    End If
                End With
            Next
            If ClosingNotifications.Count > 0 Then
                CloseTimer.Start()
            End If
            If NotificationList.Count > 0 AndAlso StartTimer Then
                TopTimer.Start()
            End If
        End If
    End Sub
    Public Sub Spin(Optional SwitchOn As Boolean = True)
        If SwitchOn Then
            LG.Spin(30, 10)
            NotificationCounter.Text = NotificationList.Count.ToString
            DisplayEmptyLabel = False
        Else
            LG.StopSpin()
            NotificationCounter.Text = NotificationList.Count.ToString
            DisplayEmptyLabel = True
        End If
    End Sub
    Public Property DisplayEmptyLabel As Boolean
        Get
            Return varDisplayEmptyLabel
        End Get
        Set(value As Boolean)
            varDisplayEmptyLabel = value
            RefreshLabel()
        End Set
    End Property
    Private Sub RefreshLabel()
        If Not varDisplayEmptyLabel Then
            EmptyLabel.Hide()
        End If
        If CInt(NotificationCounter.Text) > 0 Then
            NotificationCounter.BackColor = Color.FromArgb(0, 99, 157)
            EmptyLabel.Hide()
        Else
            NotificationCounter.BackColor = Color.FromArgb(180, 180, 180)
            If varDisplayEmptyLabel Then
                EmptyLabel.Show()
            End If
        End If
    End Sub
    Private Sub CounterChanged() Handles NotificationCounter.TextChanged
        RefreshLabel()
    End Sub
    Public Sub New(ParentWindow As MultiTabWindow, HeaderText As String)
        MyBase.New(ParentWindow)
        Hide()
        SuspendLayout()
        Size = New Size(500, 500)
        Header = New FullWidthControl(Me)
        With Header
            .BackColor = Color.FromArgb(220, 220, 220)
            .ForeColor = Color.FromArgb(60, 60, 60)
            .Text = HeaderText
            .TextAlign = ContentAlignment.MiddleLeft
            .Padding = New Padding(50, 0, 0, 0)
        End With
        Dim ButtonSize As New Size(30, 30)
        'Dim CenterRect As New Rectangle(Point.Empty, New Size(4 * (40 + 1), 40))
        'With CenterRect
        '    .X = Header.Width \ 2 - .Width \ 2
        '    .Y = Header.Height \ 2 - .Height \ 2
        'End With
        Dim FirstLeft As Integer = Header.Width - ((HeaderButtons.Length + 1) * 31)
        For i As Integer = 0 To HeaderButtons.Length - 1
            HeaderButtons(i) = New PictureBox
            With HeaderButtons(i)
                .BackgroundImage = My.Resources.RedigerProfilIcon
                .Size = ButtonSize
                .BackgroundImageLayout = ImageLayout.Center
                .BackColor = Color.FromArgb(230, 230, 230)
                .Parent = Header
                '.Location = New Point(CenterRect.Left + (40 + 1) * i, CenterRect.Top)
                .Location = New Point(FirstLeft + (31) * i, Header.Height \ 2 - 15)
                .Tag = i
                AddHandler .Click, AddressOf HeaderButton_Click
            End With
        Next
        HeaderButtons(0).BackgroundImage = My.Resources.klokke
        HeaderButtons(1).BackgroundImage = My.Resources.EgenerklaeringIcon
        HeaderButtons(2).BackgroundImage = My.Resources.NeedleArt
        HeaderButtons(3).BackgroundImage = My.Resources.rør

        With NotificationContainer
            .Hide()
            .Parent = Me
            .Location = New Point(0, Header.Bottom)
            .Size = New Size(500, 500 - .Top)
        End With
        With RefreshButton
            .Size = New Size(30, 30)
            .BackgroundImage = My.Resources.Refresh
            .Parent = Header
            .Left = Header.Width - 30
        End With
        With NotificationCounter
            .Parent = Header
            .Size = New Size(Header.Height, Header.Height)
            .BackColor = Color.FromArgb(180, 180, 180)
            .ForeColor = Color.White
            .Font = New Font(.Font.FontFamily, 15)
            .TextAlign = ContentAlignment.MiddleCenter
            .Text = "0"
        End With
        With EmptyLabel
            .Hide()
            .Parent = Me
            .AutoSize = False
            .Size = New Size(100, 40)
            .BackColor = Color.White
            .ForeColor = Color.FromArgb(80, 80, 80)
            .Text = "Ingenting å vise"
            .TextAlign = ContentAlignment.MiddleCenter
            .Location = New Point(Width \ 2 - .Width \ 2, (Height - .Height + Header.Bottom) \ 2)
        End With
        With LoadingSurface
            .Hide()
            .Size = New Size(50, 50)
            .Parent = Me
            .Location = New Point(Width \ 2 - .Width \ 2, (Height - .Height + Header.Bottom) \ 2)
        End With
        CloseTimer.AutoReset = False
        TopTimer.AutoReset = False
        LG = New LoadingGraphics(Of PictureBox)(LoadingSurface)
        With LG
            .Stroke = 3
            .Pen.Color = Color.FromArgb(230, 50, 80)
        End With
        BackColor = Color.White
        Show()
        ResumeLayout(True)
    End Sub
    Private Sub HeaderButton_Click(Sender As Object, e As EventArgs)
        Dim SenderPic As PictureBox = DirectCast(Sender, PictureBox)
        Parent.Index = DirectCast(SenderPic.Tag, Integer)
    End Sub
    Public Sub AddNotification(Text As String, ID As Object, ClickAction As Action(Of StaffNotification, StaffNotificationEventArgs), Color As Color, RelatedElement As Object, ElementType As DatabaseElementType)
        Dim NewNotification As New StaffNotification(Me, Text, ID, ClickAction, Color, RelatedElement, ElementType)
        With NotificationList
            'If .Count > 0 Then
            '    NewNotification.Top = .Last.Bottom + 20
            'Else
            '    NewNotification.Top = Header.Bottom + 20
            'End If
            .Insert(0, NewNotification)
            NotificationList(0).Top = Header.Bottom + 20
            Dim iLast As Integer = .Count - 1
            If iLast > 0 Then
                For i As Integer = 1 To iLast
                    NotificationList(i).Top = NotificationList(i - 1).Bottom + 20
                Next
            End If
            NotificationCounter.Text = CStr(.Count)
        End With
        varDisplayEmptyLabel = True
        LG.StopSpin()
    End Sub
    Public Sub AddNotification(NewNotification As StaffNotification)
        With NotificationList
            'If .Count > 0 Then
            '    NewNotification.Top = .Last.Bottom + 20
            'Else
            '    NewNotification.Top = Header.Bottom + 20
            'End If
            .Insert(0, NewNotification)
            NotificationList(0).Top = Header.Bottom + 20
            Dim iLast As Integer = .Count - 1
            If iLast > 0 Then
                For i As Integer = 1 To iLast
                    NotificationList(i).Top = NotificationList(i - 1).Bottom + 20
                Next
            End If
            NotificationCounter.Text = CStr(.Count)
        End With
        varDisplayEmptyLabel = True
        LG.StopSpin()
    End Sub
    Protected Friend Sub CloseNotification(Sender As StaffNotification)
        With NotificationList
            Dim iLast As Integer = .Count - 1
            Dim MatchFound As Boolean
            For i As Integer = 0 To iLast
                If NotificationList(i) Is Sender Then
                    ClosingNotifications.Add(NotificationList(i))
                    MatchFound = True
                    Exit For
                End If
            Next
            If MatchFound Then
                CloseTimer.Start()
            End If
        End With
    End Sub
    Public Sub RemoveNotification(ID As Object)
        With NotificationList
            Dim iLast As Integer = .Count - 1
            Dim MatchFound As Boolean
            For i As Integer = 0 To iLast
                Dim Notification As StaffNotification = NotificationList(i)
                If Notification.ID.Equals(ID) Then
                    Notification.Dispose()
                    .RemoveAt(i)
                    MatchFound = True
                    Exit For
                End If
            Next
            If MatchFound AndAlso .Count > 0 Then
                iLast = .Count - 1
                NotificationList(0).Top = Header.Bottom + 20
                If iLast > 0 Then
                    For i As Integer = 1 To iLast
                        NotificationList(i).Top = NotificationList(i - 1).Bottom + 20
                    Next
                End If
            Else

            End If
            NotificationCounter.Text = CStr(.Count)
        End With
    End Sub
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing Then
                For Each Pic As PictureBox In HeaderButtons
                    RemoveHandler Pic.Click, AddressOf HeaderButton_Click
                Next
                LG.Dispose()
                TopTimer.Dispose()
                CloseTimer.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub
End Class

Public Enum DatabaseElementType As Integer
    None = 0
    Egenerklæring = 1
    Time = 2
End Enum

Public Class StaffNotification
    Inherits Control
    Private WithEvents Textlab As New Label
    Private varDefaultColor, varHoverColor, varPressColor As Color
    '   Private varColor As Color = Color.Green
    Private varID, varRelatedElement As Object
    Private varIsClosed, varCloseOnClick, varIsSelected As Boolean
    Private varClickAction As Action(Of StaffNotification, StaffNotificationEventArgs)
    Private varElementType As DatabaseElementType
    Private WithEvents DBC_GetInfo As New DatabaseClient(Credentials.Server, Credentials.Database, Credentials.UserID, Credentials.Password)
    Private varInfoLoaded As Boolean
    Private varRelatedDonor As Donor
    Public Property IsSelected As Boolean
        Get
            Return varIsSelected
        End Get
        Set(value As Boolean)
            varIsSelected = value
        End Set
    End Property
    Public Property DefaultColor As Color
        Get
            Return varDefaultColor
        End Get
        Set(value As Color)
            varDefaultColor = value
        End Set
    End Property
    Public Property HoverColor As Color
        Get
            Return varHoverColor
        End Get
        Set(value As Color)
            varHoverColor = value
        End Set
    End Property
    Public Property PressColor As Color
        Get
            Return varPressColor
        End Get
        Set(value As Color)
            varPressColor = value
        End Set
    End Property
    Public ReadOnly Property RelatedDonor As Donor
        Get
            Return varRelatedDonor
        End Get
    End Property
    Public ReadOnly Property InfoLoaded As Boolean
        Get
            Return varInfoLoaded
        End Get
    End Property
    Public ReadOnly Property ElementType As DatabaseElementType
        Get
            Return varElementType
        End Get
    End Property
    Public Property RelatedElement As Object
        Get
            Return varRelatedElement
        End Get
        Set(value As Object)
            varRelatedElement = value
        End Set
    End Property
    Public Property CloseOnClick As Boolean
        Get
            Return varCloseOnClick
        End Get
        Set(value As Boolean)
            varCloseOnClick = value
        End Set
    End Property
    Private Sub TextLab_MouseDown(Sender As Object, e As MouseEventArgs) Handles Textlab.MouseDown
        OnMouseDown(e)
    End Sub
    Private Sub TextLab_MouseUp(Sender As Object, e As MouseEventArgs) Handles Textlab.MouseUp
        OnMouseUp(e)
    End Sub
    Public Property IsClosed As Boolean
        Get
            Return varIsClosed
        End Get
        Set(value As Boolean)
            varIsClosed = value
            If value Then
                Hide()
            End If
        End Set
    End Property
    Public Shadows Property Parent As StaffNotificationContainer
        Get
            Return DirectCast(MyBase.Parent, StaffNotificationContainer)
        End Get
        Set(value As StaffNotificationContainer)
            MyBase.Parent = value
        End Set
    End Property
    Public Property ID As Object
        Get
            Return varID
        End Get
        Set(value As Object)
            varID = value
        End Set
    End Property
    Public Sub New(ParentContainer As StaffNotificationContainer, Message As String, ID As Object, ClickAction As Action(Of StaffNotification, StaffNotificationEventArgs), Color As Color, RelatedElement As Object, ElementType As DatabaseElementType)
        Hide()
        varRelatedElement = RelatedElement
        DoubleBuffered = True
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        UpdateStyles()
        MyBase.Parent = ParentContainer
        varDefaultColor = Color
        varHoverColor = ColorHelper.Multiply(varDefaultColor, 0.9)
        varPressColor = ColorHelper.Multiply(varDefaultColor, 0.8)
        varClickAction = ClickAction
        varID = ID
        Size = New Size(Parent.Width - 40, 60)
        With Textlab
            .Parent = Me
            .ForeColor = Color.White
            .Font = New Font(Font.FontFamily, 10)
            .TextAlign = ContentAlignment.MiddleCenter
            .AutoSize = True
            .MaximumSize = New Size(Width - 20, Height)
            .Text = Message
            .BackColor = Color.Transparent
        End With
        Left = 20
        BackColor = varDefaultColor
        Select Case ElementType
            Case DatabaseElementType.Time
                DBC_GetInfo.SQLQuery = "SELECT * FROM Blodgiver WHERE b_fodselsnr = @id LIMIT 1;"
                DBC_GetInfo.Execute({"@id"}, {DirectCast(varRelatedElement, StaffTimeliste.StaffTime).Fødselsnummer})
            Case DatabaseElementType.Egenerklæring
                DBC_GetInfo.SQLQuery = "SELECT A.* FROM Blodgiver A INNER JOIN Time B ON A.b_fodselsnr = B.b_fodselsnr WHERE B.time_id = @timeid;"
                DBC_GetInfo.Execute({"@timeid"}, {DirectCast(varRelatedElement, Egenerklæringsliste.Egenerklæring).TimeID.ToString})
        End Select
        Show()
    End Sub
    Private Sub DBC_GetInfo_Finished(Sender As Object, e As DatabaseListEventArgs) Handles DBC_GetInfo.ListLoaded
        If e.ErrorOccurred Then

        Else
            With e.Data.Rows(0)
                If IsDBNull(.Item(4)) Then
                    .Item(4) = "Ikke registrert"
                End If
                If IsDBNull(.Item(5)) Then
                    .Item(5) = "Ikke registrert"
                End If
                If IsDBNull(.Item(12)) Then
                    .Item(12) = "Ukjent"
                End If
                Dim Fødselsnummer As Long = DirectCast(.Item(0), Long)
                Dim Navn() As String = {DirectCast(.Item(1), String), DirectCast(.Item(2), String)}
                Dim Telefon() As String = {DirectCast(.Item(3), String), DirectCast(.Item(4), String), DirectCast(.Item(5), String)}
                Dim EpostOgAdresse() As String = {DirectCast(.Item(6), String), DirectCast(.Item(7), String)}
                Dim Postnummer As Integer = DirectCast(.Item(8), Integer)
                Dim KjønnEpostSMS() As Boolean = {DirectCast(.Item(9), Boolean), DirectCast(.Item(10), Boolean), DirectCast(.Item(11), Boolean)}
                Dim Blodtype As String = DirectCast(.Item(12), String)
                varRelatedDonor = New Donor(Fødselsnummer, Navn(0), Navn(1), Telefon, EpostOgAdresse(0), EpostOgAdresse(1), Blodtype, KjønnEpostSMS(1), KjønnEpostSMS(2), KjønnEpostSMS(0), Postnummer)
            End With
        End If
        DBC_GetInfo.Dispose()
    End Sub
    Private Sub TextLab_TextChanged() Handles Textlab.SizeChanged, Textlab.TextChanged
        With Textlab
            .Location = New Point(Width \ 2 - .Width \ 2, Height \ 2 - .Height \ 2)
        End With
    End Sub
    Public Sub Close()
        Parent.CloseNotification(Me)
    End Sub
    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        If Not varIsSelected Then
            BackColor = varDefaultColor
            If varClickAction IsNot Nothing Then
                varClickAction.Invoke(Me, New StaffNotificationEventArgs(varID))
            Else

            End If
            If varCloseOnClick Then
                Parent.CloseNotification(Me)
            End If
        End If
    End Sub
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        If Not varIsSelected Then
            BackColor = varPressColor
        End If
    End Sub
    Protected Overrides Sub OnMouseEnter(e As EventArgs)
        MyBase.OnMouseEnter(e)
        ' TODO: Check if selected
        If Not varIsSelected Then
            BackColor = varHoverColor
        End If
    End Sub
    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        If Not varIsSelected Then
            If Not Textlab.ClientRectangle.Contains(Textlab.PointToClient(MousePosition)) Then
                BackColor = varDefaultColor
            End If
        End If
    End Sub
End Class

Public NotInheritable Class StaffNotificationEventArgs
    Inherits EventArgs
    Private varID As Object
    Public ReadOnly Property ID As Object
        Get
            Return varID
        End Get
    End Property
    Public Sub New(ID As Object)
        varID = ID
    End Sub
End Class


Public Class Donor
    Private varFødselsnummer As Int64
    Private varFornavn, varEtternavn, varTelefon(), varEpost, varAdresse, varBlodtype As String
    Private varSendEpost, varSendSMS, varHankjønn As Boolean
    Private varPostnummer As Integer
    Public ReadOnly Property Fødselsnummer As Int64
        Get
            Return varFødselsnummer
        End Get
    End Property
    Public ReadOnly Property Postnummer As Integer
        Get
            Return varPostnummer
        End Get
    End Property
    Public ReadOnly Property SendEpost As Boolean
        Get
            Return varSendEpost
        End Get
    End Property
    Public ReadOnly Property SendSMS As Boolean
        Get
            Return varSendSMS
        End Get
    End Property
    Public ReadOnly Property Hankjønn As Boolean
        Get
            Return varHankjønn
        End Get
    End Property
    Public ReadOnly Property Fornavn As String
        Get
            Return varFornavn
        End Get
    End Property
    Public ReadOnly Property Etternavn As String
        Get
            Return varEtternavn
        End Get
    End Property
    Public ReadOnly Property Telefon As String()
        Get
            Return varTelefon
        End Get
    End Property
    Public ReadOnly Property Epost As String
        Get
            Return varEpost
        End Get
    End Property
    Public ReadOnly Property Adresse As String
        Get
            Return varAdresse
        End Get
    End Property
    Public Property Blodtype As String
        Get
            Return varBlodtype
        End Get
        Set(value As String)
            varBlodtype = value
        End Set
    End Property
    Public Sub New(Fødselsnummer As Int64, Fornavn As String, Etternavn As String, Telefon() As String, Epost As String, Adresse As String, Blodtype As String, SendEpost As Boolean, SendSMS As Boolean, Hankjønn As Boolean, Postnummer As Integer)
        varFødselsnummer = Fødselsnummer
        varFornavn = Fornavn
        varEtternavn = Etternavn
        varTelefon = Telefon
        varEpost = Epost
        varAdresse = Adresse
        varBlodtype = Blodtype
        varSendEpost = SendEpost
        varSendSMS = SendSMS
        varHankjønn = Hankjønn
        varPostnummer = Postnummer
    End Sub
End Class
