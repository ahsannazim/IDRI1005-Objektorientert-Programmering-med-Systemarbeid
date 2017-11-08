Option Strict On
Option Explicit On
Option Infer Off
Imports System.Threading
Imports System.Timers
Imports System.Windows.Forms
Imports System.Drawing

Public Class NotificationEventArgs
    Inherits EventArgs
End Class
Public Enum NotificationBackgroundTransition
    FadeIn
    GrowFromLeft
    GrowFromRight
    GrowFromTop
    GrowFromBottom
    Instant
End Enum
Public Enum NotificationTextTransition
    FadeIn
    PanFromTextAlign
    Instant
End Enum
Public Class Notification
    Inherits Label
    Private SC As SynchronizationContext = SynchronizationContext.Current
    Private nID, SMargin, CharCounter, StringWidth, Opacity, MinWidth, OffX, OffY As Integer
    Private MessageString As String = ""
    Private FadeInOutTimer, ShowTextTimer, DurationTimer As System.Timers.Timer
    Private FadingIn As Boolean = True, TextIn As Boolean = True, CenterAdjust As Boolean = True, TextTimerReadyToFinish As Boolean = False, FadeTimerReadyToFinish As Boolean = False
    Private FloatH As FloatX = FloatX.Left
    Private FloatV As FloatY = FloatY.Top
    Private FinalRGB As Color = Color.Blue
    Private LockObject As New Object()
    Private WhenFinishedAction As Action(Of Notification)
    Private BGTransition As NotificationBackgroundTransition
    Private FGTransition As NotificationTextTransition
    Public Property AdjustTextCenter As Boolean
        Get
            Return CenterAdjust
        End Get
        Set(value As Boolean)
            CenterAdjust = value
        End Set
    End Property
    Public Property FloatHorizontal As FloatX
        Get
            Return FloatH
        End Get
        Set(value As FloatX)
            FloatH = value
            AdjustPosition(Nothing, Nothing)
        End Set
    End Property
    Public Property FloatVertical As FloatY
        Get
            Return FloatV
        End Get
        Set(value As FloatY)
            FloatV = value
            AdjustPosition(Nothing, Nothing)
        End Set
    End Property

    ''' <summary>
    ''' Creates a new notification to be displayed. In most cases, this should be automatically done by a NotificationManager.
    ''' </summary>
    ''' <param name="ParentControl">The container (usually a Form) in which the notification is to appear in.</param>
    ''' <param name="Appearance">The appearance of the notification. Pass presets using the NotificationPreset class.</param>
    Public Sub New(ParentControl As Control, ByVal Appearance As NotificationAppearance, ByVal Message As String, ByVal Duration As Double, Optional WhenDone As Action(Of Notification) = Nothing, Optional ByVal AlignmentX As FloatX = FloatX.Left, Optional ByVal AlignmentY As FloatY = FloatY.Top, Optional ByVal ID As Integer = -1, Optional ByVal BackgroundTransition As NotificationBackgroundTransition = NotificationBackgroundTransition.FadeIn, Optional ByVal TextTransition As NotificationTextTransition = NotificationTextTransition.FadeIn)
        Hide()
        WhenFinishedAction = WhenDone
        ApplyAppearance(Appearance)
        FloatH = AlignmentX
        FloatV = AlignmentY
        FadeInOutTimer = New System.Timers.Timer(30)
        ShowTextTimer = New System.Timers.Timer(30)
        FadeInOutTimer.SynchronizingObject = Me
        ShowTextTimer.SynchronizingObject = Me
        ShowTextTimer.AutoReset = False
        FadeInOutTimer.AutoReset = False
        AddHandler FadeInOutTimer.Elapsed, AddressOf onTimerElapsed
        AddHandler ShowTextTimer.Elapsed, AddressOf onTextTimerElapsed

        If Duration > 0 Then
            DurationTimer = New System.Timers.Timer
            AddHandler DurationTimer.Elapsed, AddressOf onDurationTimer_Elapsed
            DurationTimer.SynchronizingObject = Me
            DurationTimer.AutoReset = False
            DurationTimer.Interval = Duration * 1000
        End If
        Parent = ParentControl
        Parent.Controls.Add(Me)
        MessageString = Message
        DoubleBuffered = True
        StringWidth = TextRenderer.MeasureText(MessageString, Font).Width
        BackColor = Color.FromArgb(0, BackColor)
        FadingIn = True
        If StringWidth + Padding.Horizontal > MinWidth - 20 Then
            Width = StringWidth + Padding.Horizontal + 20
        Else
            Width = MinWidth
        End If
    End Sub
    Private Sub ApplyAppearance(ByVal Appearance As NotificationAppearance)
        AutoSize = False
        If FGTransition = NotificationTextTransition.PanFromTextAlign Then
            ForeColor = Appearance.ForeColor
        Else
            ForeColor = Appearance.BackColor
            FinalRGB = Appearance.ForeColor
        End If
        BackColor = Appearance.BackColor
        Font = Appearance.Font
        Padding = Appearance.Padding
        TextAlign = Appearance.TextAlign
        MinWidth = Appearance.MinimumWidth
        Height = Appearance.DefaultHeight
        With Appearance
            OffX = .OffsetX
            OffY = .OffsetY
        End With
        Image = Appearance.BackgroundImage
            ImageAlign = Appearance.ImageAlign
        SMargin = Appearance.MarginSides
    End Sub
    Public Sub Display()
        AdjustPosition(Nothing, Nothing)
        If CenterAdjust = True Then
            CenterText()
        End If
        AddHandler Parent.Resize, AddressOf AdjustPosition
        Show()
        FadeInOutTimer.Start()
    End Sub
    Private Sub onDurationTimer_Elapsed(sender As Object, e As ElapsedEventArgs)
        'SC.Post(AddressOf StartFadeOut, Nothing)
        StartFadeOut(Nothing)
    End Sub
    Private Sub onTextTimerElapsed(sender As Object, e As ElapsedEventArgs)
        'SC.Send(AddressOf TextTick, Nothing)
        Select Case FGTransition
            Case NotificationTextTransition.PanFromTextAlign
                TextTick(Nothing)
            Case NotificationTextTransition.FadeIn
                FadeStep(FinalRGB.R, FinalRGB.G, FinalRGB.B, RStep, GStep, BStep)
        End Select
    End Sub
    Private Sub onTimerElapsed(sender As Object, e As ElapsedEventArgs)
        'SC.Send(AddressOf SetOpacity, Nothing)
        SetOpacity(Nothing)
    End Sub
    Private Sub SetOpacity(State As Object)
        'SyncLock LockObjectOpacity
        Dim RestartTimer As Boolean = False
        If FadingIn = True Then
            Opacity += 20
            If Opacity >= 255 Then
                'FadeInOutTimer.Stop()
                Opacity = 255
                FadingIn = False
                Select Case FGTransition
                    Case NotificationTextTransition.FadeIn
                        FadeInTextTest(BackColor, FinalRGB)
                    Case Else
                        ShowTextTimer.Start()
                End Select
                ShowTextTimer.Start()
            Else
                RestartTimer = True
            End If
            BackColor = Color.FromArgb(Opacity, BackColor)
        Else
            Opacity -= 20
            If Opacity <= 0 Then
                Opacity = 0
                BackColor = Color.FromArgb(Opacity, BackColor)
                FadingIn = True
                FadeTimerReadyToFinish = True
                If TextTimerReadyToFinish = True Then
                    Finish()
                End If
            Else
                BackColor = Color.FromArgb(Opacity, BackColor)
                RestartTimer = True
            End If
        End If
        If RestartTimer = True Then
            FadeInOutTimer.Start()
        End If
        'End SyncLock
    End Sub
    Private Sub DefaultWhenDone(State As Object)
        Dispose()
    End Sub
    Private Sub StartFadeOut(State As Object)
        'If DurationTimer IsNot Nothing Then
        '    DurationTimer.Stop()
        'End If
        RemoveHandler DurationTimer.Elapsed, AddressOf onDurationTimer_Elapsed
        If DurationTimer IsNot Nothing Then
            DurationTimer.Dispose()
        End If
        Select Case FGTransition
            Case NotificationTextTransition.FadeIn
                FadingIn = False
                TextIn = False
                FadeInTextTest(ForeColor, BackColor)
            Case Else
                FadingIn = False
                TextIn = False
                FadeInOutTimer.Start()
                ShowTextTimer.Start()
        End Select
    End Sub
    Private NewR, NewG, NewB, RStep, GStep, BStep As Double
    Private Sub FadeInTextTest(ByVal FromColor As Color, ByVal ToColor As Color, Optional ByVal FadeTime As Double = 0.5, Optional ByVal FPS As Double = 30)
        ShowTextTimer.Interval = (1000 / FPS)
        ForeColor = FromColor
        Text = MessageString
        FinalRGB = ToColor
        Dim TicksPerSecond As Integer = 30
        Dim DiffR As Integer = CInt(FromColor.R) - CInt(ToColor.R)
        Dim DiffG As Integer = CInt(FromColor.G) - CInt(ToColor.G)
        Dim DiffB As Integer = CInt(FromColor.B) - CInt(ToColor.B)
        Dim StepR As Double = -DiffR / ((1000 / ShowTextTimer.Interval) * FadeTime)
        Dim StepG As Double = -DiffG / ((1000 / ShowTextTimer.Interval) * FadeTime)
        Dim StepB As Double = -DiffB / ((1000 / ShowTextTimer.Interval) * FadeTime)
        NewR = ForeColor.R
        NewG = ForeColor.G
        NewB = ForeColor.B
        RStep = StepR
        GStep = StepG
        BStep = StepB
        ShowTextTimer.Start()
    End Sub
    Private Sub FadeStep(FinalR As Integer, FinalG As Integer, FinalB As Integer, StepR As Double, StepG As Double, StepB As Double)
        NewR += StepR
        NewG += StepG
        NewB += StepB
        Dim ColorCompleted(2) As Boolean

        If StepR < 0 Then
            If NewR < FinalR Then
                NewR = FinalR
                ColorCompleted(0) = True
            End If
        ElseIf StepR > 0 Then
            If NewR > FinalR Then
                NewR = FinalR
                ColorCompleted(0) = True
            End If
        Else
            ColorCompleted(0) = True
        End If

        If StepG < 0 Then
            If NewG < FinalG Then
                NewG = FinalG
                ColorCompleted(1) = True
            End If
        ElseIf StepG > 0 Then
            If NewG > FinalG Then
                NewG = FinalG
                ColorCompleted(1) = True
            End If
        Else
            ColorCompleted(1) = True
        End If

        If StepB < 0 Then
            If NewB < FinalB Then
                NewB = FinalB
                ColorCompleted(2) = True
            End If
        ElseIf StepB > 0 Then
            If NewB > FinalB Then
                NewB = FinalB
                ColorCompleted(2) = True
            End If
        Else
            ColorCompleted(2) = True
        End If
        ForeColor = Color.FromArgb(CInt(NewR), CInt(NewG), CInt(NewB))
        If TextIn Then
            If ColorCompleted(0) AndAlso ColorCompleted(1) AndAlso ColorCompleted(2) Then
                FinalRGB = BackColor
                With ForeColor
                    NewR = .R
                    NewG = .G
                    NewB = .B
                End With
                If DurationTimer IsNot Nothing Then
                    DurationTimer.Start()
                End If
            Else
                ShowTextTimer.Start()
            End If
        Else
            If ColorCompleted(0) AndAlso ColorCompleted(1) AndAlso ColorCompleted(2) Then
                TextTimerReadyToFinish = True
                If FadeTimerReadyToFinish Then
                    Finish()
                End If
                Text = ""
                FadeInOutTimer.Start()
            Else
                ShowTextTimer.Start()
            End If
        End If
    End Sub
    Private Sub TextTick(State As Object)
        'SyncLock LockObjectText
        Dim ResetTimer As Boolean = False
        If TextIn Then
            If CharCounter < MessageString.Length Then
                CharCounter += 1
                Text = MessageString.Substring(MessageString.Length - CharCounter, CharCounter)
                ResetTimer = True
            Else
                TextIn = False
                If DurationTimer IsNot Nothing Then
                    DurationTimer.Start()
                    FadeInTextTest(BackColor, Color.White, 0.5, 30)
                End If
            End If
        Else
            If CharCounter > 0 Then
                CharCounter -= 1
                Text = MessageString.Substring(MessageString.Length - CharCounter)
                ResetTimer = True
            Else
                TextTimerReadyToFinish = True
                If FadeTimerReadyToFinish Then
                    Finish()
                End If
            End If
        End If
        If ResetTimer = True Then
            ShowTextTimer.Start()
        End If
    End Sub
    Public Sub CenterText()
        Select Case TextAlign
            Case ContentAlignment.MiddleLeft, ContentAlignment.TopLeft, ContentAlignment.BottomLeft
                Padding = New Padding(CInt((Width - StringWidth) / 2), 0, 0, 0)
            Case ContentAlignment.MiddleRight, ContentAlignment.TopRight, ContentAlignment.BottomRight
                Padding = New Padding(0, 0, CInt((Width - StringWidth) / 2), 0)
            Case Else
                CenterAdjust = False
        End Select
    End Sub
    Public Sub AdjustPosition(sender As Object, e As EventArgs)
        Select Case FloatH
            Case FloatX.FillWidth
                Width = Parent.ClientSize.Width - SMargin * 2
                Left = OffX + SMargin
                If CenterAdjust = True Then
                    CenterText()
                End If
            Case FloatX.Left
                Left = OffX + SMargin
            Case FloatX.Right
                Left = Parent.ClientSize.Width - Width + OffX - SMargin
        End Select
        Select Case FloatV
            Case FloatY.Top
                Top = SMargin + OffY
            Case FloatY.Bottom
                Top = Parent.ClientSize.Height - Height - SMargin + OffY
            Case FloatY.FillHeight
                Height = Parent.ClientSize.Height - SMargin * 2
                Top = SMargin + OffY
        End Select
    End Sub
    Private Sub Finish()
        Try
            RemoveHandler Parent.Resize, AddressOf AdjustPosition
            RemoveHandler ShowTextTimer.Elapsed, AddressOf onTextTimerElapsed
            RemoveHandler FadeInOutTimer.Elapsed, AddressOf onTimerElapsed
            RemoveHandler DurationTimer.Elapsed, AddressOf onDurationTimer_Elapsed
            If DurationTimer IsNot Nothing Then
                DurationTimer.Dispose()
            End If
            FadeInOutTimer.Dispose()
            ShowTextTimer.Dispose()
        Catch
            If DurationTimer IsNot Nothing Then
                DurationTimer.Dispose()
            End If
            If FadeInOutTimer IsNot Nothing Then
                FadeInOutTimer.Dispose()
            End If
            If ShowTextTimer IsNot Nothing Then
                ShowTextTimer.Dispose()
            End If
        Finally
            If WhenFinishedAction IsNot Nothing Then
                WhenFinishedAction.Invoke(Me)
            End If
        End Try
    End Sub
End Class
Public Enum FloatX
    Left
    Right
    FillWidth
End Enum
Public Enum FloatY
    Top
    Bottom
    FillHeight
End Enum
Public Class NotificationPreset
    Public Shared ReadOnly Property OffRedAlert As NotificationAppearance
        Get
            Return New NotificationAppearance(Color.FromArgb(185, 6, 40), Color.White, 40, 200, 0, 0, 0)
        End Get
    End Property
    Public Shared ReadOnly Property GreenSuccess As NotificationAppearance
        Get
            Return New NotificationAppearance(Color.FromArgb(87, 179, 69), Color.White, 40, 200, 0, 0, 0)
        End Get
    End Property
    Public Shared ReadOnly Property RedAlert As NotificationAppearance
        Get
            Return New NotificationAppearance(Color.Red, Color.White, 40, 200, 0, 0, 0)
        End Get
    End Property
    Public Shared ReadOnly Property PlainWhite As NotificationAppearance
        Get
            Return New NotificationAppearance(Color.FromArgb(240, 240, 240), Color.FromArgb(50, 50, 50), 40, 100, 0, 0, 10)
        End Get
    End Property
    Public Shared ReadOnly Property BlueNotice As NotificationAppearance
        Get
            Return New NotificationAppearance(Color.FromArgb(0, 72, 83), Color.White,, 200, 0, 0, 10)
        End Get
    End Property
End Class
Public Class NotificationAppearance
#Region "Fields"
    Private BColor, FColor As Color
    Private TAlign As ContentAlignment
    Private DHeight, SMargin, MinWidth, OffX, OffY As Integer
    Private BGImage As Bitmap
    Private varFont As Font = Label.DefaultFont
    Private varPadding As Padding = New Padding(0, 0, 0, 0)
    Private ImgAlign As ContentAlignment
#End Region
    Public Sub New(Optional ByVal BackColor As Color = Nothing, Optional ByVal ForeColor As Color = Nothing, Optional ByVal DefaultHeight As Integer = 40, Optional ByVal MinimumWidth As Integer = 150, Optional ByVal OffsetX As Integer = 0, Optional ByVal OffsetY As Integer = 0, Optional ByVal MarginSides As Integer = 10, Optional ByVal TextAlign As ContentAlignment = ContentAlignment.MiddleLeft, Optional ByVal BackgroundImage As Bitmap = Nothing, Optional ByVal DefaultPadding As Padding = Nothing, Optional ByVal ImageAlign As ContentAlignment = ContentAlignment.TopLeft)
        ImgAlign = ImageAlign
        TAlign = TextAlign
        DHeight = DefaultHeight
        MinWidth = MinimumWidth
        SMargin = MarginSides
        OffX = OffsetX
        OffY = OffsetY
        BGImage = BackgroundImage
        If Not BackColor = Nothing Then
            BColor = BackColor
        End If
        If Not ForeColor = Nothing Then
            FColor = ForeColor
        End If
        If Not DefaultPadding = Nothing Then
            varPadding = DefaultPadding
        End If
    End Sub
#Region "Properties"
    Public ReadOnly Property Presets As NotificationPreset
        Get
            Return New NotificationPreset
        End Get
    End Property
    Public Property BackColor As Color
        Get
            Return BColor
        End Get
        Set(value As Color)
            BColor = value
        End Set
    End Property
    Public Property ForeColor As Color
        Get
            Return FColor
        End Get
        Set(value As Color)
            FColor = value
        End Set
    End Property
    Public Property MarginSides As Integer
        Get
            Return SMargin
        End Get
        Set(value As Integer)
            SMargin = value
        End Set
    End Property
    Public Property MinimumWidth As Integer
        Get
            Return MinWidth
        End Get
        Set(value As Integer)
            MinWidth = value
        End Set
    End Property
    Public Property DefaultHeight As Integer
        Get
            Return DHeight
        End Get
        Set(value As Integer)
            DHeight = value
        End Set
    End Property
    Public Property BackgroundImage As Bitmap
        Get
            Return BGImage
        End Get
        Set(value As Bitmap)
            BGImage = value
        End Set
    End Property
    Public Property OffsetX As Integer
        Get
            Return OffX
        End Get
        Set(value As Integer)
            OffX = value
        End Set
    End Property
    Public Property OffsetY As Integer
        Get
            Return OffY
        End Get
        Set(value As Integer)
            OffY = value
        End Set
    End Property
    Public Property TextAlign As ContentAlignment
        Get
            Return TAlign
        End Get
        Set(value As ContentAlignment)
            TAlign = value
        End Set
    End Property
    Public Property Font As Font
        Get
            Return varFont
        End Get
        Set(value As Font)
            varFont = value
        End Set
    End Property
    Public Property Padding As Padding
        Get
            Return varPadding
        End Get
        Set(value As Padding)
            varPadding = value
        End Set
    End Property
    Public Property ImageAlign As ContentAlignment
        Get
            Return ImgAlign
        End Get
        Set(value As ContentAlignment)
            ImgAlign = value
        End Set
    End Property
#End Region
End Class