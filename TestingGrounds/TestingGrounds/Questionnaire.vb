Option Strict On
Option Explicit On
Option Infer Off

Imports System.Threading
Imports System.Timers

Public Class FormNavigationButton
    Inherits Control
    Private ButtonType As FormNavigationButtonType
    Private DrawRect, UpperRect As Rectangle
    Private LowerBrush As New SolidBrush(Color.FromArgb(0, 50, 60))
    Private UpperBrush As New SolidBrush(ColorHelper.Add(Color.FromArgb(0, 50, 60), 20))
    Private ArrowBrush As New SolidBrush(ColorHelper.Add(Color.FromArgb(0, 50, 60), 100))
    Private ArrowFont As New Font(FontFamily.GenericSansSerif, 20, FontStyle.Bold)
    Private ArrowPoints() As Point
    Public Sub New(ByVal NavButtonType As FormNavigationButtonType)
        ButtonType = NavButtonType
        DrawRect = New Rectangle(New Point(0, 0), New Size(0, 0))
        UpperRect = New Rectangle(New Point(0, 0), New Size(0, 0))
    End Sub
    Private Sub RefreshGDI() Handles Me.Resize, Me.VisibleChanged
        With DrawRect
            .Width = Width - 1
            .Height = Height - 1
        End With
        Select Case ButtonType
            Case FormNavigationButtonType.PreviousButton
                Dim FirstPoint As New Point(DrawRect.Width \ 2 - 5, DrawRect.Height \ 2)
                Dim SecondPoint As New Point(DrawRect.Width \ 2 + 2, DrawRect.Height \ 2 - 7)
                Dim ThirdPoint As New Point(DrawRect.Width \ 2 + 2, DrawRect.Height \ 2 + 7)
                ArrowPoints = {FirstPoint, SecondPoint, ThirdPoint}
            Case Else
                Dim FirstPoint As New Point(DrawRect.Width \ 2 + 5, DrawRect.Height \ 2)
                Dim SecondPoint As New Point(DrawRect.Width \ 2 - 2, DrawRect.Height \ 2 - 7)
                Dim ThirdPoint As New Point(DrawRect.Width \ 2 - 2, DrawRect.Height \ 2 + 7)
                ArrowPoints = {FirstPoint, SecondPoint, ThirdPoint}
        End Select
        With UpperRect
            .Width = Width - 1
            .Height = Height \ 2
        End With
    End Sub
    Public Enum FormNavigationButtonType
        NextButton
        PreviousButton
    End Enum
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        With e.Graphics
            .FillRectangle(LowerBrush, DrawRect)
            .FillRectangle(UpperBrush, UpperRect)
            .FillPolygon(ArrowBrush, ArrowPoints)
        End With
    End Sub
End Class

Public Class QuestionnaireResult
    Private ResultsArr() As FlatFormResult
    Protected Friend Sub New(ByVal Results() As FlatFormResult)
        ResultsArr = Results
    End Sub
    Public Function Value(ByVal Form As Integer, ByVal Row As Integer, ByVal Field As Integer) As Object
        Return ResultsArr(Form).Value(Row, Field)
    End Function
    Public Function Header(ByVal Form As Integer, ByVal Row As Integer, ByVal Field As Integer) As String
        Return ResultsArr(Form).Header(Row, Field)
    End Function
    Public Function FieldType(ByVal Form As Integer, ByVal Row As Integer, ByVal Field As Integer) As FormElementType
        Return ResultsArr(Form).FieldType(Row, Field)
    End Function
End Class

Public Class Questionnaire
    Inherits Control
    Private Class FormNavigation
        Inherits Label
        Private NextButton, PrevButton As FormNavigationButton
        Private BtnSpacing As Integer
        Public Event NextClicked()
        Public Event PreviousClicked()
        Private Sub RefreshButtons() Handles Me.SizeChanged, Me.VisibleChanged
            If PrevButton IsNot Nothing Then
                With PrevButton
                    .Left = (Width - .Width - BtnSpacing) \ 2
                    .Top = 0
                End With
                With NextButton
                    .Left = (Width - .Width + BtnSpacing) \ 2
                    .BringToFront()
                    .Top = 0
                End With
            End If
        End Sub
        Public Sub New(ParentQuestionnaire As Questionnaire, ByVal Height As Integer, ByVal ButtonSpacing As Integer)
            BtnSpacing = ButtonSpacing
            With Me
                .Parent = ParentQuestionnaire
                .Left = 0
                .Top = 0
                .Height = Height 'endre til = Height
                .Width = ParentQuestionnaire.Width
                .BringToFront()
                .Show()
            End With
            PrevButton = New FormNavigationButton(FormNavigationButton.FormNavigationButtonType.PreviousButton)
            NextButton = New FormNavigationButton(FormNavigationButton.FormNavigationButtonType.NextButton)
            AddHandler PrevButton.Click, AddressOf OnPreviousClicked
            AddHandler NextButton.Click, AddressOf OnNextClicked
            With PrevButton
                .Parent = Me
                .Width = 30
                .Height = 30
                .Left = 0
                .Top = 0
            End With
            With NextButton
                .Parent = Me
                .Width = 30
                .Height = 30
                .Left = Width - .Width
                .Top = 0
            End With
            With Me
                .Show()
            End With
        End Sub
        Private Sub OnNextClicked(sender As Object, e As EventArgs)
            RaiseEvent NextClicked()
        End Sub
        Private Sub OnPreviousClicked(sender As Object, e As EventArgs)
            RaiseEvent PreviousClicked()
        End Sub
    End Class
    Private SC As SynchronizationContext = SynchronizationContext.Current
    Private FormList As List(Of FlatForm)
    Private FormNav As FormNavigation
    Private varFormIndex As Integer = -1
    Private PanTimer As Timers.Timer
    Public Function Result() As QuestionnaireResult
        Dim iLast As Integer = FormList.Count - 1
        Dim ResultArr(iLast) As FlatFormResult
        For i As Integer = 0 To FormList.Count - 1
            ResultArr(i) = FormList(i).Result
        Next
        Return New QuestionnaireResult(ResultArr)
    End Function
    Public Overloads ReadOnly Property Forms(ByVal Index As Integer) As FlatForm
        Get
            Return FormList(Index)
        End Get
    End Property
    Public Overloads ReadOnly Property Forms As List(Of FlatForm)
        Get
            Return FormList
        End Get
    End Property
    Public Property FormIndex As Integer
        Get
            Return varFormIndex
        End Get
        Set(value As Integer)
            If varFormIndex >= 0 Then
                FormList(varFormIndex).Hide()
            End If
            varFormIndex = value
            With FormList(varFormIndex)
                .Display()
                .Show()
            End With
        End Set
    End Property
    Public Sub New(Optional ByVal NavigationHeight As Integer = 40, Optional ByVal NextPreviousButtonSpacing As Integer = 100)
        Initialize(NavigationHeight, NextPreviousButtonSpacing)
    End Sub
    Public Sub New(Parent As Control, Optional ByVal NavigationHeight As Integer = 40, Optional ByVal NextPreviousButtonSpacing As Integer = 100)
        Me.Parent = Parent
        Initialize(NavigationHeight, NextPreviousButtonSpacing)
    End Sub
    Public Sub New(ByVal Rectangle As Rectangle, Optional ByVal NavigationHeight As Integer = 40, Optional ByVal NextPreviousButtonSpacing As Integer = 100)
        ' Study before deciding to correct:
        PanTimer.AutoReset = False
        With Rectangle
            Left = .Left
            Top = .Top
            Width = .Width
            Height = .Height
        End With
        Initialize(NavigationHeight, NextPreviousButtonSpacing)
    End Sub
    Public Sub New(ByVal Width As Integer, ByVal Height As Integer, Optional ByVal NavigationHeight As Integer = 40, Optional ByVal NextPreviousButtonSpacing As Integer = 100)
        With Me
            .Width = Width
            .Height = Height
        End With
        Initialize(NavigationHeight, NextPreviousButtonSpacing)
    End Sub
    Public Overloads Sub Display(Optional ByVal StartFormIndex As Integer = 0)
        Show()
        FormIndex = StartFormIndex
        RefreshNav()
    End Sub
    Private Sub Initialize(ByVal NavHeight As Integer, ByVal NavBtnSpacing As Integer)
        FormList = New List(Of FlatForm)
        Hide()
        PanTimer = New Timers.Timer(1000 / 60)
        AddHandler PanTimer.Elapsed, AddressOf PanTimerTick
        FormNav = New FormNavigation(Me, NavHeight, NavBtnSpacing)
        AddHandler FormNav.NextClicked, AddressOf OnNextClicked
        AddHandler FormNav.PreviousClicked, AddressOf OnPreviousClicked
        SetStyle(ControlStyles.SupportsTransparentBackColor, False)
    End Sub
    Private Sub OnNextClicked()
        ShowNext()
    End Sub
    Private Sub OnPreviousClicked()
        ShowPrevious()
    End Sub
    Private Sub PanToNextTest()
        PanTimer.Start()
    End Sub
    Private Sub PanTimerTick(sender As Object, e As ElapsedEventArgs)
        SC.Send(AddressOf ReduceLeft, Nothing)
    End Sub
    Private Sub ReduceLeft(State As Object)
        With FormList(varFormIndex)
            .Left -= 40
            If .Right <= 0 Then
                .Hide()
                ' TODO: Reset position
            Else
                PanTimer.Start()
            End If
        End With
    End Sub
    Public Sub Add(Form As FlatForm)
        With Form
            .Hide()
            .Parent = Me
            .Left = CInt((Width / 2) - (.Width / 2))
            .Top = CInt(((Height - FormNav.Height) / 2) - (.Height / 2))
        End With
        FormList.Add(Form)
        If varFormIndex = -1 Then
            varFormIndex = 0
        End If
    End Sub
    Private Sub RefreshNav() Handles Me.SizeChanged, Me.VisibleChanged
        If FormNav IsNot Nothing Then
            With FormNav
                .SuspendLayout()
                .Left = 0
                .Top = Height - .Height
                .Width = Width
                .ResumeLayout()
            End With
        End If
        ' Change to if list isnot nothing (if buggy)
        If varFormIndex >= 0 Then
            With FormList(varFormIndex)
                .Top = (Height - FormNav.Height) \ 2 - .Height \ 2
                .Left = CInt(Width / 2 - .Width / 2)
            End With
        End If
    End Sub
    Public Sub NewForm(ByVal Width As Integer, ByVal Height As Integer, ByVal FieldSpacing As Integer)
        Dim NF As New FlatForm(Width, Height, FieldSpacing)
        FormList.Add(NF)
    End Sub
    Public Sub ShowNext()
        If varFormIndex < FormList.Count - 1 Then
            FormIndex += 1
        End If
    End Sub
    Public Sub ShowPrevious()
        If varFormIndex > 0 Then
            FormIndex -= 1
        End If
    End Sub
End Class