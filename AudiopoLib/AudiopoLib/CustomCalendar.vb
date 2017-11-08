Option Strict On
Option Explicit On
Option Infer Off
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

''' <summary>
''' A highly customizable calendar control. For interaction with individual calendar days, handle the Click, MouseEnter and MouseLeave events.
''' </summary>
Public NotInheritable Class CustomCalendar
    Inherits ContainerControl
    Implements IDisposable
#Region "Fields"
    Private WHeader As WeekHeader
    Private MHeader As MonthHeader
    Private Const MaxSquares As Integer = 42
    Private M As Integer = Date.Now.Month
    Private SquareArr(41) As CalendarDay
    Protected Friend DaysOfWeek(), MonthNames() As String
    Private RHeight, CWidth, varSpacingX, varSpacingY, varMargins(3) As Integer
    Private ParentContainer As Control
    Private varAutoShrink As Boolean = True, varHideEmptyRows As Boolean = True

    Private DictCustomStates As New Dictionary(Of Integer, CalendarDayStyle)
    Private AppliedStylesList As New List(Of DateStylePair)

    Private PrevStyle, CurrStyle, NextStyle As CalendarDayStyle

    Public Shadows Event MouseEnter(Sender As CalendarDay)
    Public Shadows Event MouseLeave(Sender As CalendarDay)
    Public Shadows Event Click(Sender As CalendarDay)

    Private IsDisplayed As Boolean
#End Region
#Region "Classes"
    Private NotInheritable Class DateStylePair
        Private varDate As Date
        Private varStyleKey As Integer
        Public Property MyDate As Date
            Get
                Return varDate
            End Get
            Set(value As Date)
                varDate = value
            End Set
        End Property
        Public Property StyleKey As Integer
            Get
                Return varStyleKey
            End Get
            Set(value As Integer)
                varStyleKey = value
            End Set
        End Property
        Public Sub New(MyDate As Date, StyleKey As Integer)
            varDate = MyDate
            varStyleKey = StyleKey
        End Sub
    End Class
    Private NotInheritable Class WeekHeader
        Implements IDisposable
        Private WeekDays() As Label = {New Label, New Label, New Label, New Label, New Label, New Label, New Label}
        Private WeekDaysString() As String = {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"}
        Private WithEvents ParentContainer As CustomCalendar
        Private H As Integer
        Private ColWidth As Integer
        Private Sub ParentBGChanged() Handles ParentContainer.BackColorChanged
            Dim BG As Color = ParentContainer.BackColor
            For Each Lab As Label In WeekDays
                Lab.BackColor = BG
            Next
        End Sub
        Public Property Parent As CustomCalendar
            Get
                Return ParentContainer
            End Get
            Set(value As CustomCalendar)
                ParentContainer = value
                For i As Integer = 0 To 6
                    WeekDays(i).Parent = value
                Next
                SetDayNames()
            End Set
        End Property
        Public ReadOnly Property Height As Integer
            Get
                Return H
            End Get
        End Property
        ''' <summary>
        ''' Gets the header label displaying the name of the specified day of the week.
        ''' </summary>
        ''' <param name="DayOfTheWeek">Zero-based index of the day of the week (Monday = 0)</param>
        Public Function GetLabel(ByVal DayOfTheWeek As Integer) As Label
            Return WeekDays(DayOfTheWeek)
        End Function
        Public Sub SetDayNames()
            WeekDaysString = ParentContainer.DaysOfWeek
            For i As Integer = 1 To 6
                WeekDays(i - 1).Text = WeekDaysString(i)
            Next
            WeekDays(6).Text = WeekDaysString(0)
        End Sub
        Public Sub New(Parent As CustomCalendar, Left As Integer, Top As Integer, Spacing As Integer, Optional Width As Integer = 100, Optional ByVal Height As Integer = 100)
            ColWidth = Width
            H = Height
            For i As Integer = 0 To 6
                With WeekDays(i)
                    .Parent = Parent
                    .AutoSize = False
                    .TextAlign = ContentAlignment.MiddleCenter
                    .BackColor = Parent.BackColor
                    .Text = WeekDaysString(i)
                    .Width = ColWidth
                    .Left = Left + i * (Width + Spacing)
                    .Top = Top
                End With
            Next
            ParentContainer = Parent
        End Sub
#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    For i As Integer = 0 To 6
                        WeekDays(i).Dispose()
                        ParentContainer = Nothing
                    Next
                End If
            End If
            disposedValue = True
        End Sub
        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
        End Sub
#End Region
    End Class
    Private NotInheritable Class MonthHeader
        Implements IDisposable
        Protected Friend Enum MonthArrow
            Left = 0
            Right = 1
        End Enum
#Region "Fields"
        Private MLabel As Label
        Private WithEvents ArrowLeft, ArrowRight As PictureBox
        Private ParentControl As CustomCalendar
        Private PointsL(), PointsR() As Point
        Private CArrowDefault As Color = Color.FromArgb(0, 173, 185)
        Private CArrowHover As Color = Color.FromArgb(0, 123, 135)
        Private BrushL, BrushR As Brush
        Public Event ArrowClicked(ByVal Arrow As MonthArrow)
#End Region
        Public Property ArrowColorDefault As Color
            Get
                Return CArrowDefault
            End Get
            Set(value As Color)
                CArrowDefault = value
                BrushL.Dispose()
                BrushR.Dispose()
                BrushL = New SolidBrush(value)
                BrushR = New SolidBrush(value)
                ArrowLeft.Invalidate()
                ArrowRight.Invalidate()
            End Set
        End Property
        Public Property ArrowColorHover As Color
            Get
                Return CArrowHover
            End Get
            Set(value As Color)
                CArrowHover = value
            End Set
        End Property
        Public Property MonthString As String
            Get
                Return MLabel.Text
            End Get
            Set(value As String)
                MLabel.Text = value
            End Set
        End Property
        Public Sub New(Parent As CustomCalendar, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer)
            ParentControl = Parent
            BrushL = New SolidBrush(ArrowColorDefault)
            BrushR = New SolidBrush(ArrowColorDefault)
            MLabel = New Label
            ArrowLeft = New PictureBox
            ArrowRight = New PictureBox
            With MLabel
                .Location = New Point(Left, Top)
                .Size = New Size(Width, Height)
                .TextAlign = ContentAlignment.MiddleCenter
                .BackColor = Parent.BackColor
                .Parent = Parent
                .Text = Parent.MonthNames(Date.Now.Month - 1)
            End With
            With ArrowLeft
                .Size = New Size(16, 16)
                .Location = New Point(Left + (Width \ 2) - (.Width \ 2) - 50, Top + (Height \ 2) - (.Height \ 2))
                .Parent = Parent
                .BringToFront()
            End With
            With ArrowRight
                .Size = New Size(16, 16)
                .Location = New Point(Left + (Width \ 2) - (.Width \ 2) + 50, Top + (Height \ 2) - (.Height \ 2))
                .Parent = Parent
                .BringToFront()
            End With
            PointsL = {New Point(ArrowLeft.Width - 4, 0), New Point(ArrowLeft.Width - 4, ArrowLeft.Height), New Point(4, (ArrowLeft.Height \ 2))}
            PointsR = {New Point(4, 0), New Point(4, ArrowRight.Height), New Point(12, (ArrowRight.Height \ 2))}
        End Sub
        Private Sub PaintLeftArrow(Sender As Object, e As PaintEventArgs) Handles ArrowLeft.Paint
            e.Graphics.FillPolygon(BrushL, PointsL)
        End Sub
        Private Sub PaintRightArrow(Sender As Object, e As PaintEventArgs) Handles ArrowRight.Paint
            e.Graphics.FillPolygon(BrushR, PointsR)
        End Sub
        Private Sub EnterLeft() Handles ArrowLeft.MouseEnter
            BrushL.Dispose()
            BrushL = New SolidBrush(CArrowHover)
            ArrowLeft.Refresh()
        End Sub
        Private Sub LeaveLeft() Handles ArrowLeft.MouseLeave
            BrushL.Dispose()
            BrushL = New SolidBrush(CArrowDefault)
            ArrowLeft.Refresh()
        End Sub
        Private Sub EnterRight() Handles ArrowRight.MouseEnter
            BrushR.Dispose()
            BrushR = New SolidBrush(CArrowHover)
            ArrowRight.Refresh()
        End Sub
        Private Sub LeaveRight() Handles ArrowRight.MouseLeave
            BrushR.Dispose()
            BrushR = New SolidBrush(CArrowDefault)
            ArrowRight.Refresh()
        End Sub
        Private Sub ClickLeft() Handles ArrowLeft.MouseDown
            RaiseEvent ArrowClicked(MonthArrow.Left)
        End Sub
        Private Sub ClickRight() Handles ArrowRight.MouseDown
            RaiseEvent ArrowClicked(MonthArrow.Right)
        End Sub
#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls
        ' IDisposable
        Protected Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    MLabel.Dispose()
                    ArrowLeft.Dispose()
                    ArrowRight.Dispose()
                    BrushL.Dispose()
                    BrushR.Dispose()
                End If
            End If
            disposedValue = True
        End Sub
        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
        End Sub
#End Region
    End Class
    Public NotInheritable Class CalendarDay
        Inherits Label
        Private LabWeek As New Label
        Private GB As LinearGradientBrush
        Private myDate As Date
        Private DrawGrad As Boolean
        Private varAppliedStyle As CalendarDayStyle
        Public Area As CalendarArea = CalendarArea.Undefined
        Public Property LastStyleApplied As CalendarDayStyle
            Get
                Return varAppliedStyle
            End Get
            Set(value As CalendarDayStyle)
                varAppliedStyle = value
            End Set
        End Property
        Public Sub SetColors(ByRef Style As CalendarDayStyle)
            varAppliedStyle = Style
            If DrawGrad Then
                If GB IsNot Nothing Then
                    GB.Dispose()
                End If
                GB = New LinearGradientBrush(DisplayRectangle, Color.FromArgb(70, ColorHelper.FillRemainingRGB(Style.BackgroundColor, 0.7)), Color.Transparent, LinearGradientMode.Vertical)
            End If
            With Style
                MyBase.BackColor = .BackgroundColor
                LabWeek.BackColor = .BackgroundColor
                LabWeek.ForeColor = .DayOfWeekColor
                ForeColor = .DayColor
            End With
        End Sub
        Public Shadows Property Parent As CustomCalendar
            Get
                Return DirectCast(MyBase.Parent, CustomCalendar)
            End Get
            Set(Parent As CustomCalendar)
                MyBase.Parent = Parent
            End Set
        End Property
        Public Property Day As Date
            Get
                Return myDate
            End Get
            Set(value As Date)
                myDate = value
                Text = CStr(value.Day)
                Dim DayToUpper As String = Parent.DaysOfWeek(value.DayOfWeek).ToUpper
                If LabWeek.Text <> DayToUpper Then
                    LabWeek.Text = DayToUpper
                End If
            End Set
        End Property
        Public Overrides Property BackColor As Color
            Get
                Return MyBase.BackColor
            End Get
            Set(value As Color)
                If MyBase.BackColor <> value Then
                    If DrawGrad Then
                        If GB IsNot Nothing Then
                            GB.Dispose()
                        End If
                        GB = New LinearGradientBrush(DisplayRectangle, Color.FromArgb(70, ColorHelper.FillRemainingRGB(value, 0.7)), Color.Transparent, LinearGradientMode.Vertical)
                    End If
                    MyBase.BackColor = value
                    LabWeek.BackColor = BackColor
                End If
            End Set
        End Property
        Public Property DayOfWeekColor As Color
            Get
                Return LabWeek.ForeColor
            End Get
            Set(value As Color)
                If LabWeek.ForeColor <> value Then
                    LabWeek.ForeColor = value
                End If
            End Set
        End Property
        Public Property DrawGradient As Boolean
            Get
                Return DrawGrad
            End Get
            Set(value As Boolean)
                If value <> DrawGrad Then
                    DrawGrad = value
                    If DrawGrad Then
                        If GB IsNot Nothing Then
                            GB.Dispose()
                        End If
                        GB = New LinearGradientBrush(DisplayRectangle, Color.FromArgb(70, ColorHelper.FillRemainingRGB(BackColor, 0.7)), Color.Transparent, LinearGradientMode.Vertical)
                    End If
                    Invalidate()
                End If
            End Set
        End Property
        Public Sub New(Width As Integer, Height As Integer)
            SuspendLayout()
            ResizeRedraw = False
            SetStyle(ControlStyles.UserPaint, True)
            SetStyle(ControlStyles.AllPaintingInWmPaint, True)
            SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
            DoubleBuffered = True
            AutoSize = False
            With Me
                .Size = New Size(Width, Height)
                .TextAlign = ContentAlignment.MiddleCenter
                .Font = New Font(.Font.FontFamily, 20, FontStyle.Bold)
            End With
            With LabWeek
                .Parent = Me
                .BackColor = BackColor
                .AutoSize = False
                .Size = New Size(Width, Height \ 4)
                .TextAlign = ContentAlignment.MiddleCenter
                .Font = New Font(.Font.FontFamily, 6, FontStyle.Bold)
                .BringToFront()
            End With
            ResumeLayout()
        End Sub
        Protected Overrides Sub OnPaint(e As PaintEventArgs)
            If DrawGrad Then
                e.Graphics.FillRectangle(GB, ClientRectangle)
            End If
            MyBase.OnPaint(e)
        End Sub
        Protected Overloads Overrides Sub Dispose(disposing As Boolean)
            LabWeek.Dispose()
            If GB IsNot Nothing Then
                GB.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub
    End Class
    Public Enum CalendarArea
        Undefined = 0
        PreviousMonth = 1
        CurrentMonth = 2
        NextMonth = 3
    End Enum
    Public Class CalendarDayStyle
        Private DColor, BGColor, WColor As Color
        Public Sub New(ByVal DayColor As Color, ByVal BackgroundColor As Color, ByVal DayOfWeekColor As Color)
            DColor = DayColor
            BGColor = BackgroundColor
            WColor = DayOfWeekColor
        End Sub

        Public Property DayColor As Color
            Get
                Return DColor
            End Get
            Set(value As Color)
                DColor = value
            End Set
        End Property
        Public Property BackgroundColor As Color
            Get
                Return BGColor
            End Get
            Set(value As Color)
                BGColor = value
            End Set
        End Property
        Public Property DayOfWeekColor As Color
            Get
                Return WColor
            End Get
            Set(value As Color)
                WColor = value
            End Set
        End Property
    End Class
#End Region
    ''' <summary>
    ''' Replaces the default names for Monday through Sunday with the specified string array.
    ''' </summary>
    ''' <param name="Names">An array of string representing the days of the week, starting with Sunday and ending with Saturday.</param>
    ''' <param name="Refresh">Specifies whether or not to call the RefreshCalendar method. Specify False if several methods with this optional parameter are called in succession</param>
    Public Sub SetDayNames(Names() As String, Optional Refresh As Boolean = True)
        DaysOfWeek = Names
        If Refresh Then
            RefreshCalendar()
        End If
    End Sub
    ''' <summary>
    ''' Applies the specified custom style (that has been added using the 
    ''' </summary>
    ''' <param name="Dates"></param>
    ''' <param name="StyleKey"></param>
    ''' <param name="Refresh">Specifies whether or not to call the RefreshCalendar method. Specify False if several methods with this optional parameter are called in succession</param>
    Public Sub ApplyCustomStyle(Dates() As Date, StyleKey As Integer, Optional Refresh As Boolean = True)
        For Each D As Date In Dates
            AppliedStylesList.Add(New DateStylePair(D, StyleKey))
        Next
        If Refresh Then
            RefreshCalendar()
        End If
    End Sub
    ''' <summary>
    ''' Use this method if all dates with a given custom style should be reset to the default style.
    ''' </summary>
    ''' <param name="StyleKey"></param>
    ''' <param name="RemoveFromDictionary">If true, removes the custom style from the dictionary. If the style is to be used again, it must be added again.</param>
    ''' <param name="Refresh">Specifies whether or not to call the RefreshCalendar method. Specify False if several methods with this optional parameter are called in succession</param>
    Public Overloads Sub RemoveCustomStyle(StyleKey As Integer, Optional RemoveFromDictionary As Boolean = False, Optional Refresh As Boolean = True)
        Dim iLast As Integer = AppliedStylesList.Count - 1
        For i As Integer = iLast To 0 Step -1
            With AppliedStylesList(i)
                If .StyleKey = StyleKey Then
                    AppliedStylesList.RemoveAt(i)
                End If
            End With
        Next
        If RemoveFromDictionary Then
            DictCustomStates.Remove(StyleKey)
        End If
        If Refresh Then
            RefreshCalendar()
        End If
    End Sub
    ''' <summary>
    ''' Use this (slow) method if only some dates with a given style should be reset to the default style.
    ''' </summary>
    ''' <param name="Dates">The dates that have a custom style applied to them that should be removed.</param>
    '''   ''' <param name="Refresh">Specifies whether or not to call the RefreshCalendar method. Specify False if several methods with this optional parameter are called in succession</param>
    Public Overloads Sub RemoveCustomStyle(Dates() As Date, Optional Refresh As Boolean = True)
        Dim DatesList As List(Of Date) = Dates.ToList
        DatesList.Capacity = Dates.Length
        Dim Matches As List(Of DateStylePair) = AppliedStylesList.FindAll(Function(Pair As DateStylePair) As Boolean
                                                                              Dim iLast As Integer = DatesList.Count - 1
                                                                              For i As Integer = 0 To iLast
                                                                                  With Pair
                                                                                      If .MyDate.Date = DatesList(i).Date Then
                                                                                          DatesList.RemoveAt(i)
                                                                                          Return True
                                                                                          Exit For
                                                                                      End If
                                                                                  End With
                                                                              Next
                                                                              Return False
                                                                          End Function)
        If Matches IsNot Nothing Then
            For Each Match As DateStylePair In Matches
                Match = Nothing
            Next
        End If
        Dim nLast As Integer = AppliedStylesList.Count - 1
        If nLast >= 0 Then
            For n As Integer = nLast To 0 Step -1
                If AppliedStylesList(n) Is Nothing Then
                    AppliedStylesList.RemoveAt(n)
                End If
            Next
        End If
        If Refresh Then
            RefreshCalendar()
        End If
    End Sub
    ''' <summary>
    ''' Shows the calendar without calling the Display method. Stylistic changes to the calendar only get applied if the Display method has been shown.
    ''' </summary>
    ''' <param name="Refresh">Specifies whether or not to call the RefreshCalendar method. Specify True if stylistic changes have been made to the calendar while it was hidden.</param>
    ''' <param name="SuppressDisplay">Specifies whether or not to display the form if it has not already been displayed (not recommended).</param>
    Public Shadows Sub Show(Refresh As Boolean, Optional SuppressDisplay As Boolean = False)
        If IsDisplayed Then
            If Refresh Then
                RefreshCalendar()
            End If
            MyBase.Show()
        ElseIf Not SuppressDisplay Then
            Display()
        End If
    End Sub
    ''' <summary>
    ''' If the Display method has not yet been called, displays the calendar. If the Display method has been called, shows the calendar without calling RefreshCalendar.
    ''' </summary>
    Public Shadows Sub Show()
        If IsDisplayed Then
            MyBase.Show()
        Else
            Display()
        End If
    End Sub
    ''' <summary>
    ''' Applies changes and displays the calendar. Call this method the first time the calendar is shown. Changes to the calendar may not display correctly unless this method has been called.
    ''' </summary>
    Public Sub Display()
        IsDisplayed = True
        RefreshCalendar()
        MyBase.Show()
    End Sub
    Private Shadows Sub OnMouseEnter(Sender As Object, e As EventArgs)
        RaiseEvent MouseEnter(DirectCast(Sender, CalendarDay))
    End Sub
    Private Shadows Sub OnMouseLeave(Sender As Object, e As EventArgs)
        RaiseEvent MouseLeave(DirectCast(Sender, CalendarDay))
    End Sub
    Private Shadows Sub OnClick(Sender As Object, e As EventArgs)
        RaiseEvent Click(DirectCast(Sender, CalendarDay))
    End Sub
    ''' <summary>
    ''' Gets or sets the color of the next and previous buttons when they are not hovered.
    ''' </summary>
    Public Property ArrowColorDefault As Color
        Get
            Return MHeader.ArrowColorDefault
        End Get
        Set(value As Color)
            MHeader.ArrowColorDefault = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the color of the next and previous buttons when they are hovered.
    ''' </summary>
    Public Property ArrowColorHover As Color
        Get
            Return MHeader.ArrowColorHover
        End Get
        Set(value As Color)
            MHeader.ArrowColorHover = value
        End Set
    End Property
    ''' <summary>
    ''' Gets the number of the currently displaying month, or sets which month to display.
    ''' </summary>
    ''' <returns>An integer value between 1 and 12 (inclusive).</returns>
    Public Property CurrentMonth As Integer
        Get
            Return M
        End Get
        Set(value As Integer)
            If value >= 1 AndAlso value <= 12 Then
                M = value
                RefreshCalendar()
            End If
        End Set
    End Property

    Public Sub New(PreviousMonthStyle As CalendarDayStyle, CurrentMonthStyle As CalendarDayStyle, NextMonthStyle As CalendarDayStyle, Left As Integer, ByVal Top As Integer, Optional ByVal ColumnWidth As Integer = 100, Optional ByVal RowHeight As Integer = 100, Optional ByVal SpacingX As Integer = 20, Optional ByVal SpacingY As Integer = 20, Optional ByVal HeaderMarginBottom As Integer = 10, Optional ByVal HeaderHeight As Integer = 30, Optional WeekDays() As String = Nothing, Optional Months() As String = Nothing)
        Hide()
        Dim StylesArr(2) As CalendarDayStyle
        StylesArr(0) = PreviousMonthStyle
        StylesArr(1) = CurrentMonthStyle
        StylesArr(2) = NextMonthStyle
        Location = New Point(Left, Top)
        CommonNew(ColumnWidth, SpacingX, SpacingY, RowHeight, HeaderHeight, StylesArr, WeekDays, Months)
    End Sub
    Public Sub New(Left As Integer, ByVal Top As Integer, Optional ByVal ColumnWidth As Integer = 100, Optional ByVal RowHeight As Integer = 100, Optional ByVal SpacingX As Integer = 20, Optional ByVal SpacingY As Integer = 20, Optional ByVal HeaderMarginBottom As Integer = 10, Optional ByVal HeaderHeight As Integer = 30, Optional WeekDays() As String = Nothing, Optional Months() As String = Nothing)
        Hide()
        Dim StylesArr(2) As CalendarDayStyle
        StylesArr(0) = New CalendarDayStyle(Color.White, Color.LightGray, Color.White)
        StylesArr(1) = New CalendarDayStyle(Color.White, Color.FromArgb(0, 173, 185), Color.White)
        StylesArr(2) = New CalendarDayStyle(Color.White, Color.LightGray, Color.White)
        Location = New Point(Left, Top)
        CommonNew(ColumnWidth, SpacingX, SpacingY, RowHeight, HeaderHeight, StylesArr, WeekDays, Months)
    End Sub
    Private Sub CommonNew(ByVal ColumnWidth As Integer, ByVal SpacingX As Integer, ByVal SpacingY As Integer, ByVal RowHeight As Integer, ByVal HeaderHeight As Integer, Styles() As CalendarDayStyle, WeekDays() As String, Months() As String)
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        DoubleBuffered = True
        ResizeRedraw = False

        If WeekDays Is Nothing Then
            WeekDays = {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"}
        End If
        If Months Is Nothing Then
            Months = {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"}
        End If
        PrevStyle = Styles(0)
        CurrStyle = Styles(1)
        NextStyle = Styles(2)
        DaysOfWeek = WeekDays
        MonthNames = Months

        MHeader = New MonthHeader(Me, Left, Top, (ColumnWidth * 7) + (SpacingX * 6), 50)
        AddHandler MHeader.ArrowClicked, AddressOf OnArrowClicked
        WHeader = New WeekHeader(Me, Left, Top + 50, SpacingX, ColumnWidth, HeaderHeight) ' Legg til alle de andre parameterne også
        RHeight = RowHeight
        CWidth = ColumnWidth
        varSpacingX = SpacingX
        varSpacingY = SpacingY
        BuildSquares()
    End Sub
    ''' <summary>
    ''' Specifies whether or not the SizeToContent method should be automatically called as necessary.
    ''' </summary>
    Public Property AutoShrink As Boolean
        Get
            Return varAutoShrink
        End Get
        Set(value As Boolean)
            varAutoShrink = value
            If value Then
                SizeToContent()
            End If
        End Set
    End Property
    ''' <summary>
    ''' Sets the width and height of the calendar to the right and bottom edges of its elements.
    ''' </summary>
    Public Sub SizeToContent()
        Dim CalcHeight As Integer
        If varAutoShrink Then
            With SquareArr(41)
                If .Visible Then
                    CalcHeight = .Bottom
                Else
                    CalcHeight = SquareArr(41 - 7).Bottom
                End If
            End With
        Else
            CalcHeight = SquareArr(41).Bottom
        End If
        With Me
            .Width = SquareArr(6).Right
            .Height = CalcHeight
        End With
    End Sub
    Private Sub OnArrowClicked(Arrow As MonthHeader.MonthArrow)
        Select Case Arrow
            Case MonthHeader.MonthArrow.Left
                PreviousMonth()
            Case MonthHeader.MonthArrow.Right
                NextMonth()
        End Select
    End Sub
    ''' <summary>
    ''' Applies changes to the calendar. Call this method only if exceptional circumstances causes the calendar to display incorrectly.
    ''' </summary>
    Public Sub RefreshCalendar()
        If IsDisplayed Then
            SuspendLayout()
            Hide()
            MHeader.MonthString = MonthNames(M - 1)
            Dim D As New Date(Date.Now.Year, M, 1)
            Dim DInt As Integer = D.DayOfWeek
            If DInt = 0 Then
                DInt = 7
            End If
            Dim DaysInMonthPrevious As Integer
            If D.Month > 1 Then
                DaysInMonthPrevious = Date.DaysInMonth(D.Year, D.Month - 1)
            Else
                DaysInMonthPrevious = Date.DaysInMonth(D.Year - 1, 12)
            End If
            Dim DaysInMonthPreviousYear As Integer = Date.DaysInMonth(D.Year - 1, 12)
            Dim DaysInMonthCurrent As Integer = Date.DaysInMonth(D.Year, M)
            Dim DYear As Integer = D.Year
            Dim DMonth As Integer = D.Month
            Dim Day As Integer
            For i As Integer = 1 To MaxSquares
                If i < DInt Then
                    With SquareArr(i - 1)
                        .Area = CalendarArea.PreviousMonth
                        If D.Month > 1 Then
                            Day = DaysInMonthPrevious - DInt + i + 1
                            .Day = New Date(DYear, DMonth - 1, Day)
                        Else
                            Day = DaysInMonthPreviousYear - DInt + i + 1
                            .Day = New Date(DYear - 1, 12, Day)
                        End If
                    End With
                ElseIf i - DInt < DaysInMonthCurrent Then
                    Day = i - DInt + 1
                    With SquareArr(i - 1)
                        .Area = CalendarArea.CurrentMonth
                        .Day = New Date(DYear, DMonth, Day)
                    End With
                Else
                    Day = i - (DaysInMonthCurrent + DInt) + 1
                    With SquareArr(i - 1)
                        .Area = CalendarArea.NextMonth
                        If D.Month < 12 Then
                            .Day = New Date(DYear, DMonth + 1, Day)
                        Else
                            .Day = New Date(DYear + 1, 1, Day)
                        End If
                    End With
                End If
            Next
            Dim LastRowStart As Integer = 7 * 5
            Dim LastDayIndex As Integer = DInt + DaysInMonthCurrent - 2
            If Not SquareArr(0).Visible Then
                For i As Integer = 0 To LastRowStart - 1
                    SquareArr(i).Show()
                Next
            End If
            If LastDayIndex < LastRowStart AndAlso varHideEmptyRows Then
                For i As Integer = LastRowStart To MaxSquares - 1
                    SquareArr(i).Hide()
                Next
            ElseIf Not SquareArr(MaxSquares - 1).Visible Then
                For i As Integer = LastRowStart To MaxSquares - 1
                    SquareArr(i).Show()
                Next
            End If
            For i As Integer = 0 To 41
                With SquareArr(i)
                    Select Case .Area
                        Case CalendarArea.CurrentMonth
                            .SetColors(CurrStyle)
                        Case CalendarArea.PreviousMonth
                            .SetColors(PrevStyle)
                        Case CalendarArea.NextMonth
                            .SetColors(NextStyle)
                    End Select
                End With
            Next
            For i As Integer = 0 To 41
                With SquareArr(i)
                    Dim SquareDate As Date = .Day
                    Dim Match As DateStylePair = AppliedStylesList.Find(Function(Pair As DateStylePair) As Boolean
                                                                            With Pair.MyDate
                                                                                If .Date = SquareDate.Date Then
                                                                                    Return True
                                                                                Else
                                                                                    Return False
                                                                                End If
                                                                            End With
                                                                        End Function)
                    If Match IsNot Nothing Then
                        Dim Style As CalendarDayStyle = DictCustomStates(Match.StyleKey)
                        If Style IsNot Nothing Then
                            .SetColors(Style)
                        End If
                    End If
                End With
            Next
            WHeader.SetDayNames()
            If varAutoShrink Then
                SizeToContent()
            End If
            ResumeLayout(True)
            Show()
            Refresh()
        End If
    End Sub
    Private Sub BuildSquares()
        SuspendLayout()
        Dim WHeight As Integer = WHeader.Height
        Dim Counter As Integer = 0
        For n As Integer = 0 To 5
            For i As Integer = 0 To 6
                Dim Square As New CalendarDay(CWidth, RHeight)
                With Square
                    .Location = New Point(Left + i * (varSpacingX + CWidth), WHeight + Top + n * (varSpacingY + RHeight) + 50)
                    .Tag = (n + 1) * (i + 1)
                    .Parent = Me
                End With
                AddHandler Square.MouseEnter, AddressOf OnMouseEnter
                AddHandler Square.MouseLeave, AddressOf OnMouseLeave
                AddHandler Square.Click, AddressOf OnClick
                SquareArr(Counter) = Square
                Counter += 1
            Next
        Next
        'RefreshCalendar()
        ResumeLayout(True)
    End Sub
    ''' <summary>
    ''' Specifies whether or not to draw subtle gradients on the Calendar's days. The initial and final color are determined by the CalendarDay's BackColor property. Default = False.
    ''' </summary>
    Public WriteOnly Property DrawGradient As Boolean
        Set(value As Boolean)
            For i As Integer = 0 To 41
                SquareArr(i).DrawGradient = value
            Next
        End Set
    End Property
    Private Sub NextMonth()
        CurrentMonth += 1
    End Sub
    Private Sub PreviousMonth()
        CurrentMonth -= 1
    End Sub
    ''' <summary>
    ''' Gets the CalendarDay control currently representing the specified date.
    ''' </summary>
    ''' <returns>The CalendarDay displaying the specified date, or Nothing if the date is not included in the current view.</returns>
    Public Overloads ReadOnly Property Day(ByVal DayDate As Date) As CalendarDay
        Get
            For i As Integer = 0 To 41
                With SquareArr(i).Day
                    If .Date = DayDate.Date Then
                        Return SquareArr(i)
                    End If
                End With
            Next
            Return Nothing
        End Get
    End Property
    ''' <summary>
    ''' Gets the CalendarDay at the specified index (going left to right, including days in the previous and next month's area, if any).
    ''' </summary>
    ''' <param name="Index">An integer value between 0 and 41 inclusive.</param>
    ''' <returns>The CalendarDay at the specified index (going left to right, including days in the previous and next month's area, if any).</returns>
    Public Overloads ReadOnly Property Day(ByVal Index As Integer) As CalendarDay
        Get
            Return SquareArr(Index)
        End Get
    End Property
    ''' <summary>
    ''' Gets an array of all 42 CalendarDay controls, which include both visible and hidden days.
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Days As CalendarDay()
        Get
            Return SquareArr
        End Get
    End Property
    ''' <summary>
    ''' Adds the specified custom style to a dictionary, so that it can be applied with the specified key using the ApplyCustomStyle method.
    ''' </summary>
    ''' <param name="Key">A unique key that can be used to identify the style when it is applied.</param>
    ''' <param name="Style">The style to be applied to specific dates.</param>
    Public Sub AddCustomStyle(Key As Integer, Style As CalendarDayStyle)
        DictCustomStates.Add(Key, Style)
    End Sub
#Region "IDisposable Support"
    Private disposedValue As Boolean
    Protected Overrides Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                RemoveHandler MHeader.ArrowClicked, AddressOf OnArrowClicked
                MHeader.Dispose()
                WHeader.Dispose()
                For i As Integer = 0 To MaxSquares - 1
                    RemoveHandler SquareArr(i).MouseEnter, AddressOf OnMouseEnter
                    RemoveHandler SquareArr(i).MouseLeave, AddressOf OnMouseLeave
                    RemoveHandler SquareArr(i).Click, AddressOf OnClick
                    SquareArr(i).Dispose()
                Next
                SquareArr = Nothing
            End If
        End If
        disposedValue = True
        MyBase.Dispose(disposing)
    End Sub
#End Region
End Class