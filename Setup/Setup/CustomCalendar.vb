'Option Strict On
'Option Explicit On
'Option Infer Off
'Imports System.Drawing.Drawing2D

'Public NotInheritable Class CustomCalendar
'    Implements IDisposable
'    Private NotInheritable Class WeekHeader
'        Implements IDisposable
'        Private WeekDays() As Label = {New Label, New Label, New Label, New Label, New Label, New Label, New Label}
'        Private WeekDaysString() As String = {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"}
'        Private ParentContainer As Control
'        Private H As Integer
'        Private ColWidth As Integer
'        Public ReadOnly Property Height As Integer
'            Get
'                Return H
'            End Get
'        End Property
'        ''' <summary>
'        ''' Gets the header label displaying the name of the specified day of the week.
'        ''' </summary>
'        ''' <param name="DayOfTheWeek">Zero-based index of the day of the week (Monday = 0)</param>
'        Public Function GetLabel(ByVal DayOfTheWeek As Integer) As Label
'            Return WeekDays(DayOfTheWeek)
'        End Function
'        Public Sub Refresh()
'            For i As Integer = 0 To 6
'                With WeekDays(i)
'                    .Show()
'                End With
'            Next
'        End Sub
'        Public Sub New(Parent As Control, Left As Integer, Top As Integer, Spacing As Integer, Optional Width As Integer = 100, Optional ByVal Height As Integer = 100)
'            ColWidth = Width
'            H = Height
'            For i As Integer = 0 To 6
'                WeekDays(i).Hide()
'            Next
'            For i As Integer = 0 To 6
'                With WeekDays(i)
'                    .Parent = Parent
'                    .AutoSize = False
'                    .TextAlign = ContentAlignment.MiddleCenter
'                    .BackColor = Color.Transparent
'                    .ForeColor = Color.Black
'                    .Text = WeekDaysString(i)
'                    .Width = ColWidth
'                    .Left = Left + i * (Width + Spacing)
'                    Debug.Print(Width & ", " & Spacing)
'                    .Top = Top
'                End With
'            Next
'            ParentContainer = Parent
'        End Sub
'#Region "IDisposable Support"
'        Private disposedValue As Boolean ' To detect redundant calls

'        ' IDisposable
'        Protected Sub Dispose(disposing As Boolean)
'            If Not disposedValue Then
'                If disposing Then
'                    For i As Integer = 0 To 6
'                        WeekDays(i).Dispose()
'                        ParentContainer = Nothing
'                    Next
'                End If
'            End If
'            disposedValue = True
'        End Sub
'        Public Sub Dispose() Implements IDisposable.Dispose
'            Dispose(True)
'        End Sub
'#End Region
'    End Class
'    Private NotInheritable Class MonthHeader
'        Implements IDisposable
'        Protected Friend Enum MonthArrow
'            Left = 0
'            Right = 1
'        End Enum
'#Region "Fields"
'        Private MLabel As Label
'        Private WithEvents ArrowLeft, ArrowRight As PictureBox
'        Private ParentControl As Control
'        Private PointsL(), PointsR() As Point
'        Private CArrowDefault As Color = Color.FromArgb(0, 173, 185)
'        Private CArrowHover As Color = Color.FromArgb(0, 123, 135)
'        Private BrushL, BrushR As Brush
'        Public Event ArrowClicked(ByVal Arrow As MonthArrow)
'#End Region
'        Public Property ArrowColorDefault As Color
'            Get
'                Return CArrowDefault
'            End Get
'            Set(value As Color)
'                CArrowDefault = value
'                BrushL.Dispose()
'                BrushR.Dispose()
'                BrushL = New SolidBrush(value)
'                BrushR = New SolidBrush(value)
'                ArrowLeft.Invalidate()
'                ArrowRight.Invalidate()
'            End Set
'        End Property
'        Public Property ArrowColorHover As Color
'            Get
'                Return CArrowHover
'            End Get
'            Set(value As Color)
'                CArrowHover = value
'            End Set
'        End Property
'        Public Property MonthString As String
'            Get
'                Return MLabel.Text
'            End Get
'            Set(value As String)
'                MLabel.Text = value
'            End Set
'        End Property
'        Public Sub New(Parent As Control, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer)
'            ParentControl = Parent
'            BrushL = New SolidBrush(ArrowColorDefault)
'            BrushR = New SolidBrush(ArrowColorDefault)
'            MLabel = New Label
'            ArrowLeft = New PictureBox
'            ArrowRight = New PictureBox
'            With MLabel
'                .Hide()
'                .Left = Left
'                .Top = Top
'                .Width = Width
'                .Height = Height
'                .TextAlign = ContentAlignment.MiddleCenter
'                .BackColor = Color.Transparent
'                .Parent = Parent
'                .Text = MonthName(DateTime.Now.Month)
'            End With
'            With ArrowLeft
'                .Hide()
'                .Width = 16
'                .Height = 16
'                .Left = Left + (Width \ 2) - (.Width \ 2) - 50
'                .Parent = Parent
'                .Top = Top + (Height \ 2) - (.Height \ 2)
'                .BringToFront()
'            End With
'            With ArrowRight
'                .Hide()
'                .Width = 16
'                .Height = 16
'                .Left = Left + (Width \ 2) - (.Width \ 2) + 50
'                .Parent = Parent
'                .Top = Top + (Height \ 2) - (.Height \ 2)
'                .BringToFront()
'            End With
'            PointsL = {New Point(ArrowLeft.Width - 4, 0), New Point(ArrowLeft.Width - 4, ArrowLeft.Height), New Point(4, (ArrowLeft.Height \ 2))}
'            PointsR = {New Point(4, 0), New Point(4, ArrowRight.Height), New Point(12, (ArrowRight.Height \ 2))}
'            MLabel.Show()
'            ArrowLeft.Show()
'            ArrowRight.Show()
'        End Sub
'        Private Sub PaintLeftArrow(Sender As Object, e As PaintEventArgs) Handles ArrowLeft.Paint
'            Dim SenderX As PictureBox = DirectCast(Sender, PictureBox)
'            e.Graphics.FillPolygon(BrushL, PointsL)
'        End Sub
'        Private Sub PaintRightArrow(Sender As Object, e As PaintEventArgs) Handles ArrowRight.Paint
'            Dim SenderX As PictureBox = DirectCast(Sender, PictureBox)
'            e.Graphics.FillPolygon(BrushR, PointsR)
'        End Sub
'        Private Sub EnterLeft() Handles ArrowLeft.MouseEnter
'            BrushL.Dispose()
'            BrushL = New SolidBrush(CArrowHover)
'            ArrowLeft.Invalidate()
'        End Sub
'        Private Sub LeaveLeft() Handles ArrowLeft.MouseLeave
'            BrushL.Dispose()
'            BrushL = New SolidBrush(CArrowDefault)
'            ArrowLeft.Invalidate()
'        End Sub
'        Private Sub EnterRight() Handles ArrowRight.MouseEnter
'            BrushR.Dispose()
'            BrushR = New SolidBrush(CArrowHover)
'            ArrowRight.Invalidate()
'        End Sub
'        Private Sub LeaveRight() Handles ArrowRight.MouseLeave
'            BrushR.Dispose()
'            BrushR = New SolidBrush(CArrowDefault)
'            ArrowRight.Invalidate()
'        End Sub
'        Private Sub ClickLeft() Handles ArrowLeft.MouseDown
'            RaiseEvent ArrowClicked(MonthArrow.Left)
'        End Sub
'        Private Sub ClickRight() Handles ArrowRight.MouseDown
'            RaiseEvent ArrowClicked(MonthArrow.Right)
'        End Sub
'#Region "IDisposable Support"
'        Private disposedValue As Boolean ' To detect redundant calls
'        ' IDisposable
'        Protected Sub Dispose(disposing As Boolean)
'            If Not disposedValue Then
'                If disposing Then
'                    MLabel.Dispose()
'                    ArrowLeft.Dispose()
'                    ArrowRight.Dispose()
'                    BrushL.Dispose()
'                    BrushR.Dispose()
'                End If
'            End If
'            disposedValue = True
'        End Sub
'        Public Sub Dispose() Implements IDisposable.Dispose
'            Dispose(True)
'        End Sub
'#End Region
'    End Class
'    Public NotInheritable Class CalendarDay
'        Inherits Label
'        Private LabWeek As New Label
'        Private GB As LinearGradientBrush
'        Private Shared WeekDays() As String = {"SUNDAY", "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY"}
'        Private myDate As Date
'        Private DrawGrad As Boolean = True
'        Private DrawShad As Boolean
'        Public Area As CalendarArea = CalendarArea.Undefined
'        Public Property WeekDayNames As String()
'            Get
'                Return WeekDays
'            End Get
'            Set(value As String())
'                If Not value.Length = 7 Then
'                    Throw New Exception("The array of week day names must contain 7 elements.")
'                Else
'                    WeekDays = value
'                End If
'            End Set
'        End Property
'        Public Property Day As Date
'            Get
'                Return myDate
'            End Get
'            Set(value As Date)
'                myDate = value
'                Me.Text = CStr(value.Day)
'                LabWeek.Text = WeekDays(value.DayOfWeek)
'            End Set
'        End Property
'        Public Overrides Property BackColor As Color
'            Get
'                Return MyBase.BackColor
'            End Get
'            Set(value As Color)
'                GB.Dispose()
'                GB = New LinearGradientBrush(Me.DisplayRectangle, Color.FromArgb(70, ColorHelper.FillRemainingRGB(Me.BackColor, 0.7)), Color.Transparent, LinearGradientMode.Vertical)
'                MyBase.BackColor = value
'            End Set
'        End Property
'        Public Property DayOfWeekColor As Color
'            Get
'                Return LabWeek.ForeColor
'            End Get
'            Set(value As Color)
'                LabWeek.ForeColor = value
'            End Set
'        End Property
'        Public Property DrawGradient As Boolean
'            Get
'                Return DrawGrad
'            End Get
'            Set(value As Boolean)
'                DrawGrad = value
'                Me.Invalidate()
'            End Set
'        End Property
'        Public Sub New(Width As Integer, Height As Integer)
'            Hide()
'            DoubleBuffered = True
'            LabWeek.Hide()
'            AutoSize = False
'            With Me
'                .Width = Width
'                .Height = Height
'                .TextAlign = ContentAlignment.MiddleCenter
'                '.ForeColor = Color.White
'                .Font = New Font(.Font.FontFamily, 20, FontStyle.Bold)
'            End With
'            With LabWeek
'                .Parent = Me
'                '.ForeColor = Color.White
'                .BackColor = Color.Transparent
'                .AutoSize = False
'                .Top = 0
'                .Left = 0
'                .Width = Me.Width
'                .Height = Me.Height \ 4
'                .TextAlign = ContentAlignment.MiddleCenter
'                .BringToFront()
'                .Font = New Font(.Font.FontFamily, 6, FontStyle.Bold)
'                .Show()
'            End With
'            GB = New LinearGradientBrush(Me.DisplayRectangle, Color.FromArgb(70, ColorHelper.FillRemainingRGB(Me.BackColor, 0.6)), Color.Transparent, LinearGradientMode.Vertical)
'            Debug.Print(LabWeek.Width & ", " & LabWeek.Height & ", " & LabWeek.ForeColor.R)
'        End Sub
'        Protected Overrides Sub OnPaint(e As PaintEventArgs)
'            e.Graphics.SmoothingMode = SmoothingMode.HighSpeed
'            If DrawGrad Then
'                e.Graphics.FillRectangle(GB, Me.ClientRectangle)
'            End If
'            MyBase.OnPaint(e)
'        End Sub
'        Protected Overloads Overrides Sub Dispose(disposing As Boolean)
'            LabWeek.Dispose()
'            GB.Dispose()
'            MyBase.Dispose(disposing)
'        End Sub
'    End Class
'    Public Enum CalendarArea
'        Undefined = 0
'        PreviousMonth = 1
'        CurrentMonth = 2
'        NextMonth = 3
'    End Enum
'    Public NotInheritable Class CalendarCustomState

'    End Class
'    Public Enum CalendarElement
'        DayOfWeek = 0
'        DayOfMonth = 1
'        Month = 2
'        Background = 3
'    End Enum
'    Public Enum CalendarState
'        Enabled = 0
'        Disabled = 1
'        Marked = 2
'    End Enum

'#Region "Fields"
'    Private WHeader As WeekHeader
'    Private MHeader As MonthHeader
'    Private Const MaxSquares As Integer = 42
'    Private M As Integer = DateTime.Now.Month
'    ' Private SquareList As New List(Of CalendarDay)(36)
'    Private SquareArr(42) As CalendarDay
'    Private LocationPoint As New Point()
'    Private RHeight, CWidth, varSpacingX, varSpacingY As Integer
'    Private ParentContainer As Control
'    Private AutoColors As Boolean = True
'    Private varHideEmptyRows As Boolean = True
'    Private DictCustomStates As Dictionary(Of Integer, CalendarCustomState)
'    Private PrevStyle, CurrStyle, NextStyle As CalendarDayStyle
'    Private DrawShad As Boolean
'    Public Event MouseEnter(Sender As CalendarDay)
'    Public Event MouseLeave(Sender As CalendarDay)
'    Public Event Click(Sender As CalendarDay)
'    Public WriteOnly Property DrawGradient As Boolean
'        Set(value As Boolean)
'            For i As Integer = 0 To 42
'                SquareArr(i).DrawGradient = value
'            Next
'        End Set
'    End Property
'    Public Class CalendarDayStyle
'        Private DColor, BGColor, WColor As Color
'        Public Sub New(ByVal DayColor As Color, ByVal BackgroundColor As Color, ByVal DayOfWeekColor As Color)
'            DColor = DayColor
'            BGColor = BackgroundColor
'            WColor = DayOfWeekColor
'        End Sub
'        Public Property DayColor As Color
'            Get
'                Return DColor
'            End Get
'            Set(value As Color)
'                DColor = value
'            End Set
'        End Property
'        Public Property BackgroundColor As Color
'            Get
'                Return BGColor
'            End Get
'            Set(value As Color)
'                BGColor = value
'            End Set
'        End Property
'        Public Property DayOfWeekColor As Color
'            Get
'                Return WColor
'            End Get
'            Set(value As Color)
'                WColor = value
'            End Set
'        End Property
'    End Class
'#End Region
'    Private Sub OnMouseEnter(Sender As Object, e As EventArgs)
'        RaiseEvent MouseEnter(DirectCast(Sender, CalendarDay))
'    End Sub
'    Private Sub OnMouseLeave(Sender As Object, e As EventArgs)
'        RaiseEvent MouseLeave(DirectCast(Sender, CalendarDay))
'    End Sub
'    Private Sub OnClick(Sender As Object, e As EventArgs)
'        RaiseEvent Click(DirectCast(Sender, CalendarDay))
'    End Sub
'    Public Property ArrowColorDefault As Color
'        Get
'            Return MHeader.ArrowColorDefault
'        End Get
'        Set(value As Color)
'            MHeader.ArrowColorDefault = value
'        End Set
'    End Property
'    Public Property ArrowColorHover As Color
'        Get
'            Return MHeader.ArrowColorHover
'        End Get
'        Set(value As Color)
'            MHeader.ArrowColorHover = value
'        End Set
'    End Property
'    ''' <summary>
'    ''' 
'    ''' </summary>
'    ''' <param name="Element"></param>
'    ''' <param name="DefaultColor"></param>
'    ''' <returns>Color on hover, color on disabled, color on mouse down, color on </returns>
'    Private Function AutoGenerateColor(ByVal Element As CalendarElement, DefaultColor As Color) As Color()
'        Dim ColorHover As Color
'        Dim ColorDisabled As Color
'        Dim ColorMouseDown As Color
'        Select Case Element
'            Case CalendarElement.Background
'                ColorHover = ColorHelper.Multiply(DefaultColor, -0.3)
'                ColorDisabled = Color.FromArgb(100, DefaultColor)
'                ColorMouseDown = ColorHelper.Multiply(DefaultColor, -0.5)
'            Case CalendarElement.DayOfMonth
'            Case CalendarElement.DayOfWeek
'            Case CalendarElement.Month
'        End Select
'        Return {ColorHover, ColorDisabled, ColorMouseDown}
'    End Function
'    'Public Property ColorCurrentMonth As Color
'    '    Get
'    '        Return BGCurrentMonth
'    '    End Get
'    '    Set(value As Color)
'    '        BGCurrentMonth = value
'    '        Refresh()
'    '    End Set
'    'End Property
'    Public Property Left As Integer
'        Get
'            Return LocationPoint.X
'        End Get
'        Set(value As Integer)
'            LocationPoint.X = value
'        End Set
'    End Property
'    Public Property Top As Integer
'        Get
'            Return LocationPoint.Y
'        End Get
'        Set(value As Integer)
'            LocationPoint.Y = value
'        End Set
'    End Property
'    Public Property CurrentMonth As Integer
'        Get
'            Return M
'        End Get
'        Set(value As Integer)
'            If value >= 1 AndAlso value <= 12 Then
'                M = value
'                Refresh()
'            End If
'        End Set
'    End Property
'    Public Sub New(Parent As Control, PreviousMonthStyle As CalendarDayStyle, CurrentMonthStyle As CalendarDayStyle, NextMonthStyle As CalendarDayStyle, Left As Integer, ByVal Top As Integer, Optional ByVal ColumnWidth As Integer = 100, Optional ByVal RowHeight As Integer = 100, Optional ByVal SpacingX As Integer = 20, Optional ByVal SpacingY As Integer = 20, Optional ByVal HeaderMarginBottom As Integer = 10, Optional ByVal HeaderHeight As Integer = 30)
'        'BGCurrentMonth = Color.Navy
'        LocationPoint = New Point(Left, Top)
'        ParentContainer = Parent
'        PrevStyle = PreviousMonthStyle
'        CurrStyle = CurrentMonthStyle
'        NextStyle = NextMonthStyle
'        MHeader = New MonthHeader(Parent, Left, Top, (ColumnWidth * 7) + (SpacingX * 6), 50)
'        AddHandler MHeader.ArrowClicked, AddressOf OnArrowClicked
'        WHeader = New WeekHeader(Parent, Left, Top + 50, SpacingX, ColumnWidth, HeaderHeight) ' Legg til alle de andre parameterne også
'        RHeight = RowHeight
'        CWidth = ColumnWidth
'        varSpacingX = SpacingX
'        varSpacingY = SpacingY
'        BuildSquares()
'    End Sub
'    Public Sub New(Parent As Control, Left As Integer, ByVal Top As Integer, Optional ByVal ColumnWidth As Integer = 100, Optional ByVal RowHeight As Integer = 100, Optional ByVal SpacingX As Integer = 20, Optional ByVal SpacingY As Integer = 20, Optional ByVal HeaderMarginBottom As Integer = 10, Optional ByVal HeaderHeight As Integer = 30)
'        'BGCurrentMonth = Color.Navy
'        LocationPoint = New Point(Left, Top)
'        ParentContainer = Parent
'        PrevStyle = New CalendarDayStyle(Color.White, Color.LightGray, Color.White)
'        CurrStyle = New CalendarDayStyle(Color.White, Color.FromArgb(0, 173, 185), Color.White)
'        NextStyle = New CalendarDayStyle(Color.White, Color.LightGray, Color.White)
'        MHeader = New MonthHeader(Parent, Left, Top, (ColumnWidth * 7) + (SpacingX * 6), 50)
'        AddHandler MHeader.ArrowClicked, AddressOf OnArrowClicked
'        WHeader = New WeekHeader(Parent, Left, Top + 50, SpacingX, ColumnWidth, HeaderHeight) ' Legg til alle de andre parameterne også
'        RHeight = RowHeight
'        CWidth = ColumnWidth
'        varSpacingX = SpacingX
'        varSpacingY = SpacingY
'        BuildSquares()
'    End Sub
'    Private Sub OnArrowClicked(Arrow As MonthHeader.MonthArrow)
'        Select Case Arrow
'            Case MonthHeader.MonthArrow.Left
'                PreviousMonth()
'            Case MonthHeader.MonthArrow.Right
'                NextMonth()
'        End Select
'    End Sub
'    Private Sub Refresh()
'        ParentContainer.SuspendLayout()
'        MHeader.MonthString = MonthName(M)
'        Dim D As New Date(Date.Now.Year, M, 1)
'        Dim DInt As Integer = D.DayOfWeek
'        If DInt = 0 Then
'            DInt = 7
'        End If
'        'Dim SquareListCount As Integer = SquareList.Count
'        Dim Day As Integer

'        Dim DaysInMonthPrevious As Integer
'        If D.Month > 1 Then
'            DaysInMonthPrevious = Date.DaysInMonth(D.Year, D.Month - 1)
'        Else
'            DaysInMonthPrevious = Date.DaysInMonth(D.Year - 1, 12)
'        End If
'        Dim DaysInMonthPreviousYear As Integer = Date.DaysInMonth(D.Year - 1, 12)
'        Dim DaysInMonthCurrent As Integer = Date.DaysInMonth(D.Year, M)
'        Dim DYear As Integer = D.Year
'        Dim DMonth As Integer = D.Month

'        For i As Integer = 1 To MaxSquares
'            If i < DInt Then
'                With SquareArr(i - 1)
'                    .Area = CalendarArea.PreviousMonth
'                    If D.Month > 1 Then
'                        Day = DaysInMonthPrevious - DInt + i + 1
'                        .Day = New Date(DYear, DMonth - 1, Day)
'                    Else
'                        Day = DaysInMonthPreviousYear - DInt + i + 1
'                        .Day = New Date(DYear - 1, 12, Day)
'                    End If
'                    .BackColor = PrevStyle.BackgroundColor
'                    .ForeColor = PrevStyle.DayColor
'                    Debug.Print(PrevStyle.DayColor.G.ToString)
'                    .DayOfWeekColor = PrevStyle.DayOfWeekColor
'                    Debug.Print("SETTING")
'                End With
'            ElseIf i - DInt < DaysInMonthCurrent Then
'                Day = i - DInt + 1
'                With SquareArr(i - 1)
'                    .Area = CalendarArea.CurrentMonth
'                    .Day = New Date(DYear, DMonth, Day)
'                    .BackColor = CurrStyle.BackgroundColor
'                    .ForeColor = CurrStyle.DayColor
'                    .DayOfWeekColor = CurrStyle.DayOfWeekColor
'                End With
'            Else
'                Day = i - (DaysInMonthCurrent + DInt) + 1
'                With SquareArr(i - 1)
'                    .Area = CalendarArea.NextMonth
'                    If D.Month < 12 Then
'                        .Day = New Date(DYear, DMonth + 1, Day)
'                    Else
'                        .Day = New Date(DYear + 1, 1, Day)
'                    End If
'                    .BackColor = NextStyle.BackgroundColor
'                    .ForeColor = NextStyle.DayColor
'                    .DayOfWeekColor = NextStyle.DayOfWeekColor
'                End With
'            End If
'        Next
'        Dim LastRowStart As Integer = 7 * 5
'        Dim LastDayIndex As Integer = DInt + DaysInMonthCurrent - 2
'        Debug.Print("LastRowStart: " & LastRowStart & ", LastDayIndex: " & LastDayIndex & ", days in month: " & DaysInMonthCurrent)
'        If Not SquareArr(0).Visible Then
'            For i As Integer = 0 To LastRowStart - 1
'                SquareArr(i).Show()
'            Next
'        End If
'        If LastDayIndex < LastRowStart AndAlso varHideEmptyRows Then
'            For i As Integer = LastRowStart To MaxSquares - 1
'                SquareArr(i).Hide()
'                Debug.Print("SETTING FALSE")
'            Next
'        ElseIf Not SquareArr(MaxSquares - 1).Visible Then
'            For i As Integer = LastRowStart To MaxSquares - 1
'                SquareArr(i).Show()
'            Next
'        End If
'        WHeader.Refresh()
'        ParentContainer.ResumeLayout()
'    End Sub
'    Private Sub BuildSquares()
'        ParentContainer.SuspendLayout()
'        Dim WHeight As Integer = WHeader.Height
'        Dim Counter As Integer = 0
'        For n As Integer = 0 To 5
'            For i As Integer = 0 To 6
'                Dim Square As New CalendarDay(CWidth, RHeight)
'                With Square
'                    .Hide()
'                    .Left = Left + i * (varSpacingX + CWidth)
'                    .Top = WHeight + Top + n * (varSpacingY + RHeight) + 50
'                    .Tag = (n + 1) * (i + 1)
'                End With
'                ParentContainer.Controls.Add(Square)
'                AddHandler Square.MouseEnter, AddressOf OnMouseEnter
'                AddHandler Square.MouseLeave, AddressOf OnMouseLeave
'                AddHandler Square.Click, AddressOf OnClick
'                SquareArr(Counter) = Square
'                Counter += 1
'            Next
'        Next
'        ParentContainer.ResumeLayout()
'        Refresh()
'    End Sub
'    Private Sub NextMonth()
'        CurrentMonth += 1
'    End Sub
'    Private Sub PreviousMonth()
'        CurrentMonth -= 1
'    End Sub
'#Region "IDisposable Support"
'    Private disposedValue As Boolean
'    Protected Sub Dispose(disposing As Boolean)
'        If Not disposedValue Then
'            If disposing Then
'                RemoveHandler MHeader.ArrowClicked, AddressOf OnArrowClicked
'                MHeader.Dispose()
'                WHeader.Dispose()
'                For i As Integer = 0 To MaxSquares - 1
'                    RemoveHandler SquareArr(i).MouseEnter, AddressOf OnMouseEnter
'                    RemoveHandler SquareArr(i).MouseLeave, AddressOf OnMouseLeave
'                    RemoveHandler SquareArr(i).Click, AddressOf OnClick
'                    SquareArr(i).Dispose()
'                Next
'            End If
'            SquareArr = Nothing
'        End If
'        disposedValue = True
'    End Sub
'    Public Sub Dispose() Implements IDisposable.Dispose
'        Dispose(True)
'    End Sub
'#End Region
'End Class