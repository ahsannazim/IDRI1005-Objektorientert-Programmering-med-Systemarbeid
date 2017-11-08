Option Strict On
Option Explicit On
Option Infer Off
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Text
Imports System.Windows.Forms

Public Class FormField
    Inherits Control
    Private WithEvents labHeader As Label
    Private varValue As Object = Nothing
    Private varSecondaryValue As Object = Nothing
    Private varSpacingBottom As Integer = 10
    Private varDrawGradient As Boolean
    Private varDrawBorder(3) As Boolean
    Private varDrawBorderOnHeader As Boolean = True
    Private varDrawDotsOnHeader As Boolean = True
    Protected Friend varFieldType As FormElementType
    Public Event ValueChanged(Sender As FormField, ByVal Value As Object)
    Protected Friend Event SecondaryValueChanged(Sender As FormField, ByVal Value As Object)
    Protected Event HeaderVisibleChanged(ByVal Visible As Boolean)
    Protected Friend varIsInputType As Boolean = True
    Private varTopColor, varBottomColor As Color
    Protected MeType As FormElementType
    Protected Friend GradBrush As LinearGradientBrush
    Private DashedPen As Pen
    Private varBorderPen As New Pen(Color.FromArgb(220, 220, 220))
    Private varDrawDash(3) As Boolean
    Protected varMinMax() As Integer = {0, 200}
    Protected varRequired, varIsNumeric, varIsValid As Boolean
    Public Event ValidChanged(Sender As FormField)

    Protected varRequireSpecificValue As Boolean
    Protected varRequiredValue As Object = Nothing
    Public Overloads Sub RequireSpecificValue()
        varRequireSpecificValue = False
    End Sub
    Public Overloads Sub RequireSpecificValue(Value As Object)
        varRequireSpecificValue = True
        varRequiredValue = Value
    End Sub
    Public Property IsValid As Boolean
        Get
            Return varIsValid
        End Get
        Set(value As Boolean)
            Dim ChangeOccurred As Boolean
            If value <> varIsValid Then ChangeOccurred = True
            varIsValid = value
            OnValidChanged(ChangeOccurred)
        End Set
    End Property
    Protected Overridable Sub OnValidChanged(ChangeOccurred As Boolean)
        If varIsValid Then
            varBorderPen.Color = Color.FromArgb(220, 220, 220)
        Else
            varBorderPen.Color = Color.FromArgb(255, 50, 50)
        End If
        If ChangeOccurred Then
            RaiseEvent ValidChanged(Me)
        End If
        Invalidate(True)
    End Sub
    Public Enum ElementSide
        Left = 0
        Top = 1
        Right = 2
        Bottom = 3
    End Enum
    Public Overridable Function Validate() As Boolean
        Return True
    End Function
    Protected Overridable Sub OnNumericChanged()

    End Sub
    Public Property Numeric As Boolean
        Get
            Return varIsNumeric
        End Get
        Set(value As Boolean)
            varIsNumeric = value
        End Set
    End Property
    Public Property Required As Boolean
        Get
            Return varRequired
        End Get
        Set(value As Boolean)
            varRequired = value
        End Set
    End Property
    Public Property MinLength As Integer
        Get
            Return varMinMax(0)
        End Get
        Set(value As Integer)
            varMinMax(0) = value
        End Set
    End Property
    Public Property MaxLength As Integer
        Get
            Return varMinMax(1)
        End Get
        Set(value As Integer)
            varMinMax(1) = value
            OnMaxLengthChanged()
        End Set
    End Property
    Protected Overridable Sub OnMaxLengthChanged()

    End Sub
    Protected Friend Overridable Sub Clear(ByVal ClearNonInput As Boolean)

    End Sub
    Public Property DrawBorderOnHeader As Boolean
        Get
            Return varDrawBorderOnHeader
        End Get
        Set(value As Boolean)
            varDrawBorderOnHeader = value
        End Set
    End Property
    Public Property DrawDotsOnHeader As Boolean
        Get
            Return varDrawDotsOnHeader
        End Get
        Set(value As Boolean)
            varDrawDotsOnHeader = value
        End Set
    End Property
    Private Sub OnHeaderTextAlignChanged() Handles labHeader.TextAlignChanged
        Select Case labHeader.TextAlign
            Case ContentAlignment.TopLeft
                labHeader.Padding = New Padding(6, 2, 0, 0)
            Case ContentAlignment.TopCenter
                labHeader.Padding = New Padding(0, 2, 0, 0)
            Case ContentAlignment.TopRight
                labHeader.Padding = New Padding(0, 2, 6, 0)
            Case ContentAlignment.MiddleLeft
                labHeader.Padding = New Padding(6, 0, 0, 0)
            Case ContentAlignment.MiddleCenter
                labHeader.Padding = New Padding(0, 0, 0, 0)
            Case ContentAlignment.MiddleRight
                labHeader.Padding = New Padding(0, 0, 6, 0)
            Case ContentAlignment.BottomLeft
                labHeader.Padding = New Padding(6, 0, 0, 2)
            Case ContentAlignment.BottomCenter
                labHeader.Padding = New Padding(0, 0, 0, 2)
            Case ContentAlignment.BottomRight
                labHeader.Padding = New Padding(0, 0, 6, 2)
        End Select
    End Sub
    Public ReadOnly Property BorderPen As Pen
        Get
            Return varBorderPen
        End Get
    End Property
    Public Property DrawBorder(ByVal Side As ElementSide) As Boolean
        Get
            Return varDrawBorder(Side)
        End Get
        Set(value As Boolean)
            varDrawBorder(Side) = value
            Invalidate()
        End Set
    End Property
    Public Property DrawDashedSepararators(ByVal Side As ElementSide) As Boolean
        Get
            Return varDrawDash(Side)
        End Get
        Set(value As Boolean)
            varDrawDash(Side) = value
            Invalidate()
        End Set
    End Property
    Public ReadOnly Property IsInputType As Boolean
        Get
            Return varIsInputType
        End Get
    End Property
    Public Property DrawGradient As Boolean
        Get
            Return varDrawGradient
        End Get
        Set(value As Boolean)
            varDrawGradient = value
            Invalidate()
        End Set
    End Property
    Protected Friend Property TopColor As Color
        Get
            Return varTopColor
        End Get
        Set(value As Color)
            varTopColor = value
            If GradBrush IsNot Nothing Then
                GradBrush.Dispose()
            End If
            GradBrush = New LinearGradientBrush(Point.Empty, New Point(0, Bottom), varTopColor, varBottomColor)
            Invalidate()
        End Set
    End Property
    Protected Friend Property BottomColor As Color
        Get
            Return varBottomColor
        End Get
        Set(value As Color)
            varBottomColor = value
            If GradBrush IsNot Nothing Then
                GradBrush.Dispose()
            End If
            GradBrush = New LinearGradientBrush(Point.Empty, New Point(0, Bottom), varTopColor, varBottomColor)
            Invalidate()
        End Set
    End Property
    ' TODO: Override in all derived classes for which pressing enter og space should have an effect.
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        MyBase.OnKeyDown(e)
        If e.KeyCode = Keys.Space OrElse e.KeyCode = Keys.Enter Then
            OnUserClick()
        End If
    End Sub
    Protected Overridable Sub OnUserClick()

    End Sub
    Protected Overridable Sub OnValueChanged(ByVal Value As Object)
        RaiseEvent ValueChanged(Me, Value)
        IsValid = True
    End Sub
    Protected Overridable Sub OnSecondaryValueChanged(ByVal Value As Object)
        RaiseEvent SecondaryValueChanged(Me, Value)
    End Sub
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        With e.Graphics
            .SmoothingMode = SmoothingMode.AntiAlias
            '.CompositingQuality = CompositingQuality.HighQuality
            If varDrawGradient Then
                .FillRectangle(GradBrush, ClientRectangle)
            End If
            'If varDrawDash(0) Then



            If varDrawDash(0) Then
                .DrawLine(DashedPen, Point.Empty, New Point(0, Height))
            End If
            If varDrawDash(1) Then
                .DrawLine(DashedPen, Point.Empty, New Point(Width, 0))
            End If
            If varDrawDash(2) Then
                .DrawLine(DashedPen, New Point(Width, 0), New Point(Width, Height))
            End If
            If varDrawDash(3) Then
                .DrawLine(DashedPen, New Point(0, Height), New Point(Width, Height))
            End If

            If varDrawBorder(0) Then
                .DrawLine(varBorderPen, Point.Empty, New Point(0, Height))
            End If
            If varDrawBorder(1) Then
                .DrawLine(varBorderPen, Point.Empty, New Point(Width, 0))
            End If
            If varDrawBorder(2) Then
                .DrawLine(varBorderPen, New Point(Width - 1, 0), New Point(Width - 1, Height))
            End If
            If varDrawBorder(3) Then
                .DrawLine(varBorderPen, New Point(0, Height - 1), New Point(Width, Height - 1))
            End If

            'End If
            MyBase.OnPaint(e)
        End With
    End Sub
    'Protected Overrides Sub OnGotFocus(e As EventArgs)
    '    MyBase.OnGotFocus(e)
    '    If Not MeType = FormElementType.TextField Then
    '        BackColor = ColorHelper.Multiply(BackColor, 0.4)
    '        DrawGradient = False
    '    End If
    'End Sub
    'Protected Overrides Sub OnLostFocus(e As EventArgs)
    '    MyBase.OnLostFocus(e)
    '    If Not MeType = FormElementType.TextField Then
    '        BackColor = ColorHelper.Multiply(BackColor, 2.5)
    '        DrawGradient = True
    '    End If
    'End Sub

    'To reduce flicker
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim params As CreateParams = MyBase.CreateParams
            params.ExStyle = params.ExStyle Or &H2000000
            Return params
        End Get
    End Property

    Public Sub New(ByVal FieldSpacing As Integer, ByVal FieldType As FormElementType, Style As FormFieldStyle)
        DoubleBuffered = True
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        UpdateStyles()

        'Using NH As New HatchBrush(HatchStyle.DottedGrid, ColorHelper.Add(Color.FromArgb(5, 53, 68), 20))
        Using NH As New HatchBrush(HatchStyle.DottedGrid, ColorHelper.Add(Style.HeaderBG, 20), ColorHelper.Add(Style.HeaderBG, -50))
            DashedPen = New Pen(NH)
        End Using
        Dim DrawBordersArr() As Boolean = Style.DrawBorders
        For i As Integer = 0 To 3
            varDrawBorder(i) = DrawBordersArr(i)
        Next
        varFieldType = FieldType
        varSpacingBottom = FieldSpacing
        SetStyle(ControlStyles.Selectable, False)
        labHeader = New Label
        With labHeader
            .TabStop = False
            .Location = Point.Empty
            .Width = Width
            .Height = Style.HeaderHeight
            .Font = New Font(.Font.FontFamily, 8)
            .BackColor = Style.HeaderBG
            .ForeColor = Style.HeaderFG
            .TextAlign = ContentAlignment.MiddleLeft
            .Padding = New Padding(6, 0, 0, 0)
            .Parent = Me
            AddHandler .Paint, AddressOf PaintLines
        End With
        If FieldType = FormElementType.TextField Then
            BackColor = Style.TextBG
            ForeColor = Style.TextFG
        Else
            BackColor = Style.BodyBG
            ForeColor = Style.BodyFG
        End If
        varTopColor = Color.FromArgb(50, ColorHelper.Multiply(BackColor, 1.5))
        varBottomColor = Color.FromArgb(50, BackColor)
    End Sub
    Private Sub PaintLines(Sender As Object, e As PaintEventArgs)
        Dim LabSize As Size = labHeader.ClientSize
        With e.Graphics
            If varDrawDotsOnHeader Then
                If varDrawDash(0) Then
                    .DrawLine(DashedPen, Point.Empty, New Point(0, LabSize.Height))
                End If
                If varDrawDash(1) Then
                    .DrawLine(DashedPen, Point.Empty, New Point(LabSize.Width, 0))
                End If
                If varDrawDash(2) Then
                    .DrawLine(DashedPen, New Point(LabSize.Width, 0), New Point(LabSize.Width, LabSize.Height))
                End If
            End If
            If varDrawBorderOnHeader Then
                If varDrawBorder(0) Then
                    .DrawLine(varBorderPen, Point.Empty, New Point(0, LabSize.Height))
                End If
                If varDrawBorder(1) Then
                    .DrawLine(varBorderPen, Point.Empty, New Point(LabSize.Width, 0))
                End If
                If varDrawBorder(2) Then
                    .DrawLine(varBorderPen, New Point(LabSize.Width - 1, 0), New Point(LabSize.Width - 1, LabSize.Height))
                End If
            End If
        End With
    End Sub
    Protected Overrides Sub OnSizeChanged(e As EventArgs)
        MyBase.OnSizeChanged(e)
        labHeader.Width = Width
        If GradBrush IsNot Nothing Then
            GradBrush.Dispose()
        End If
        GradBrush = New LinearGradientBrush(Point.Empty, New Point(0, Bottom), TopColor, BottomColor)
    End Sub
    Protected Overrides Sub OnVisibleChanged(e As EventArgs)
        MyBase.OnVisibleChanged(e)
        labHeader.Width = Width
    End Sub
    Public ReadOnly Property Header As Label
        Get
            Return labHeader
        End Get
    End Property
    Public Overridable Property Value As Object
        Get
            Return varValue
        End Get
        Set(value As Object)
            varValue = value
            OnValueChanged(value)
        End Set
    End Property
    Public Overridable Property SecondaryValue As Object
        Get
            Return varSecondaryValue
        End Get
        Set(value As Object)
            varSecondaryValue = value
            OnSecondaryValueChanged(value)
        End Set
    End Property
    Public Overridable Overloads Sub SwitchHeader(ByVal SwitchOn As Boolean)
        With Header
            If SwitchOn Then
                .Height = 17
                .Show()
            Else
                .Height = 0
                .Hide()
            End If
        End With
        RaiseEvent HeaderVisibleChanged(SwitchOn)
    End Sub
    Public Sub Extrude(ByVal Side As FieldExtrudeSide, ByVal Amount As Integer)
        SuspendLayout()
        Select Case Side
            Case FieldExtrudeSide.Left
                Width += Amount
                Left -= Amount
            Case FieldExtrudeSide.Right
                Width += Amount
            Case FieldExtrudeSide.Bottom
                Height += Amount
        End Select
        ResumeLayout()
    End Sub
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing Then
                If GradBrush IsNot Nothing Then
                    GradBrush.Dispose()
                End If
            End If
            MyBase.Dispose(disposing)
        Catch
        End Try
    End Sub
End Class
Public Enum FieldExtrudeSide
    Left
    Right
    Bottom
End Enum
Public Enum FormElementType
    CheckBox
    TextField
    Button
    Label
    DropDown
    Radio
    Space
End Enum
Public Class FormFieldStylePresets
    Public Shared ReadOnly Property PlainWhite As FormFieldStyle
        Get
            Return New FormFieldStyle(Color.FromArgb(245, 245, 245), Color.FromArgb(70, 70, 70), Color.White, Color.Black, Color.White, Color.Black, {True, True, True, True}, 20)
        End Get
    End Property
End Class
Public Class FormFieldStyle
    Private varHeaderBG, varHeaderFG, varBodyBG, varBodyFG, varTextBG, varTextFG As Color
    Private varDrawBorders(3) As Boolean
    Private varHeaderHeight As Integer
    Public Overloads Property DrawBorders(ByVal Side As FormField.ElementSide) As Boolean
        Get
            Return varDrawBorders(Side)
        End Get
        Set(value As Boolean)
            varDrawBorders(Side) = value
        End Set
    End Property
    Public Overloads Property DrawBorders(ByVal Side As Integer) As Boolean
        Get
            Return varDrawBorders(Side)
        End Get
        Set(value As Boolean)
            varDrawBorders(Side) = value
        End Set
    End Property
    Public Overloads Property DrawBorders As Boolean()
        Get
            Return varDrawBorders
        End Get
        Set(value As Boolean())
            varDrawBorders = value
        End Set
    End Property
    Public Property HeaderBG As Color
        Get
            Return varHeaderBG
        End Get
        Set(value As Color)
            varHeaderBG = value
        End Set
    End Property
    Public Property HeaderFG As Color
        Get
            Return varHeaderFG
        End Get
        Set(value As Color)
            varHeaderFG = value
        End Set
    End Property
    Public Property BodyBG As Color
        Get
            Return varBodyBG
        End Get
        Set(value As Color)
            varBodyBG = value
        End Set
    End Property
    Public Property BodyFG As Color
        Get
            Return varBodyFG
        End Get
        Set(value As Color)
            varBodyFG = value
        End Set
    End Property
    Public Property TextBG As Color
        Get
            Return varTextBG
        End Get
        Set(value As Color)
            varTextBG = value
        End Set
    End Property
    Public Property TextFG As Color
        Get
            Return varTextFG
        End Get
        Set(value As Color)
            varTextFG = value
        End Set
    End Property
    Public Property HeaderHeight As Integer
        Get
            Return varHeaderHeight
        End Get
        Set(value As Integer)
            varHeaderHeight = value
        End Set
    End Property
    ''' <summary>
    ''' Initializes a new FormFieldStyle with the specified colors.
    ''' </summary>
    Public Sub New(ByVal HeaderBG As Color, ByVal HeaderFG As Color, ByVal BodyBG As Color, ByVal BodyFG As Color, ByVal TextBG As Color, ByVal TextFG As Color, ByVal DrawBorders() As Boolean, ByVal HeaderHeight As Integer)
        varHeaderBG = HeaderBG
        varHeaderFG = HeaderFG
        varBodyBG = BodyBG
        varBodyFG = BodyFG
        varTextBG = TextBG
        varTextFG = TextFG
        varDrawBorders = DrawBorders
        varHeaderHeight = HeaderHeight
    End Sub
    ''' <summary>
    ''' Initializes a new FormFieldStyle with the default colors.
    ''' </summary>
    Public Sub New()
        varHeaderBG = Color.FromArgb(5, 53, 68)
        varHeaderFG = Color.White
        varBodyBG = Color.FromArgb(7, 70, 91)
        varBodyFG = Color.White
        varTextBG = ColorHelper.Multiply(Color.FromArgb(7, 70, 91), 0.4)
        varTextFG = Color.White
        varHeaderHeight = 17
    End Sub
End Class

Public Class FormElement
    Private E As FormElementType
    Private W As Integer
    Private V As Object
    Public Sub New(ByVal Type As FormElementType, Optional ByVal Width As Integer = -1, Optional ByVal Value As Object = Nothing)
        E = Type
        W = Width
        V = Value
    End Sub
    Public Property Width As Integer
        Get
            Return W
        End Get
        Set(value As Integer)
            W = value
        End Set
    End Property
    Public Property Type As FormElementType
        Get
            Return E
        End Get
        Set(value As FormElementType)
            E = value
        End Set
    End Property
    Public Property Value As Object
        Get
            Return V
        End Get
        Set(value As Object)
            V = value
        End Set
    End Property
End Class
Public Class HeaderValuePair
    Public Header As String = "Not set"
    Public Value As Object = False
    Public Type As FormElementType
    Public Sub New(ByVal HeaderString As String, ByVal ValueObject As Object, ByVal ElementType As FormElementType)
        Header = HeaderString
        Value = ValueObject
        Type = ElementType
    End Sub
End Class

Public Class FlatForm
#Region "FlatForm"
    Inherits Control
    Private varFieldSpacing As Integer
    Private RadioContextList As New List(Of RadioButtonContext)
    Private RowList As List(Of FormRow)
    Private varRowHeight As Integer = 57
    Private varNewFieldStyle As New FormFieldStyle()
    Public Event EnterPressed()
    'To reduce flicker
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim params As CreateParams = MyBase.CreateParams
            params.ExStyle = params.ExStyle Or &H2000000
            Return params
        End Get
    End Property

    Public ReadOnly Property Last As FormField
        Get
            Dim Ret As FormField
            With RowList
                With .Item(.Count - 1).Fields
                    If .Count > 0 Then
                        Ret = .Item(.Count - 1)
                    Else
                        Return Nothing
                    End If
                End With
            End With
            Return Ret
        End Get
    End Property
    Public Property NewRowHeight As Integer
        Get
            Return varRowHeight
        End Get
        Set(value As Integer)
            varRowHeight = value
        End Set
    End Property
    Public Property NewFieldStyle As FormFieldStyle
        Get
            Return varNewFieldStyle
        End Get
        Set(value As FormFieldStyle)
            varNewFieldStyle = value
        End Set
    End Property
    Public ReadOnly Property Row(ByVal Index As Integer) As FormRow
        Get
            Return RowList(Index)
        End Get
    End Property
    Public ReadOnly Property LastRow As FormRow
        Get
            With RowList
                Return .Item(.Count - 1)
            End With
        End Get
    End Property
    Public ReadOnly Property Rows As List(Of FormRow)
        Get
            Return RowList
        End Get
    End Property
    Public Function GetRadioSeries() As String
        Dim SB As New StringBuilder
        For Each Context As RadioButtonContext In RadioContextList
            Dim Checked As Integer = Context.GetChecked
            If Checked < 0 Then
                SB.Append("X")
            Else
                SB.Append(Checked)
            End If
        Next
        Return SB.ToString
    End Function
    Public ReadOnly Property Result() As HeaderValuePair()
        Get
            Dim iLast As Integer = RowList.Count - 1
            Dim RetList As New List(Of HeaderValuePair)
            For Each R As FormRow In RowList
                For Each Item As FormField In R.Fields
                    With Item
                        Select Case .varFieldType
                            Case FormElementType.CheckBox, FormElementType.Radio, FormElementType.TextField
                                RetList.Add(New HeaderValuePair(.Header.Text, .Value, .varFieldType))
                        End Select
                    End With
                Next
            Next
            Dim RetArr() As HeaderValuePair = RetList.ToArray
            RetList = Nothing
            Return RetArr
        End Get
    End Property
    Public ReadOnly Property FieldCount As Integer
        Get
            Dim Sum As Integer = 0
            If RowList.Count > 0 Then
                For i As Integer = 0 To RowList.Count - 1
                    Sum += RowList(i).Controls.Count
                Next
            End If
            Return Sum
        End Get
    End Property
    Public ReadOnly Property RowCount As Integer
        Get
            Return RowList.Count
        End Get
    End Property
    Public ReadOnly Property Field(ByVal Row As Integer, ByVal Index As Integer) As FormField
        Get
            Return RowList(Row).Fields(Index)
        End Get
    End Property
    Public Function Validate() As Boolean
        SuspendLayout()
        Dim EverythingValid As Boolean = True
        For Each R As FormRow In RowList
            If Not R.Validate Then
                EverythingValid = False
            End If
        Next
        For Each C As RadioButtonContext In RadioContextList
            If Not C.Validate Then
                EverythingValid = False
            End If
        Next
        ResumeLayout(True)
        Return EverythingValid
    End Function
    Public Sub New(ByVal Width As Integer, ByVal Height As Integer, ByVal FieldSpacing As Integer, Optional ByRef NewFieldStyle As FormFieldStyle = Nothing)
        DoubleBuffered = True
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        UpdateStyles()

        If NewFieldStyle IsNot Nothing Then
            varNewFieldStyle = NewFieldStyle
        End If
        RowList = New List(Of FormRow)
        varFieldSpacing = FieldSpacing
        With Me
            .SuspendLayout()
            .Width = Width
            .Height = 300
            .ResumeLayout()
        End With
    End Sub
    'Public Sub HeightToContent()
    '    Height = RowList(RowList.Count - 1).Bottom
    'End Sub
    Protected Overrides Sub OnResize(e As EventArgs)
        SuspendLayout()
        MyBase.OnResize(e)
        If RowList.Count > 0 Then
            Dim Testheight As Integer = RowList(RowList.Count - 1).Bottom - varFieldSpacing
            If Height <> Testheight Then
                Height = Testheight
            End If
        End If
        ResumeLayout(True)
    End Sub
    Public Sub ClearAll(Optional ByVal ClearNonInput As Boolean = False)
        SuspendLayout()
        For Each R As FormRow In RowList
            R.ClearAll(ClearNonInput)
        Next
        ResumeLayout(True)
    End Sub
    Protected Overrides Sub OnVisibleChanged(e As EventArgs)
        MyBase.OnVisibleChanged(e)
        If RowList.Count > 0 Then
            Dim Testheight As Integer = RowList(RowList.Count - 1).Bottom - varFieldSpacing
            If Height <> Testheight Then
                Height = Testheight
            End If
        End If
    End Sub
    Public Sub MergeWithAbove(ByVal Row As Integer, ByVal FieldIndex As Integer, Optional ByVal UpperFieldIndexOffset As Integer = 0, Optional ByVal RepositionLowerChildren As Boolean = False)
        Dim Upper, Lower As FormField
        With RowList
            Upper = .Item(Row - 1).Fields(FieldIndex + UpperFieldIndexOffset)
            Lower = .Item(Row).Fields(FieldIndex)
        End With
        With Lower
            If RepositionLowerChildren Then
                .SwitchHeader(False)
            Else
                .Header.Hide()
            End If
        End With
        With Upper
            .Extrude(FieldExtrudeSide.Bottom, varFieldSpacing)
            Dim TotalHeight As Integer = .Height + Lower.Height
            Dim Ratio As Double = .Height / TotalHeight
            .TopColor = Color.FromArgb(50, ColorHelper.Multiply(.BackColor, 2))
            .BottomColor = Color.FromArgb(50, ColorHelper.Mix(Color.FromArgb(0, .BackColor), .TopColor, Ratio))
        End With
        With Lower
            .TopColor = Upper.BottomColor
            .BottomColor = Color.FromArgb(0, .BackColor)
        End With
    End Sub
    Public Sub Display()
        With RowList
            Dim counter As Integer = 0
            Dim iLast As Integer = .Count - 1
            For i As Integer = 0 To iLast
                For Each F As FormField In .Item(i).Fields
                    If Not F.varFieldType = FormElementType.Label Then
                        F.TabIndex = counter
                        counter += 1
                    End If
                Next
                .Item(i).Display()
            Next
        End With
    End Sub
    Public Sub AddRow(Optional ByVal RowHeight As Integer = 40)
        Dim NewRow As New FormRow(varFieldSpacing)
        With NewRow
            .Hide()
            .Width = Width
            .Height = RowHeight + varFieldSpacing
            .Parent = Me
            .Left = 0
            If RowList.Count > 0 Then
                .Top = RowList(RowList.Count - 1).Bottom
            Else
                .Top = 0
            End If
        End With
        RowList.Add(NewRow)
    End Sub
    Public Sub SetTabIndex()
        Dim i As Integer = 0
        For Each R As FormRow In RowList
            For Each F As FormField In R.Fields
                F.TabIndex = i
                i += 1
            Next
        Next
    End Sub
    Public Sub AddRadioContext(IsRequired As Boolean)
        Dim R As New RadioButtonContext(IsRequired)
        RadioContextList.Add(R)
    End Sub
    Public Overloads Sub AddField(Elements() As FormElement)
        With Elements
            Dim iLast As Integer = .Length - 1
            For i As Integer = 0 To iLast
                With .ElementAt(i)
                    AddField(.Type, .Width, .Value)
                End With
            Next
        End With
    End Sub
    Public Overloads Sub AddField(ByVal FieldType As FormElementType, Optional ByVal FieldWidth As Integer = -1, Optional ByVal Value As Object = Nothing)
        Dim LastRow As FormRow
        With RowList
            If .Count > 0 Then
                LastRow = .Item(.Count - 1)
                If LastRow.CalculateTotalWidth + Math.Abs(FieldWidth) > LastRow.Width Then
                    Dim NewRow As New FormRow(varFieldSpacing)
                    With NewRow
                        .Hide()
                        .Width = Width
                        .Height = varRowHeight + varFieldSpacing
                        .Parent = Me
                        .Left = 0
                        .Top = LastRow.Bottom
                    End With
                    .Add(NewRow)
                End If
            Else
                Dim NewRow As New FormRow(varFieldSpacing)
                With NewRow
                    .Hide()
                    .Width = Width
                    .Height = varRowHeight + varFieldSpacing
                    .Parent = Me
                    .Left = 0
                    .Top = 0
                End With
                .Add(NewRow)
            End If
            'If .Count = 0 Then
            '    Dim NewRow As New FormRow(varFieldSpacing)
            '    With NewRow
            '        .Hide()
            '        .Parent = Me
            '        .Top = 0
            '        .Left = 0
            '        .Width = Width
            '        .Height = varRowHeight + varFieldSpacing
            '        '.BackColor = Color.FromArgb(7, 70, 91)
            '        .BackColor = Color.Green
            '    End With
            '    .Add(NewRow)
            'End If
            LastRow = .Item(.Count - 1)
            Dim NewControl As FormField
            Select Case FieldType
                Case FormElementType.CheckBox
                    NewControl = New FormCheckBox(varFieldSpacing, varNewFieldStyle)
                Case FormElementType.TextField
                    NewControl = New FormTextField(varFieldSpacing, varNewFieldStyle)
                    With NewControl
                        '.BackColor = ColorHelper.Multiply(.BackColor, 0.4)
                        .DrawGradient = False
                    End With
                Case FormElementType.Button
                    NewControl = New FormButton(varFieldSpacing, varNewFieldStyle)
                Case FormElementType.DropDown
                    NewControl = New FormDropDown(varFieldSpacing, varNewFieldStyle)
                Case FormElementType.Label
                    NewControl = New FormLabel(varFieldSpacing, varNewFieldStyle)
                Case FormElementType.Radio
                    With RadioContextList
                        If .Count = 0 Then .Add(New RadioButtonContext(True))
                        With .Item(.Count - 1)
                            NewControl = New FormRadioButton(.Count, (varFieldSpacing), varNewFieldStyle)
                            .Add(DirectCast(NewControl, FormRadioButton))
                        End With
                    End With
                Case FormElementType.Space
                    NewControl = New FormSpace(varFieldSpacing, varNewFieldStyle)
                Case Else
                    NewControl = Nothing
            End Select
            'With NewControl
            '    .Height = LastRow.Height - varFieldSpacing
            'End With
            If Value IsNot Nothing Then
                NewControl.Value = Value
            End If
            With LastRow
                If FieldWidth < 0 Then
                    .AddField(NewControl, .CalculateRemainingWidth)
                Else
                    .AddField(NewControl, FieldWidth)
                End If
            End With
        End With
    End Sub
    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            Try
                If RadioContextList IsNot Nothing Then
                    With RadioContextList
                        Dim nLast As Integer = .Count - 1
                        For n As Integer = 0 To nLast
                            .Item(n).Dispose()
                        Next
                    End With
                End If
                RadioContextList = Nothing
                If RowList IsNot Nothing Then
                    With RowList
                        Dim iLast As Integer = .Count - 1
                        For i As Integer = 0 To iLast
                            .Item(i).Dispose()
                        Next
                    End With
                End If
            Catch
            End Try
        End If
        MyBase.Dispose(disposing)
    End Sub
#End Region
    Public Class FormRow
        Inherits Control
        Private varFieldSpacing As Integer
        Private FieldList As List(Of FormField)

        'To reduce flicker
        Protected Overrides ReadOnly Property CreateParams() As CreateParams
            Get
                Dim params As CreateParams = MyBase.CreateParams
                params.ExStyle = params.ExStyle Or &H2000000
                Return params
            End Get
        End Property

        Public ReadOnly Property Fields As List(Of FormField)
            Get
                Return FieldList
            End Get
        End Property
        Protected Friend Sub ClearAll(ByVal ClearNonInput As Boolean)
            For Each Field As FormField In FieldList
                Field.Clear(ClearNonInput)
            Next
        End Sub
        Protected Friend ReadOnly Property FieldSpacing As Integer
            Get
                Return varFieldSpacing
            End Get
        End Property
        Protected Friend Function Validate() As Boolean
            Dim EverythingValid As Boolean = True
            For Each F As FormField In FieldList
                If Not F.Validate Then
                    EverythingValid = False
                End If
            Next
            Return EverythingValid
        End Function
        Public Sub RemoveGaps()
            SuspendLayout()
            With FieldList
                Dim Part As Integer = Width \ .Count
                For i As Integer = 0 To .Count - 1
                    With .Item(i)
                        .Width = Part
                        .Left = Part * i
                    End With
                Next
                With .Item(.Count - 1)
                    .Width = Width - .Left
                End With
            End With
            ResumeLayout()
        End Sub
        Protected Friend Sub Display()
            With Controls
                Dim iLast As Integer = .Count - 1
                For i As Integer = 0 To iLast
                    .Item(i).Show()
                Next
            End With
            Show()
        End Sub
        Protected Overrides Sub OnResize(e As EventArgs)
            MyBase.OnResize(e)
            With FieldList
                Dim iLast As Integer = .Count - 1
                If iLast >= 0 Then
                    For i As Integer = 0 To iLast
                        .Item(i).Height = Height - varFieldSpacing
                    Next
                End If
            End With
        End Sub
        Protected Friend Sub New(ByVal FieldSpacing As Integer)
            DoubleBuffered = True
            SetStyle(ControlStyles.UserPaint, True)
            SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
            SetStyle(ControlStyles.AllPaintingInWmPaint, True)
            SetStyle(ControlStyles.Selectable, False)
            UpdateStyles()

            FieldList = New List(Of FormField)
            TabStop = False
            'GradSurface = New Control
            'With GradSurface
            '.Parent = Me
            '.Top = 0
            '.Left = 0
            '.SendToBack()
            'End With
            varFieldSpacing = FieldSpacing
            'GradBrush = New LinearGradientBrush(New Point(0, 0), New Point(Left, Bottom), ColorFirst, ColorSecond)
        End Sub
        Protected Friend Sub AddField(Item As FormField, ByVal FieldWidth As Integer)
            With FieldList
                If .Count > 0 Then
                    Item.Left = .Item(.Count - 1).Right + varFieldSpacing
                    'Else
                    '    Item.Left = 0
                End If
                With Item
                    '.Top = 0
                    .Height = Height - varFieldSpacing
                    .Width = FieldWidth
                End With
                Controls.Add(Item)
                .Add(Item)
            End With
        End Sub
        Protected Friend Function CalculateTotalWidth() As Integer
            Dim OtherTotal As Integer
            With FieldList
                If .Count > 0 Then
                    OtherTotal = .Item(.Count - 1).Right + varFieldSpacing
                End If
            End With
            Return OtherTotal
        End Function
        Protected Friend Function CalculateRemainingWidth() As Integer
            ' Calculate total width
            Dim OtherTotal As Integer
            With FieldList
                If .Count > 0 Then
                    OtherTotal = .Item(.Count - 1).Right + varFieldSpacing
                End If
            End With
            ' Do not count the left and right side
            'OtherTotal -= varFieldSpacing * 2
            ' Return difference
            ' TODO: No loop
            Return Width - OtherTotal
        End Function
        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                With FieldList
                    Dim iLast As Integer = .Count - 1
                    For i As Integer = 0 To iLast
                        .Item(i).Dispose()
                    Next
                End With
            End If
            MyBase.Dispose(disposing)
        End Sub
    End Class
    Private Class FormSpace
        Inherits FormField
        Public Sub New(ByVal FieldSpacing As Integer, Style As FormFieldStyle)
            MyBase.New(FieldSpacing, FormElementType.CheckBox, Style)
        End Sub
    End Class
    Private Class FormCheckBox
        Inherits FormField
        Private WithEvents TextLab As Label
        Private WithEvents Check As PictureBox
        Private FocusedBrush As New SolidBrush(Color.Gray)
        Private varChecked As Boolean
        Private Shared varCheckBoxBorderPen As New Pen(Color.FromArgb(240, 240, 240))
        Private CheckBorderRect As Rectangle
        Private CheckBrush As New SolidBrush(Color.Black)
        Public Overrides Function Validate() As Boolean
            If varRequired And Not DirectCast(Value, Boolean) Then
                IsValid = False
            ElseIf varRequireSpecificValue AndAlso Not (varRequiredValue.Equals(Value)) Then
                IsValid = False
            End If
            Return varIsValid
        End Function
        Private Property Checked As Boolean
            Get
                Return varChecked
            End Get
            Set(value As Boolean)
                varChecked = value
                Me.Value = value
                If Check IsNot Nothing Then
                    Check.Invalidate()
                End If
            End Set
        End Property
        Protected Friend Overrides Sub Clear(ByVal ClearNonInput As Boolean)
            MyBase.Clear(ClearNonInput)
            Value = False
        End Sub
        Protected Overrides Sub OnUserClick()
            MyBase.OnUserClick()
            If varChecked Then
                Checked = False
            Else
                Checked = True
            End If
        End Sub
        Private Sub OnHeaderVisibleChanged() Handles Me.HeaderVisibleChanged
            With TextLab
                .SuspendLayout()
                .Top = Header.Height
                .Height = Height - .Top
                .Width = Width
                .ResumeLayout()
            End With
            With Check
                .Top = CInt(TextLab.Top + (TextLab.Height / 2) - .Height / 2)
            End With
        End Sub
        Private Sub OnCheckPaint(Sender As Object, e As PaintEventArgs) Handles Check.Paint
            With e.Graphics
                .DrawRectangle(varCheckBoxBorderPen, CheckBorderRect)
                If varChecked Then
                    .SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                    .DrawString(ChrW(&H2713), SystemFonts.MessageBoxFont, CheckBrush, Point.Empty)
                End If
            End With
        End Sub
        Protected Overrides Sub OnValueChanged(Value As Object)
            MyBase.OnValueChanged(Value)
            Try
                varChecked = DirectCast(Value, Boolean)
            Catch
                Throw New InvalidCastException("This FormField instance expects the value object to be a boolean value (checked or not checked).")
            End Try
            If Check IsNot Nothing Then
                Check.Invalidate()
            End If
        End Sub
        Protected Overrides Sub OnSecondaryValueChanged(Value As Object)
            MyBase.OnSecondaryValueChanged(Value)
            TextLab.Text = DirectCast(Value, String)
        End Sub
        Private Sub OnCheckMouseUp(Sender As Object, e As EventArgs) Handles Check.MouseUp
            If varChecked Then
                varChecked = False
            Else
                varChecked = True
            End If
            Value = varChecked
            Check.Invalidate()
            OnValueChanged(varChecked)
        End Sub
        Public Sub New(ByVal FieldSpacing As Integer, Style As FormFieldStyle)
            MyBase.New(FieldSpacing, FormElementType.CheckBox, Style)
            SetStyle(ControlStyles.Selectable, True)
            MeType = FormElementType.CheckBox
            Check = New PictureBox
            TextLab = New Label
            Hide()
            With Check
                .Hide()
                .Parent = Me
                .Width = 16
                .Height = 16
                .BackColor = Color.White
                .Left = 10
                .Show()
            End With
            CheckBorderRect = New Rectangle(Point.Empty, New Size(15, 15))
            With TextLab
                .Hide()
                .Parent = Me
                '.BackColor = Color.FromArgb(7, 70, 91)
                .BackColor = Color.Transparent
                .Padding = New Padding(34, 0, 0, 0)
                .Left = 0
                .Top = 14
                .ForeColor = Style.BodyFG
                .Text = "Meld meg på"
                .FlatStyle = FlatStyle.Flat
                .TextAlign = ContentAlignment.MiddleLeft
                .Show()
            End With
            Value = False
            Show()
        End Sub
        Protected Overrides Sub OnSizeChanged(e As EventArgs)
            MyBase.OnSizeChanged(e)
            With TextLab
                .SuspendLayout()
                .Top = Header.Height
                .Height = Height - .Top
                .Width = Width
                .ResumeLayout()
            End With
            With Check
                .Top = CInt(TextLab.Top + (TextLab.Height / 2) - .Height / 2)
            End With
        End Sub
        Protected Overrides Sub OnVisibleChanged(e As EventArgs)
            MyBase.OnVisibleChanged(e)
            With TextLab
                .SuspendLayout()
                .Top = Header.Height
                .Height = Height - .Top
                .Width = Width
                .ResumeLayout()
            End With
            With Check
                .Top = CInt(TextLab.Top + (TextLab.Height / 2) - .Height / 2)
            End With
        End Sub
        Protected Overrides Sub Dispose(disposing As Boolean)
            TextLab.Dispose()
            Check.Dispose()
            varCheckBoxBorderPen.Dispose()
            CheckBrush.Dispose()
            FocusedBrush.Dispose()
            MyBase.Dispose(disposing)
        End Sub
    End Class

    Private Class FormButton
        Inherits FormField
        Private TextLab As New Label
        Public Sub New(ByVal FieldSpacing As Integer, Style As FormFieldStyle)
            MyBase.New(FieldSpacing, FormElementType.Button, Style)
            MeType = FormElementType.Button
        End Sub
        Protected Overrides Sub OnSecondaryValueChanged(Value As Object)
            MyBase.OnSecondaryValueChanged(Value)
            TextLab.Text = DirectCast(Value, String)
        End Sub
        ' TODO: Handle OnHeaderVisibleChanged
        Protected Overrides Sub Dispose(disposing As Boolean)
            TextLab.Dispose()
            MyBase.Dispose(disposing)
        End Sub
    End Class

    Public Class FormTextField
        Inherits FormField
        Private PaddingLeft As Integer = 10
        Private WithEvents varTextField As TextBox
        Private varPlaceHolder As String = ""
        Private varDrawPlaceHolder As Boolean
        Private PlaceHolderPoint As Point = Point.Empty
        Private WithEvents PlaceHolderSurface As New Control
        Private Sub SendPlaceHolderSurfaceToBack() Handles PlaceHolderSurface.Click, PlaceHolderSurface.GotFocus
            TextField.Focus()
            TextField.BringToFront()
            PlaceHolderSurface.Hide()
            If Not varIsValid Then
                IsValid = True
            End If
        End Sub
        Protected Overrides Sub OnMaxLengthChanged()
            MyBase.OnMaxLengthChanged()
            varTextField.MaxLength = varMinMax(1)
        End Sub
        Public Property PlaceHolder As String
            Get
                Return varPlaceHolder
            End Get
            Set(value As String)
                varPlaceHolder = value
            End Set
        End Property
        Protected Friend Overrides Sub Clear(ByVal ClearNonInput As Boolean)
            MyBase.Clear(ClearNonInput)
            varTextField.Clear()
        End Sub
        Public Overrides Function Validate() As Boolean
            IsValid = True
            If varRequired And (TextField.Text = "" OrElse TextField.Text Is Nothing) Then
                IsValid = False
            ElseIf varRequireSpecificValue AndAlso Not varRequiredValue.Equals(Value) Then
                IsValid = False
            ElseIf varIsNumeric AndAlso Not IsNumeric(TextField.Text) Then
                IsValid = False
            Else
                Dim TFTL As Integer = TextField.TextLength
                If varMinMax(0) > TFTL OrElse varMinMax(1) < TFTL Then
                    IsValid = False
                End If
            End If
            Return IsValid
        End Function
        Private Sub PlaceHolder_Paint(Sender As Object, e As PaintEventArgs) Handles PlaceHolderSurface.Paint
            If varPlaceHolder <> "" Then
                Using NewBrush As New SolidBrush(Color.FromArgb(70, ForeColor))
                    With PlaceHolderSurface
                        e.Graphics.DrawString(varPlaceHolder, varTextField.Font, NewBrush, PlaceHolderPoint)
                    End With
                End Using
            End If
        End Sub
        Private Sub Field_KeyPress(Sender As Object, e As KeyPressEventArgs) Handles varTextField.KeyPress
            If varIsNumeric AndAlso Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
                e.Handled = True
            End If
        End Sub
        Private Sub Field_TextChanged() Handles varTextField.TextChanged
            Value = varTextField.Text
            If TextField.Text = "" AndAlso Not TextField.Focused Then
                TextField.SendToBack()
                PlaceHolderSurface.Show()
            Else
                TextField.BringToFront()
                PlaceHolderSurface.Hide()
            End If
            With PlaceHolderSurface
                PlaceHolderPoint = New Point(PaddingLeft - 6, .Height \ 2 - TextRenderer.MeasureText(varTextField.Text, varTextField.Font).Height \ 2 - 6)
            End With
        End Sub
        Private Sub Field_LostFocused() Handles varTextField.LostFocus
            If TextField.Text = "" Then
                TextField.SendToBack()
                PlaceHolderSurface.Show()
            End If
        End Sub
        Public ReadOnly Property TextField As TextBox
            Get
                Return varTextField
            End Get
        End Property
        Public Shadows Property Font As Font
            Get
                Return MyBase.Font
            End Get
            Set(value As Font)
                MyBase.Font = value
            End Set
        End Property
        Private Sub Me_BackColorChanged(Sender As Object, e As EventArgs) Handles Me.BackColorChanged
            If varTextField IsNot Nothing Then
                varTextField.BackColor = BackColor
            End If
        End Sub

        Public Sub New(ByVal FieldSpacing As Integer, Style As FormFieldStyle)
            MyBase.New(FieldSpacing, FormElementType.TextField, Style)
            MeType = FormElementType.TextField
            varTextField = New TextBox
            Hide()
            With PlaceHolderSurface
                .Parent = Me
                .Width = Width - 2
                .Left = 1
                .Top = Header.Bottom + 1
                .Height = Height - .Top - 2
                .Cursor = Cursors.IBeam
                ' Sett font
            End With
            With varTextField
                .Parent = Me
                .Left = PaddingLeft
                .Width = Width - PaddingLeft * 2
                .Height = 16
                .Top = CInt((Height / 2) - (.Height / 2) + (Header.Height / 2))
                .ForeColor = Style.TextFG
                .BorderStyle = BorderStyle.None
                .SendToBack()
            End With
            BackColor = Style.TextBG
            Value = ""
            Show()
        End Sub
        Protected Overrides Sub OnSizeChanged(e As EventArgs)
            MyBase.OnSizeChanged(e)
            With varTextField
                .Width = Width - PaddingLeft * 2
                .Top = CInt((Height / 2) - (.Height / 2) + (Header.Height / 2))
            End With
            With PlaceHolderSurface
                .Width = Width - 2
                .Top = Header.Bottom + 1
                .Height = Height - .Top - 2
                PlaceHolderPoint = New Point(PaddingLeft - 6, .Height \ 2 - TextRenderer.MeasureText(varTextField.Text, varTextField.Font).Height \ 2 - 6)
            End With
        End Sub
        Protected Overrides Sub OnVisibleChanged(e As EventArgs)
            MyBase.OnVisibleChanged(e)
            With varTextField
                .Width = Width - PaddingLeft * 2
                .Top = CInt((Height / 2) - (.Height / 2) + (Header.Height / 2))
            End With
            With PlaceHolderSurface
                .Width = Width - 2
                .Top = Header.Bottom + 1
                .Height = Height - .Top - 2
                PlaceHolderPoint = New Point(PaddingLeft - 6, .Height \ 2 - TextRenderer.MeasureText(varTextField.Text, varTextField.Font).Height \ 2 - 6)
            End With
        End Sub
        Protected Overrides Sub OnValueChanged(Value As Object)
            MyBase.OnValueChanged(Value)
            varTextField.Text = DirectCast(Value, String)
        End Sub
        Protected Overrides Sub OnSecondaryValueChanged(Value As Object)
            MyBase.OnSecondaryValueChanged(Value)
            varTextField.Text = DirectCast(Value, String)
            If Not varIsValid Then
                IsValid = True
            End If
        End Sub
    End Class

    Private Class FormLabel
        Inherits FormField
        Private WithEvents TextLab As Label
        Private Sub OnTextLabAlignChanged() Handles TextLab.TextAlignChanged
            Select Case TextLab.TextAlign
                Case ContentAlignment.TopLeft
                    TextLab.Padding = New Padding(6, 3, 0, 0)
                Case ContentAlignment.TopCenter
                    TextLab.Padding = New Padding(0, 3, 0, 0)
                Case ContentAlignment.TopRight
                    TextLab.Padding = New Padding(0, 3, 6, 0)
                Case ContentAlignment.MiddleLeft
                    TextLab.Padding = New Padding(6, 0, 0, 0)
                Case ContentAlignment.MiddleCenter
                    TextLab.Padding = New Padding(0, 0, 0, 0)
                Case ContentAlignment.MiddleRight
                    TextLab.Padding = New Padding(0, 0, 6, 0)
                Case ContentAlignment.BottomLeft
                    TextLab.Padding = New Padding(6, 0, 0, 3)
                Case ContentAlignment.BottomCenter
                    TextLab.Padding = New Padding(0, 0, 0, 3)
                Case ContentAlignment.BottomRight
                    TextLab.Padding = New Padding(0, 0, 6, 3)
            End Select
        End Sub
        Public Sub New(ByVal FieldSpacing As Integer, Style As FormFieldStyle)
            MyBase.New(FieldSpacing, FormElementType.Label, Style)
            MeType = FormElementType.Label
            varIsInputType = False
            SetStyle(ControlStyles.Selectable, False)
            TextLab = New Label
            Hide()
            With TextLab
                .TabStop = False
                .Parent = Me
                '.BackColor = Color.FromArgb(7, 70, 91)
                .BackColor = Color.Transparent
                .ForeColor = Style.BodyFG
                .AutoSize = False
                .TextAlign = ContentAlignment.MiddleLeft
                .Padding = New Padding(6, 0, 6, 0)
                .Left = 1
                .Top = Header.Height
                .Width = Width - 2
            End With
            Value = "Change the default text using the Value property."
        End Sub
        Protected Friend Overrides Sub Clear(ClearNonInput As Boolean)
            MyBase.Clear(ClearNonInput)
            If ClearNonInput Then
                Value = ""
            End If
        End Sub
        Protected Overrides Sub OnValueChanged(Value As Object)
            MyBase.OnValueChanged(Value)
            TextLab.Text = DirectCast(Value, String)
        End Sub
        Protected Overrides Sub OnSecondaryValueChanged(Value As Object)
            MyBase.OnSecondaryValueChanged(Value)
            TextLab.Text = DirectCast(Value, String)
        End Sub
        Protected Overrides Sub OnSizeChanged(e As EventArgs)
            MyBase.OnSizeChanged(e)
            With TextLab
                .SuspendLayout()
                .Width = Width
                .Height = Height - Header.Height
                .Top = Header.Height
                .ResumeLayout()
                .PerformLayout()
            End With
        End Sub
        Protected Overrides Sub OnVisibleChanged(e As EventArgs)
            MyBase.OnVisibleChanged(e)
            With TextLab
                .SuspendLayout()
                .Width = Width
                .Height = Height - Header.Height
                .Top = Header.Height
                .ResumeLayout()
            End With
        End Sub
        Private Sub OnHeaderVisibleChanged() Handles Me.HeaderVisibleChanged
            With TextLab
                .SuspendLayout()
                .Width = Width
                .Height = Height - Header.Height
                .Top = Header.Height
                .ResumeLayout()
            End With
        End Sub
        Protected Overrides Sub Dispose(disposing As Boolean)
            TextLab.Dispose()
            MyBase.Dispose(disposing)
        End Sub
    End Class

    Private Class FormDropDown
        Inherits FormField
        Public Sub New(ByVal FieldSpacing As Integer, Style As FormFieldStyle)
            MyBase.New(FieldSpacing, FormElementType.DropDown, Style)
            MeType = FormElementType.DropDown
        End Sub
        Protected Overrides Sub Dispose(disposing As Boolean)
            MyBase.Dispose(disposing)
        End Sub
    End Class

    Public Class FormRadioButton
        Inherits FormField
        Private WithEvents CheckSurface As PictureBox
        Private WithEvents TextLab As Label
        Private varChecked As Boolean
        Private CheckBrush As New SolidBrush(Color.White)
        Private DotBrush As New SolidBrush(Color.Black)
        Private HoverBrush As New SolidBrush(Color.LightGray)
        Private RadioBorderPen As New Pen(Color.FromArgb(120, 120, 120), 1.5)
        Private Hovering As Boolean
        Private varIndex As Integer
        Public ReadOnly Property Checked As Boolean
            Get
                Return varChecked
            End Get
        End Property
        Private Sub Check()
            varChecked = True
            CheckSurface.Refresh()
            Value = varChecked
        End Sub
        Public ReadOnly Property RadioIndex As Integer
            Get
                Return varIndex
            End Get
        End Property
        Protected Overrides Sub OnUserClick()
            MyBase.OnUserClick()
            If Not varChecked Then
                Value = True
            End If
        End Sub
        Public Overrides Function Validate() As Boolean
            If varRequireSpecificValue AndAlso Not (varRequiredValue.Equals(Value)) Then
                IsValid = False
            Else
                IsValid = True
            End If
            Return IsValid
        End Function
        Protected Friend Overrides Sub Clear(ByVal ClearNonInput As Boolean)
            MyBase.Clear(ClearNonInput)
            Value = False
        End Sub
        Public Sub New(ByVal RadioContextIndex As Integer, ByVal FieldSpacing As Integer, Style As FormFieldStyle)
            MyBase.New(FieldSpacing, FormElementType.Radio, Style)
            SuspendLayout()
            MeType = FormElementType.Radio
            CheckSurface = New PictureBox
            TextLab = New Label
            Hide()
            varIndex = RadioContextIndex
            With CheckSurface
                .Parent = Me
                .Size = New Size(16, 16)
                '.BackColor = Color.FromArgb(7, 70, 91)
                .BackColor = Color.Transparent
                .Left = 10
            End With
            With TextLab
                .Hide()
                .AutoSize = False
                .Parent = Me
                '.BackColor = Color.FromArgb(7, 70, 91)
                .BackColor = Color.Transparent
                .Padding = New Padding(34, 0, 0, 0)
                .Location = New Point(0, Header.Height)
                .Height = Height - Header.Height
                .ForeColor = Style.BodyFG
                .Text = "Meld meg på"
                .FlatStyle = FlatStyle.Flat
                .TextAlign = ContentAlignment.MiddleLeft
                .SendToBack()
                .Show()
            End With
            Value = False
            Show()
            ResumeLayout()
        End Sub
        Private Sub CheckSurface_Paint(Sender As Object, e As PaintEventArgs) Handles CheckSurface.Paint
            Dim DotRect, BorderRect As Rectangle
            With CheckSurface
                DotRect = New Rectangle(0, 0, .Width - 1, .Height - 1)
                BorderRect = New Rectangle(0, 0, .Width - 1, .Height - 1)
            End With
            With e.Graphics
                .SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                .FillEllipse(CheckBrush, DotRect)
                .DrawEllipse(RadioBorderPen, BorderRect)
                If varChecked Then
                    DotRect.Inflate(-4, -4)
                    .FillEllipse(DotBrush, DotRect)
                ElseIf Hovering Then
                    DotRect.Inflate(-4, -4)
                    .FillEllipse(HoverBrush, DotRect)
                End If
            End With
        End Sub
        Private Sub CheckSurface_MouseEnter(Sender As Object, e As EventArgs) Handles CheckSurface.MouseEnter
            Hovering = True
            CheckSurface.Refresh()
        End Sub
        Private Sub CheckSurface_MouseLeave(Sender As Object, e As EventArgs) Handles CheckSurface.MouseLeave
            Hovering = False
            CheckSurface.Refresh()
        End Sub
        Private Sub CheckSurface_Click(Sender As Object, e As EventArgs) Handles CheckSurface.Click
            If Hovering Then
                varChecked = True
            End If
            CheckSurface.Refresh()
            OnValueChanged(varChecked)
        End Sub
        Protected Overrides Sub OnSizeChanged(e As EventArgs)
            MyBase.OnSizeChanged(e)
            SuspendLayout()
            With TextLab
                .Width = Width
                .Height = Height - Header.Height
                .Top = Header.Height
            End With
            With Header
                CheckSurface.Top = .Height + (Height - .Height) \ 2 - 8
                CheckSurface.Left = 10
            End With
            ResumeLayout()
        End Sub
        Protected Overrides Sub OnVisibleChanged(e As EventArgs)
            MyBase.OnVisibleChanged(e)
            SuspendLayout()
            With TextLab
                .Width = Width
                .Height = Height - Header.Height
                .Top = Header.Height
            End With
            'With CheckSurface
            '    .Top = CInt((Height / 2) - (.Top / 2)) + TextLab.Top - 2
            'End With
            With Header
                CheckSurface.Top = .Height + (Height - .Height) \ 2 - 8
                CheckSurface.Left = 10
            End With
            ResumeLayout()
        End Sub
        Protected Overrides Sub OnValueChanged(Value As Object)
            MyBase.OnValueChanged(Value)
            varChecked = DirectCast(Value, Boolean)
            Value = varChecked
            CheckSurface.Refresh()
        End Sub
        Protected Overrides Sub OnSecondaryValueChanged(Value As Object)
            MyBase.OnSecondaryValueChanged(Value)
            TextLab.Text = DirectCast(Value, String)
        End Sub
        Protected Overrides Sub Dispose(disposing As Boolean)
            TextLab.Dispose()
            CheckSurface.Dispose()
            CheckBrush.Dispose()
            DotBrush.Dispose()
            RadioBorderPen.Dispose()
            HoverBrush.Dispose()
            MyBase.Dispose(disposing)
        End Sub
    End Class

    Public Class RadioButtonContext
        Implements IDisposable
        Private varIsRequired As Boolean
        Private IsValid As Boolean = True
        Private ControlList As List(Of FormRadioButton)
        Public Function GetChecked() As Integer
            Dim Ret As Integer = -1
            Dim Counter As Integer = 0
            For Each C As FormRadioButton In ControlList
                If C.Checked Then
                    Ret = Counter
                    Exit For
                Else
                    Counter += 1
                End If
            Next
            Return Ret
        End Function
        Public Property IsRequired As Boolean
            Get
                Return varIsRequired
            End Get
            Set(value As Boolean)
                varIsRequired = value
            End Set
        End Property
        Public ReadOnly Property Count As Integer
            Get
                Return ControlList.Count
            End Get
        End Property
        Public Function Validate() As Boolean
            IsValid = True
            If varIsRequired Then
                Dim SomethingSelected As Boolean
                For Each R As FormRadioButton In ControlList
                    If DirectCast(R.Value, Boolean) Then
                        SomethingSelected = True
                        Exit For
                    End If
                Next
                If Not SomethingSelected Then
                    IsValid = False
                    For Each R As FormRadioButton In ControlList
                        R.IsValid = False
                    Next
                End If
            End If
            Return IsValid
        End Function
        Protected Friend Sub New(IsRequired As Boolean)
            varIsRequired = IsRequired
            ControlList = New List(Of FormRadioButton)(4)
        End Sub
        Public Sub Add(R As FormRadioButton)
            ControlList.Add(R)
            AddHandler R.ValueChanged, AddressOf OnValueChanged
        End Sub
        Private Sub OnValueChanged(Sender As FormField, Value As Object)
            Dim SenderX As FormRadioButton = DirectCast(Sender, FormRadioButton)
            If DirectCast(Value, Boolean) Then
                With ControlList
                    Dim iLast As Integer = .Count - 1
                    For i As Integer = 0 To iLast
                        With .Item(i)
                            If .RadioIndex <> SenderX.RadioIndex Then
                                .Value = False
                            End If
                        End With
                    Next
                End With
            End If
        End Sub

#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                End If
                With ControlList
                    Dim iLast As Integer = .Count - 1
                    For i As Integer = 0 To iLast
                        RemoveHandler .Item(i).ValueChanged, AddressOf OnValueChanged
                    Next
                End With
                ControlList.Clear()
                ControlList.TrimExcess()
                ControlList = Nothing
                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(True)
            ' TODO: uncomment the following line if Finalize() is overridden above.
            ' GC.SuppressFinalize(Me)
        End Sub
#End Region
    End Class
End Class
