Option Strict On
Option Explicit On
Option Infer Off
Imports System.Drawing.Text
Imports System.Threading
Imports AudiopoLib

Public Class BlodBeholder
    Inherits Control
    Private WithEvents Empty As New PictureBox
    Private EmptyBrush As New SolidBrush(Color.FromArgb(15, 79, 117))
    Private FullBrush As New SolidBrush(Color.White)
    Private LargeFont As New Font(Font.FontFamily, 110)
    Private SmallFont As New Font(Font.FontFamily, 12)
    Private LargeText As String = "0%"
    Private TextSizes(2) As Size
    Private LargeTextSize As Size = Size.Empty
    Private SmallText() As String = {"Du har hittil donert ca. 0%", "av blodvolumet i et", "gjennomsnittlig menneske."}
    Private SC As SynchronizationContext = SynchronizationContext.Current
    Private WithEvents SlideTimer As New Timers.Timer(1000 \ 50)
    Private Initial, Current, Goal As Double
    Private Const MengdeBlodIMenneskeILiter As Double = 4.5
    Private Const XLast As Integer = 40
    Private CurrentX As Integer = 0
    Private WithEvents DBC_GetBloodAmount As New DatabaseClient(Credentials.Server, Credentials.Database, Credentials.UserID, Credentials.Password)
    Public Sub GetBlood()
        DBC_GetBloodAmount.SQLQuery = "SELECT A.ant_liter FROM Blodtapping A INNER JOIN Time B ON A.time_id = B.time_id WHERE B.b_fodselsnr = @nr;"
        DBC_GetBloodAmount.Execute({"@nr"}, {CurrentLogin.PersonalNumber})
    End Sub
    Private Sub DBC_Finished(Sender As Object, e As DatabaseListEventArgs) Handles DBC_GetBloodAmount.ListLoaded
        If e.ErrorOccurred Then

        Else
            Dim Sum As Double = 0
            For Each R As DataRow In e.Data.Rows
                Sum += DirectCast(R.Item(0), Double)
            Next
            TotalBloodInLiters = Sum
        End If
    End Sub
    Public Sub New(EmptyBitmap As Bitmap, FullBitmap As Bitmap)
        DoubleBuffered = True
        With Empty
            .Parent = Me
            .BackgroundImage = EmptyBitmap
            .Size = .BackgroundImage.Size
            .Top = 0
        End With
        With SlideTimer
            .AutoReset = False
        End With
        BackgroundImage = FullBitmap
        Size = BackgroundImage.Size
        LargeTextSize = TextRenderer.MeasureText(LargeText, LargeFont)
        TextSizes = {TextRenderer.MeasureText(SmallText(0), SmallFont), TextRenderer.MeasureText(SmallText(1), SmallFont), TextRenderer.MeasureText(SmallText(2), SmallFont)}
    End Sub
    Private Sub Empty_MouseEnter(sender As Object, e As EventArgs) Handles Empty.MouseEnter
        OnMouseEnter(e)
    End Sub
    Private Sub Empty_MouseLeave(sender As Object, e As EventArgs) Handles Empty.MouseLeave
        OnMouseLeave(e)
    End Sub
    Private Sub Empty_MouseUp(sender As Object, e As EventArgs) Handles Empty.MouseUp
        OnClick(e)
    End Sub
    Private Sub Empty_DoubleClick(sender As Object, e As EventArgs) Handles Empty.DoubleClick
        OnDoubleClick(e)
    End Sub
    Public WriteOnly Property TotalBloodInLiters As Double
        Set(value As Double)
            Dim NewPercentage As Double = Math.Floor(value * 100 / MengdeBlodIMenneskeILiter)
            LargeText = NewPercentage & "%"
            LargeTextSize = TextRenderer.MeasureText(LargeText, LargeFont)
            SmallText = {"Du har hittil donert ca. " & NewPercentage & "%", "av blodvolumet i et", "gjennomsnittlig menneske."}
            TextSizes = {TextRenderer.MeasureText(SmallText(0), SmallFont), TextRenderer.MeasureText(SmallText(1), SmallFont), TextRenderer.MeasureText(SmallText(2), SmallFont)}
            SlideToPercentage(NewPercentage)
        End Set
    End Property
    Private Sub EmptyPaint(sender As Object, e As PaintEventArgs) Handles Empty.Paint
        With e.Graphics
            .TextRenderingHint = TextRenderingHint.AntiAlias
            ' TODO: Make more efficient
            .DrawString(LargeText, LargeFont, EmptyBrush, New Point((Width - LargeTextSize.Width) \ 2, (Height - LargeTextSize.Height) \ 2 - 60))
            .DrawString(SmallText(0), SmallFont, EmptyBrush, New Point((Width - TextSizes(0).Width) \ 2, (Height - TextSizes(0).Height) \ 2 + 50))
            .DrawString(SmallText(1), SmallFont, EmptyBrush, New Point((Width - TextSizes(1).Width) \ 2, (Height - TextSizes(1).Height) \ 2 + 70))
            .DrawString(SmallText(2), SmallFont, EmptyBrush, New Point((Width - TextSizes(2).Width) \ 2, (Height - TextSizes(2).Height) \ 2 + 90))
        End With
    End Sub
    Protected Overrides Sub OnPaintBackground(pevent As PaintEventArgs)
        MyBase.OnPaintBackground(pevent)
        With pevent.Graphics
            .TextRenderingHint = TextRenderingHint.AntiAlias
            ' TODO: Make more efficient
            .DrawString(LargeText, LargeFont, FullBrush, New Point((Width - LargeTextSize.Width) \ 2, (Height - LargeTextSize.Height) \ 2 - 60))
            .DrawString(SmallText(0), SmallFont, FullBrush, New Point((Width - TextSizes(0).Width) \ 2, (Height - TextSizes(0).Height) \ 2 + 50))
            .DrawString(SmallText(1), SmallFont, FullBrush, New Point((Width - TextSizes(1).Width) \ 2, (Height - TextSizes(1).Height) \ 2 + 70))
            .DrawString(SmallText(2), SmallFont, FullBrush, New Point((Width - TextSizes(2).Width) \ 2, (Height - TextSizes(2).Height) \ 2 + 90))
        End With
    End Sub
    Private Sub SlideTimer_Tick(Sender As Object, e As EventArgs) Handles SlideTimer.Elapsed
        SC.Post(AddressOf Send_Tick, Nothing)
    End Sub
    Protected Overrides Sub OnLocationChanged(e As EventArgs)
        SuspendLayout()
        MyBase.OnLocationChanged(e)
        ResumeLayout(True)
    End Sub
    Private Sub Send_Tick(State As Object)
        ' TEST
        Current = EaseInOut.GetY(Initial, Goal, CurrentX, XLast)
        CurrentX += 1
        Empty.Height = CInt(Height - (Height / 100) * Current)
        If CurrentX < XLast Then
            SlideTimer.Start()
        End If
    End Sub
    Public Sub SlideToPercentage(Percent As Double)
        Initial = Current
        CurrentX = 0
        Goal = Percent
        SlideTimer.Start()
    End Sub
End Class