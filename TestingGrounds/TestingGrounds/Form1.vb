Option Strict On
Option Explicit On
Option Infer Off

Public Class Form1
    Dim WithEvents Cal As CustomCalendar
    Dim TestForm1, TestForm2 As FlatForm
    Dim TestQuestionnaire As Questionnaire
    Dim Uni As New Univers(Me, 50, 1000, 700)
    Private Sub Setup_Tjener_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeForms()
        cmdAdd.BringToFront()
        Uni.BringToFront()
        Uni.Dock = DockStyle.Fill
        cmdAdd.BringToFront()
        Button2.BringToFront()
    End Sub

    Private Sub InitializeForms()
        Dim TestForm1 As New FlatForm(400, 300, 10)
        With TestForm1
            .SuspendLayout()
            .AddField(FormElementType.CheckBox, 120)
            .AddField(FormElementType.Label)
            .AddField(FormElementType.CheckBox, 120)
            .AddField(FormElementType.CheckBox, 160)
            .AddField(FormElementType.CheckBox)
            .AddField(FormElementType.TextField)
            .AddField(FormElementType.Radio, 120, True)
            .AddField(FormElementType.Radio)
            .AddRadioContext()
            .AddField(FormElementType.Radio, 120, True)
            .AddField(FormElementType.Radio)
            .Field(1, 2).Extrude(FieldExtrudeSide.Left, 10)
            .Field(0, 1).Value = "Dette er et eksempel på både vertikal og horisontal sammensmelting." & vbNewLine & " "
            .Field(1, 2).SecondaryValue = "Test"
            .MergeWithAbove(1, 1)
            .MergeWithAbove(1, 2, -1)
            With .Field(2, 0)
                .Value = "Hei"
                .Header.Text = "Har du noen ekstra kommentarer?"
            End With
            .ResumeLayout()
            '.HeightToContent()
        End With
        Dim TestForm2 As New FlatForm(400, 300, 10)
        With TestForm2
            .SuspendLayout()
            .AddField(FormElementType.Label, 260)
            .AddField(FormElementType.CheckBox)
            .AddField(FormElementType.CheckBox, 100)
            .AddField(FormElementType.CheckBox, 150)
            .AddField(FormElementType.Radio)
            .AddField(FormElementType.Label)
            .AddField(FormElementType.Label)
            .AddField(FormElementType.Radio)
            .AddField(FormElementType.Radio, 120)
            .AddField(FormElementType.Radio)
            .Field(1, 1).Extrude(FieldExtrudeSide.Left, 10)
            .Field(0, 0).Value = "Vi kan lage helt vilkårlige skjema med stor stilistisk frihet." & vbNewLine & " "
            With .Field(0, 1)
                .Header.BackColor = Color.DeepPink
                .Header.Text = "Eksempel"
                .Header.ForeColor = ColorHelper.Multiply(.Header.ForeColor, 0.4)
                .SecondaryValue = "Stilistisk frihet"
                .BackColor = Color.HotPink
                .Value = True
            End With
            .MergeWithAbove(1, 0)
            .MergeWithAbove(1, 1, -1)
            With .Field(2, 0)
                .Value = "Ayy"
                .Header.Text = "Faktisk skjema kommer snart"
            End With
            .ResumeLayout()
            '.HeightToContent()
        End With

        TestQuestionnaire = New Questionnaire(Me)
        With TestQuestionnaire
            .Width = 500
            .Height = 550
            .Top = 0
            .Add(TestForm1)
            .Add(TestForm2)
            .Display()
        End With
    End Sub
    Private Sub CalMouseEnter(Sender As CustomCalendar.CalendarDay) Handles Cal.MouseEnter
        Select Case Sender.Area
            Case CustomCalendar.CalendarArea.CurrentMonth
                Sender.BackColor = ColorHelper.Add(Sender.BackColor, 50)
        End Select
        Debug.Print(Sender.Day.ToShortDateString)
    End Sub
    Private Sub CalMouseLeave(Sender As CustomCalendar.CalendarDay) Handles Cal.MouseLeave
        Select Case Sender.Area
            Case CustomCalendar.CalendarArea.CurrentMonth
                Sender.BackColor = ColorHelper.Add(Sender.BackColor, -50)
        End Select
    End Sub
    Private Sub CalClick(Sender As CustomCalendar.CalendarDay) Handles Cal.Click
        Select Case Sender.Area
            Case CustomCalendar.CalendarArea.CurrentMonth
                Sender.BackColor = Color.Aqua
        End Select
        Dim Shad As New BoxShadow(Sender, Color.FromArgb(0, 50, 60))
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Cal.Dispose()
        TestForm1.Dispose()
    End Sub

    Private Sub cmdAdd_Click(sender As Object, e As EventArgs) Handles cmdAdd.Click
        Uni.Clear()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Uni.Air(New FastPoint(ClientSize.Width \ 2, ClientSize.Height \ 2), 20.0, 100)
    End Sub

    Private Sub Me_Resize() Handles MyBase.Resize
        If TestQuestionnaire IsNot Nothing Then
            With TestQuestionnaire
                .Left = (ClientRectangle.Width - .Width) \ 2
                .Top = (ClientRectangle.Height - .Height) \ 2
            End With
        End If
        With Uni
            .Size = ClientSize
        End With
    End Sub
End Class
