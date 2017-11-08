Imports AudiopoLib

Public Class RedigerProfilTab
    Inherits Tab
    Private Personalia As New FlatForm(400, 300, 3, FormFieldStylePresets.PlainWhite)
    Private WithEvents TopBar As New TopBar(Me)
    Private FormPanel As New BorderControl(Color.FromArgb(210, 210, 210))
    Private FormInfo As New Label
    Private InfoLab As New InfoLabel
    Private WithEvents SendKnapp As New TopBarButton(FormPanel, My.Resources.NesteIcon, "Neste steg", New Size(0, 36))
    Private AvbrytKnapp As New TopBarButton(FormPanel, My.Resources.AvbrytIcon, "Avbryt", New Size(0, 36), True)
    Private LayoutTool As New FormLayoutTools(Me)
    Private Footer As New Footer(Me)
    Private WithEvents DBC As New DatabaseClient(Credentials.Server, Credentials.Database, Credentials.UserID, Credentials.Password)
    Private FirstHeader As New FullWidthControl(FormPanel)
    Private NotifManager As New NotificationManager(FirstHeader)
    Private Sub TopBarButtonClick(Sender As TopBarButton, e As EventArgs) Handles TopBar.ButtonClick
        ResetForm()
        If Not Sender.IsLogout Then
            ResetForm()
            Parent.Index = 2
        Else
            Logout()
        End If
    End Sub
    Private Sub SendClick() Handles SendKnapp.Click
        If Personalia.Validate Then
            Dim Result() As HeaderValuePair = Personalia.Result
            Dim DataArr(12) As String
            For i As Integer = 0 To 3
                DataArr(i + 1) = Result(i).Value.ToString
            Next
            If DirectCast(Result(4).Value, Boolean) Then
                DataArr(5) = "1"
            Else
                DataArr(5) = "0"
            End If
            For i As Integer = 6 To 9
                DataArr(i) = Result(i).Value.ToString
            Next
            If DirectCast(Result(10).Value, Boolean) Then
                DataArr(10) = "1"
            Else
                DataArr(10) = "0"
            End If
            If DirectCast(Result(11).Value, Boolean) Then
                DataArr(11) = "1"
            Else
                DataArr(11) = "0"
            End If
            If DirectCast(Result(12).Value, Boolean) Then
                DataArr(12) = "1"
            Else
                DataArr(12) = "0"
            End If
            DataArr(0) = CurrentLogin.PersonalNumber
            Personalia.Hide()
            FormInfo.Hide()
            SendKnapp.Hide()
            With SendKnapp
                .IconImage = My.Resources.OKIcon
                With .Label
                    .Text = "Opprett bruker"
                End With
                .Left = Personalia.Right - .Width
                With .Label
                    .Font = New Font(.Font, FontStyle.Bold)
                    .UseCompatibleTextRendering = True
                    .Text = .Text
                End With
            End With
            AvbrytKnapp.Show()
            SendKnapp.Show()
            Dim SendForm As Boolean
            If DirectCast(Personalia.Last.Value, Boolean) Then
                If MsgBox("Dette vil medføre at brukeren din og alle opplysningene dine vil bli permanent slettet. Du blir bedt om å bekrefte valget via epost/SMS.", MsgBoxStyle.YesNo, "Er du sikker?") = MsgBoxResult.Yes Then
                    SendForm = True
                End If
            Else
                SendForm = True
            End If
            If SendForm Then
                DBC.Execute({"@nr", "@b_fornavn", "@b_etternavn", "@b_adresse", "@b_postnr", "@b_kjonn", "@b_telefon1", "@b_telefon2", "@b_telefon3", "@b_epost", "@send_epost", "@send_sms", "@slett_meg"}, DataArr)
            Else
                Personalia.Last.Value = False
            End If
        Else
            NotifManager.Display("Noe gikk galt. Vennligst forsikre deg om at skjemaet er fylt inn riktig.", NotificationPreset.OffRedAlert)
        End If
    End Sub
    Private Sub DBC_Finished(Sender As Object, e As DatabaseListEventArgs) Handles DBC.ListLoaded
        If e.ErrorOccurred Then
            NotifManager.Display("Noe gikk galt. Vennligst forsikre deg om at skjemaet er fylt inn riktig.", NotificationPreset.OffRedAlert)
        Else
            NotifManager.Display("Endringene er sendt til evaluering.", NotificationPreset.GreenSuccess)
        End If
    End Sub
    Private Sub DBC_Failed() Handles DBC.ExecutionFailed
        NotifManager.Display("Noe gikk galt. Vennligst forsikre deg om at skjemaet er fylt ut riktig.", NotificationPreset.OffRedAlert)
    End Sub
    Public Shadows Sub Show()
        FormPanel.Hide()
        MyBase.Show()
    End Sub
    Private Sub Me_VisibleChanged() Handles Me.VisibleChanged
        If Visible Then
            FormPanel.Show()
        End If
    End Sub
    Public Sub AutoFillOutForm(RelatedDonor As Donor)
        With Personalia
            .Field(0, 0).Value = RelatedDonor.Fødselsnummer.ToString
            .Field(0, 1).Value = RelatedDonor.Fornavn
            .Field(0, 2).Value = RelatedDonor.Etternavn
            .Field(1, 0).Value = RelatedDonor.Adresse
            .Field(1, 1).Value = RelatedDonor.Postnummer.ToString
            If RelatedDonor.Hankjønn Then
                .Field(2, 0).Value = True
            Else
                .Field(2, 1).Value = True
            End If
            .Field(3, 0).Value = RelatedDonor.Telefon(0)
            .Field(3, 1).Value = RelatedDonor.Telefon(1)
            .Field(3, 2).Value = RelatedDonor.Telefon(2)
            .Field(4, 0).Value = RelatedDonor.Epost
            .Field(5, 0).Value = RelatedDonor.SendEpost
            .Field(6, 0).Value = RelatedDonor.SendSMS
            .Last.Value = False
        End With
    End Sub
    Public Sub New(Window As MultiTabWindow)
        MyBase.New(Window)
        DoubleBuffered = True
        SuspendLayout()
        BackColor = Color.FromArgb(240, 240, 240)
        With FormPanel
            .Hide()
            .Parent = Me
            .Left = 30
            .Width = 440
            .Height = 550
            .BackColor = Color.FromArgb(225, 225, 225)
        End With
        With FirstHeader
            .Width = FormPanel.Width
            .Height = 40
            .Text = "Endre/slett personopplysninger"
            .BackColor = Color.FromArgb(183, 187, 191)
            .ForeColor = Color.White
            .TextAlign = ContentAlignment.MiddleCenter
        End With
#Region "Form"
        With Personalia
            .NewRowHeight = 50
            .AddField(FormElementType.Label, 180)
            With .Last
                .Header.Text = "Fødselsnummer* (11 siffer)"
                .Value = "XXXXXXXXXXX"
            End With
            .AddField(FormElementType.TextField, 107)
            With .Last
                .Header.Text = "Fornavn*"
                .Required = True
                .MaxLength = 30
            End With
            .AddField(FormElementType.TextField)
            With .Last
                .Header.Text = "Etternavn*"
                .Required = True
                .MaxLength = 30
            End With
            .AddField(FormElementType.TextField, 290)
            With .Last
                .Header.Text = "Privatadresse*"
                .Required = True
                .MaxLength = 100
            End With
            .AddField(FormElementType.TextField)
            With .Last
                .Header.Text = "Postnummer*"
                .Required = True
                .Numeric = True
                .MinLength = 4
                .MaxLength = 4
            End With
            .NewRowHeight = 50
            .AddField(FormElementType.Radio, 200)
            With .Last
                .Value = True
                .Header.Text = "Kjønn*"
                .SecondaryValue = "Jeg er mann"
                .DrawBorder(FormField.ElementSide.Right) = False
            End With
            .AddField(FormElementType.Radio)
            With .Last
                .SecondaryValue = "Jeg er kvinne"
                .DrawBorder(FormField.ElementSide.Left) = False
                .DrawDashedSepararators(FormField.ElementSide.Left) = True
                .Extrude(FieldExtrudeSide.Left, 3)
                .DrawDotsOnHeader = False
            End With
            .AddField(FormElementType.TextField, 133)
            With .Last
                .Header.Text = "Telefon privat*"
                .Required = True
                .MaxLength = 15
            End With
            .AddField(FormElementType.TextField, 133)
            With .Last
                .Header.Text = "Telefon mobil"
            End With
            .AddField(FormElementType.TextField)
            With .Last
                .Header.Text = "Telefon arbeid"
            End With
            .AddField(FormElementType.TextField)
            With .Last
                .Header.Text = "Epost-adresse*"
                .Required = True
                .MinLength = 5
                .MaxLength = 100
            End With
            .AddField(FormElementType.CheckBox)
            .NewRowHeight = 40
            With .Last
                .Value = True
                .SecondaryValue = "Jeg ønsker å motta innkalling, påminnelser og informasjon via epost"
                .DrawBorder(FormField.ElementSide.Bottom) = False
            End With
            .AddField(FormElementType.CheckBox)
            With .Last
                .Value = True
                .SecondaryValue = "Jeg ønsker å motta innkalling, påminnelser og informasjon via SMS"
                .DrawBorder(FormField.ElementSide.Top) = False
                .DrawDashedSepararators(FormField.ElementSide.Top) = True
                .SwitchHeader(False)
            End With
            .MergeWithAbove(6, 0, 0, True)
            .NewRowHeight = 54
            .AddField(FormElementType.CheckBox)
            With .Last
                .Value = False
                .Header.Text = "Les nøye"
                .SecondaryValue = "Jeg ønsker å slette min bruker og tilhørende opplysninger."
            End With
            .Parent = FormPanel
            .Display()
            .Location = New Point((FormPanel.Width - .Width) \ 2, (FormPanel.Height + FirstHeader.Height - .Height) \ 2 - 20)
        End With
#End Region
        With TopBar
            .AddButton(My.Resources.HjemIcon, "Hjem", New Size(135, 36))
            'AddHandler .Click, AddressOf 
        End With
        With SendKnapp
            .Top = Personalia.Bottom + 10
            .Left = Personalia.Right - .Width
        End With
        With AvbrytKnapp
            .Top = SendKnapp.Top
            .Left = SendKnapp.Left - .Width - 10
            AddHandler .Click, AddressOf AvbrytKnapp_Klikk
        End With
        With FormInfo
            .Parent = FormPanel
            .Top = Personalia.Bottom + 10
            .Left = Personalia.Left
            .AutoSize = False
            .Height = SendKnapp.Height
            .Width = AvbrytKnapp.Left - .Left
            .TextAlign = ContentAlignment.MiddleLeft
            .ForeColor = Color.FromArgb(80, 80, 80)
            .Text = "* markerer obligatoriske felt"
        End With
        DBC.SQLQuery = "INSERT INTO Endringer (b_fødselsnr, b_fornavn, b_etternavn, b_telefon1, b_telefon2, b_telefon3, b_epost, b_adresse, b_postnr, b_kjonn, send_epost, send_sms, slett_meg) VALUES (@b_fornavn, @b_etternavn, @b_telefon1, @b_telefon2, @b_telefon3, @b_epost, @b_adresse, @b_postnr, @b_kjonn, @send_epost, @send_sms, @slett_meg);"
        FormPanel.Show()
        ResumeLayout()
    End Sub
    Private Sub PanInTest() Handles Me.DoubleClick
        InfoLab.PanIn()
    End Sub
    Private Sub AvbrytKnapp_Klikk(Sender As Object, e As EventArgs)
        ResetForm()
        Parent.Index = 2
    End Sub
    Private Sub ResetForm()
        FormPanel.Hide()
        FormInfo.Show()
        With Personalia
            .ClearAll()
        End With
    End Sub
    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            LayoutTool.Dispose()
            NotifManager.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub
    Protected Overrides Sub OnResize(e As EventArgs)
        SuspendLayout()
        MyBase.OnResize(e)
        ' TODO: Remove LayoutTool
        If LayoutTool IsNot Nothing Then
            With FormPanel
                .Left = Width \ 2 - .Width \ 2
                .Top = (Height + TopBar.Bottom - Footer.Height - .Height) \ 2
            End With
        End If
        ResumeLayout(True)
    End Sub
End Class

