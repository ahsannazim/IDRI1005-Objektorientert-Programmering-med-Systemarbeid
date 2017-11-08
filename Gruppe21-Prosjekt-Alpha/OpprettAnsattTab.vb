Imports AudiopoLib
Public Class OpprettAnsattTab
    Inherits Tab
    Private Personalia As New FlatForm(300, 300, 3, FormFieldStylePresets.PlainWhite)
    Private PasswordForm As New FlatForm(270, 100, 3, New FormFieldStyle(Color.FromArgb(245, 245, 245), Color.FromArgb(70, 70, 70), Color.White, Color.FromArgb(80, 80, 80), Color.White, Color.Black, {True, True, True, True}, 20))
    Private WithEvents TopBar As New TopBar(Me)
    Private FormPanel As New BorderControl(Color.FromArgb(210, 210, 210))
    Private WithEvents SendKnapp As New TopBarButton(FormPanel, My.Resources.NesteIcon, "Neste steg", New Size(0, 36),, 145)
    Private WithEvents AvbrytKnapp As New TopBarButton(FormPanel, My.Resources.AvbrytIcon, "Avbryt", New Size(0, 36), True, 145)
    Private PasswordFormVisible As Boolean = False
    Private Footer As New Footer(Me)
    Private WithEvents DBC As New DatabaseClient(Credentials.Server, Credentials.Database, Credentials.UserID, "")
    Private FirstHeader As New FullWidthControl(FormPanel)
    Private NotifManager As New NotificationManager(FirstHeader)
    Private CreateLogin As Boolean = True
    Private ValidityChecker As MySqlAdminLogin
    Private LoadingSurface As New PictureBox
    Private LG As LoadingGraphics(Of PictureBox)
    Private TSL As ThreadStarterLight
    Public Sub New(Window As MultiTabWindow)
        MyBase.New(Window)
        DoubleBuffered = True
        SuspendLayout()
        BackColor = Color.FromArgb(240, 240, 240)
        With FormPanel
            .Parent = Me
            .Size = New Size(408, 480)
            .BackColor = Color.FromArgb(225, 225, 225)
        End With
        With FirstHeader
            .Width = FormPanel.Width
            .Height = 40
            .Text = "Opprett ny profil for ansatte"
            .BackColor = Color.FromArgb(183, 187, 191)
            .ForeColor = Color.White
        End With
#Region "Form"
        With Personalia
            .NewRowHeight = 50
            .AddField(FormElementType.TextField, 97)
            With .Last
                .Header.Text = "Fødselsnummer"
                .Required = True
                .Numeric = True
                .MinLength = 11
                .MaxLength = 11
            End With
            With DirectCast(.Last, FlatForm.FormTextField)
                .PlaceHolder = "11 siffer"
            End With
            .AddField(FormElementType.TextField, 97)
            With .Last
                .Header.Text = "Fornavn"
                .Required = True
                .MaxLength = 30
            End With
            .AddField(FormElementType.TextField)
            With .Last
                .Header.Text = "Etternavn"
                .Required = True
                .MaxLength = 30
            End With
            .AddField(FormElementType.TextField, 197)
            With .Last
                .Header.Text = "Privatadresse"
                .Required = True
                .MaxLength = 100
            End With
            .AddField(FormElementType.TextField)
            With .Last
                .Header.Text = "Postnummer"
                .Required = True
                .Numeric = True
                .MinLength = 4
                .MaxLength = 4
            End With
            .NewRowHeight = 50
            .AddField(FormElementType.TextField, 147)
            With .Last
                .Header.Text = "Telefon privat"
                .Required = True
                .MinLength = 8
                .MaxLength = 15
            End With
            .AddField(FormElementType.TextField)
            With .Last
                .Header.Text = "Telefon mobil"
            End With
            .AddField(FormElementType.TextField)
            With .Last
                .Header.Text = "Epost-adresse"
                .Required = True
                .MinLength = 5
                .MaxLength = 100
            End With
            .Parent = FormPanel
            .Location = New Point(FormPanel.Width \ 2 - .Width \ 2, FormPanel.Height \ 2 - .Height \ 2)
            .Display()
        End With
#End Region
#Region "Password Form"
        With PasswordForm
            .AddField(FormElementType.TextField)
            With .Last
                .Header.Text = "Velg et brukernavn"
                .MaxLength = 50
                .Required = True
            End With
            With DirectCast(.Last, FlatForm.FormTextField)
                .PlaceHolder = "Spør administrator (maks 50)"
            End With
            .AddField(FormElementType.TextField)
            Dim FieldHeight As Integer
            With .Last
                .Header.Text = "Velg et passord"
                .DrawBorder(FormField.ElementSide.Bottom) = False
                .Required = True
                .MinLength = 8
                .MaxLength = 50
                AddHandler .ValueChanged, AddressOf PasswordChanged
                AddHandler .ValidChanged, AddressOf PasswordValidChanged
                FieldHeight = .Height - .Header.Bottom
            End With
            With DirectCast(.Last, FlatForm.FormTextField)
                .PlaceHolder = "Minst 8 tegn"
                .TextField.UseSystemPasswordChar = True
            End With
            .AddField(FormElementType.TextField)
            .MergeWithAbove(2, 0)
            With .Last
                AddHandler .ValidChanged, AddressOf PasswordValidChanged
                .DrawBorder(FormField.ElementSide.Top) = False
                .DrawDashedSepararators(FormField.ElementSide.Top) = True
                .SwitchHeader(False)
                .Height = FieldHeight
                .Required = True
                .RequireSpecificValue("")
            End With
            With DirectCast(.Last, FlatForm.FormTextField)
                .PlaceHolder = "Gjenta passordet"
                .TextField.UseSystemPasswordChar = True
            End With
            With .LastRow
                .Height = FieldHeight + 3
            End With
            .AddField(FormElementType.TextField)
            With .Last
                .Required = True
                .Header.Text = "Administratornøkkel"
            End With
            With DirectCast(.Last, FlatForm.FormTextField)
                .TextField.UseSystemPasswordChar = True
            End With
            .Parent = FormPanel
            .Hide()
            .Location = New Point(FormPanel.Width \ 2 - .Width \ 2, FormPanel.Height \ 2 - .Height \ 2)
            .Display()
        End With
#End Region
        With TopBar
            .AddButton(My.Resources.HjemIcon, "Tilbake", New Size(135, 36))
        End With
        With SendKnapp
            .Top = FormPanel.Height - .Height - 20
            .Left = Personalia.Right - .Width
        End With
        With AvbrytKnapp
            .Top = SendKnapp.Top
            .Left = Personalia.Left
        End With
        With LoadingSurface
            .Hide()
            .SendToBack()
            .Size = New Size(50, 50)
            .Parent = FormPanel
            .Location = New Point(FormPanel.Width \ 2 - .Width \ 2, FormPanel.Height \ 2 - .Height \ 2)
        End With
        LG = New LoadingGraphics(Of PictureBox)(LoadingSurface)
        With LG
            .Stroke = 3
            .Pen.Color = Color.FromArgb(230, 50, 80)
        End With
        FormPanel.Show()
        ResumeLayout()
    End Sub
    Private Sub TopBar_ButtonClick(Sender As TopBarButton, e As EventArgs) Handles TopBar.ButtonClick
        Select Case CInt(Sender.Tag)
            Case 0
                LG.StopSpin()
                Parent.Index = 1
                ResetForm()
        End Select
    End Sub
    Private Sub PasswordChanged(Sender As FormField, Value As Object)
        PasswordForm.Field(2, 0).RequireSpecificValue(PasswordForm.Field(1, 0).Value.ToString)
    End Sub
    Private Sub PasswordValidChanged(Sender As FormField)
        With PasswordForm
            .Field(1, 0).IsValid = Sender.IsValid
            .Field(2, 0).IsValid = Sender.IsValid
        End With
    End Sub
    Private Sub Opprett_Click(sender As Object, e As EventArgs) Handles SendKnapp.Click
        If Not PasswordFormVisible Then
            If Personalia.Validate Then
                Personalia.Hide()
                PasswordForm.Show()
                PasswordFormVisible = True
                SendKnapp.Text = "Registrer"
            Else
                NotifManager.Display("Skjemaet er ufullstendig eller uriktig utfylt.", NotificationPreset.OffRedAlert)
            End If
        Else
            If PasswordForm.Validate Then
                SuspendLayout()
                SendKnapp.Hide()
                PasswordForm.Hide()
                AvbrytKnapp.Hide()
                If TSL IsNot Nothing Then
                    TSL.Dispose()
                End If
                TSL = New ThreadStarterLight(AddressOf CredManager_Decode)
                With TSL
                    .WhenFinished = AddressOf CredManager_Finished
                    .Start(New String() {Application.StartupPath & "/cred/", "auth.txt", DirectCast(PasswordForm.Last.Value, String)})
                End With
                ResumeLayout(True)
                LG.Spin(30, 10)
            Else
                NotifManager.Display("Skjemaet er ufullstendig eller uriktig utfylt.", NotificationPreset.OffRedAlert)
            End If
        End If
    End Sub

    Private Function CredManager_Decode(Data As Object) As Object
        Dim CredManager As New AdminCredentials
        Dim ArgsArr() As String = DirectCast(Data, String())
        Dim CredPath As String = ArgsArr(0)
        Dim FileName As String = ArgsArr(1)
        Dim Key As String = ArgsArr(2)
        Dim GottenString As String = CredManager.Decode(CredPath, FileName, Key)
        Dim RetString() As String = Nothing
        If GottenString IsNot Nothing AndAlso GottenString <> "" Then
            RetString = CredManager.Decode(CredPath, FileName, Key).Split("%".ToCharArray)
        End If
        'CredManager.Dispose
        Return RetString
    End Function
    Private Sub CredManager_Finished(Decoded As Object)
        Dim DecodedString() As String = DirectCast(Decoded, String())
        If TSL IsNot Nothing Then
            TSL.Dispose()
        End If
        OnDecodeFinished((DecodedString IsNot Nothing AndAlso DecodedString(0) <> ""), DecodedString)
    End Sub
    Private Sub OnDecodeFinished(ByVal CorrectKey As Boolean, Data() As String)
        If CorrectKey Then
            If ValidityChecker IsNot Nothing Then
                ValidityChecker.Dispose()
            End If
            ValidityChecker = New MySqlAdminLogin(Data(0), Data(1))
            With ValidityChecker
                .WhenFinished = AddressOf OnCheckFinished
                .LoginAsync(Data(2), Data(3))
            End With
        Else
            SendKnapp.Enabled = True
            LG.StopSpin()
            With LoadingSurface
                .Hide()
                .SendToBack()
            End With
            PasswordForm.Show()
            SendKnapp.Show()
            AvbrytKnapp.Show()
            With PasswordForm.Last
                .Value = ""
                .Focus()
            End With
            NotifManager.Display("Nøkkelen var feil.", NotificationPreset.OffRedAlert)
        End If
    End Sub
    Private Sub OnCheckFinished(Valid As Boolean)
        If Valid Then
            With DBC
                Dim FirstResult() As HeaderValuePair = Personalia.Result
                Dim PasswordResult() As HeaderValuePair = PasswordForm.Result
                .SQLQuery = "INSERT INTO Ansatt (a_fodselsnr, a_fornavn, a_etternavn,  a_adresse, a_postnr, a_telefon, a_telefon2, a_epost, a_brukernavn, a_passord) VALUES (@fodselsnr, @fornavn, @etternavn, @adresse, @postnr, @telefon1, @telefon2, @epost, @brukernavn, @passord);"
                .Password = Credentials.Password
                Dim DataArr(9) As String
                For i As Integer = 0 To 7
                    DataArr(i) = DirectCast(FirstResult(i).Value, String)
                Next
                DataArr(8) = DirectCast(PasswordResult(0).Value, String)
                DataArr(9) = DirectCast(PasswordResult(1).Value, String)
                .Execute({"@fodselsnr", "@fornavn", "@etternavn", "@adresse", "@postnr", "@telefon1", "@telefon2", "@epost", "@brukernavn", "@passord"}, DataArr)
            End With
        Else
            PasswordForm.Show()
            SendKnapp.Show()
            AvbrytKnapp.Show()
            LG.StopSpin()
            NotifManager.Display("Kunne ikke koble til databasen.", NotificationPreset.OffRedAlert)
        End If
    End Sub
    Private Sub AvbrytKnapp_Klikk(Sender As Object, e As EventArgs) Handles AvbrytKnapp.Click
        Parent.Index = 5
        ResetTab()
    End Sub
    Private Sub ResetForm()
        With PasswordForm
            .ClearAll()
            .Hide()
        End With
        With Personalia
            .ClearAll()
            .Show()
        End With
        With SendKnapp
            .Text = "Neste"
        End With
        PasswordFormVisible = False
        DBC.SQLQuery = ""
        DBC.Password = ""
    End Sub
    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            RemoveHandler PasswordForm.Field(1, 0).ValueChanged, AddressOf PasswordChanged
            RemoveHandler PasswordForm.Field(1, 0).ValidChanged, AddressOf PasswordValidChanged
            RemoveHandler PasswordForm.Field(2, 0).ValidChanged, AddressOf PasswordValidChanged
            NotifManager.Dispose()
            LG.Dispose()
            TSL.Dispose()
            ValidityChecker.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub
    Protected Overrides Sub OnResize(e As EventArgs)
        SuspendLayout()
        MyBase.OnResize(e)
        If FormPanel IsNot Nothing Then
            With FormPanel
                .Left = (Width - .Width) \ 2
                .Top = TopBar.Bottom + (Height - Footer.Height - .Height - TopBar.Bottom) \ 2
            End With
            With LoadingSurface
                .Location = New Point((FormPanel.Width - .Width) \ 2, (FirstHeader.Bottom + FormPanel.Height - .Height) \ 2)
            End With
        End If
        ResumeLayout(True)
    End Sub
    Public Overrides Sub ResetTab(Optional Arguments As Object = Nothing)
        MyBase.ResetTab(Arguments)
        LG.StopSpin()
        SendKnapp.Enabled = True
        Personalia.Show()
        PasswordForm.Hide()
        SendKnapp.Show()
        AvbrytKnapp.Show()
        ResetForm()
    End Sub
    Private Sub DBC_Finished(Sender As Object, e As DatabaseListEventArgs) Handles DBC.ListLoaded
        LG.StopSpin()
        If e.ErrorOccurred Then
            PasswordForm.Show()
            SendKnapp.Show()
            AvbrytKnapp.Show()
            NotifManager.Display("Det oppsto en uventet feil.", NotificationPreset.OffRedAlert)
        Else
            Dim NewNotif As New Notification(FirstHeader, NotificationPreset.GreenSuccess, "Den nye brukeren er klar.", 2, AddressOf NotificationClosed)
            NewNotif.Display()
        End If
    End Sub
    Private Sub NotificationClosed(Sender As Notification)
        Sender.Dispose()
        Parent.Index = 5
        ResetTab()
    End Sub
    Private Sub DBC_Failed() Handles DBC.ExecutionFailed
        LG.StopSpin()
        PasswordForm.Show()
        SendKnapp.Show()
        AvbrytKnapp.Show()
        NotifManager.Display("Det oppsto en uventet feil.", NotificationPreset.OffRedAlert)
    End Sub
End Class
