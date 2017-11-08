Option Strict On
Option Explicit On
Option Infer Off
Imports System.ComponentModel
Imports System.Threading
Imports AudiopoLib

Public Class Main
    Private IsLoaded As Boolean = False
    Public Sub New()
        ' Make sure the form is loaded once it's created (while the splash screen is showing) in order to prevent a delay when it's shown for the first time.
        ' Required:
        InitializeComponent()
        ' Raise the Load event:
        MyBase.OnLoad(New EventArgs)
    End Sub
    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Checks to see if the form is already initialized. If the form has simply been shown, there should be no initialization performed.
        If Not IsLoaded Then
            Hide()
            KeyPreview = True

            ' Creates a new MultiTabWindow instance (AudiopoLib) and assigns it to the Globals.Windows reference.
            Windows = New MultiTabWindow(Me)

            ' Initialize each tab/view in the project and specify the parent MultiTabWindow as the only argument.
            ' When switching tabs, the corresponding index must be supplied, so double-check that the number is correct.
            ' If tabs are added or removed out of order, all usage of the Index property must be updated. Such usage can
            ' be found in two ways: Parent.Index = X (in classes derived from Tab), or Windows.Index = X (elsewhere). 
            PersonaliaTest = New Personopplysninger(Windows) ' Index = 0
            LoggInnTab = New LoggInnNy(Windows) ' Index = 1
            Dashboard = New DashboardTab(Windows) ' Index = 2
            Egenerklæring = New EgenerklæringTab(Windows) ' Index = 3
            Timebestilling = New TimebestillingTab(Windows) ' Index = 4
            AnsattLoggInn = New AnsattLoggInnTab(Windows) ' Index 5
            OpprettAnsatt = New OpprettAnsattTab(Windows) ' Index 6
            AnsattDashboard = New AnsattDashboardTab(Windows) ' Index = 7
            RedigerProfil = New RedigerProfilTab(Windows) ' Index = 8
            With Windows
                .BackColor = Color.FromArgb(240, 240, 240)
                ' Switches to the tab at index 1 (LoggInnTab; see above).
                .Index = 1
            End With
            ' The form is now initialized and loaded. Do not initialize again.
            IsLoaded = True
        End If
    End Sub
    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        ' If closing the form would cause a memory leak, database corruption or other errors, set e.Cancel = True.
        ' We'll leave it as is for now.
        MyBase.OnFormClosing(e)
    End Sub
    Private Sub Me_Closing(Sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        ' If memory usage gets bad, explicitly and carefully dispose of resources before ending.
        End
    End Sub
End Class

''' <summary>
''' Inherits Tab. The sign-up tab for new donors.
''' </summary>
Public Class Personopplysninger
    Inherits Tab
    ' The FlatForm (AudiopoLib) for filling out personal information.
    ' We specify a width of 400, a height of 300 (the FlatForm automatically sets its height based on its contents), a field-spacing of 3 pixels,
    ' and the FormFieldStylePresets.PlainWhite (AudiopoLib) preset. Custom FormFieldStyles (AudiopoLib) can be created.
    Private Personalia As New FlatForm(400, 300, 3, FormFieldStylePresets.PlainWhite)
    ' A separate FlatForm with small alterations to the style (thus a custom style) for the password part of the sign-up view.
    Private PasswordForm As New FlatForm(270, 100, 3, New FormFieldStyle(Color.FromArgb(245, 245, 245), Color.FromArgb(70, 70, 70), Color.White, Color.FromArgb(80, 80, 80), Color.White, Color.Black, {True, True, True, True}, 20))
    Private WithEvents TopBar As New TopBar(Me)
    Private FormPanel As New BorderControl(Color.FromArgb(210, 210, 210))
    Private PicDoktor, PicDoktorPassord, PicOpprettKontoInfo, PicSuccess As New PictureBox
    Private FormInfo As New Label
    Private InfoLab As New InfoLabel
    Private WithEvents SendKnapp As New TopBarButton(FormPanel, My.Resources.NesteIcon, "Neste steg", New Size(0, 36))
    Private AvbrytKnapp As New TopBarButton(FormPanel, My.Resources.AvbrytIcon, "Avbryt", New Size(0, 36), True)
    Private NeiTakkKnapp As New TopBarButton(FormPanel, My.Resources.NeiTakkIcon, "Nei takk", New Size(0, 36))
    Private PasswordFormVisible As Boolean
    Private LayoutTool As New FormLayoutTools(Me)
    Private Footer As New Footer(Me)
    ' The DatabaseClient (AudiopoLib) that will take care of inserting the table entries
    Private WithEvents DBC As New DatabaseClient(Credentials.Server, Credentials.Database, Credentials.UserID, Credentials.Password)
    Private FirstHeader As New FullWidthControl(FormPanel)
    ' To store the information entered into the forms by the user
    Private FormResult(), PasswordResult() As String
    ' A NotificationManager (AudiopoLib) to display messages to the user.
    Private NotifManager As New NotificationManager(FirstHeader)
    ' Will be set to false if the user declines to create an account:
    Private CreateLogin As Boolean = True
    ''' <summary>
    ''' Raised when the TopBar has one of its buttons clicked.
    ''' </summary>
    ''' <param name="Sender">The TopBarButton that was clicked.</param>
    Private Sub TopBarButtonClick(Sender As TopBarButton, e As EventArgs) Handles TopBar.ButtonClick
        ' There is only one possible course of action. Whether there are two buttons (logout + "back") or just one (logout),
        ' the result is the same, so we do not need to check if Sender.IsLogout = True.
        Logout()
    End Sub
    ''' <summary>
    ''' Handles the Click event of the submit button.
    ''' </summary>
    Private Sub SendClick() Handles SendKnapp.Click
        ' If the password form is not visible, it means the personal information form has been filled out, and the "next" button
        ' (the same button as the submit button, but with different text and design) has been clicked, so we can go ahead and get
        ' the result of the forms and attempt to submit the information to the database.
        If Not PasswordFormVisible Then
            ' FlatForm.Validate (AudiopoLib) returns True if the requirements (IsNumeric, MaxLength, etc.) specified when the fields
            ' were created are met, or draws a red border around fields with illegal values and returns False otherwise.
            If Personalia.Validate Then
                ' Get the values of all input-type fields in the form of an array of HeaderValuePairs (AudiopoLib) containing the
                ' value of the header, the primary value (True or False for checkboxes, for example), and the secondary value
                ' (the text component of checkboxes, for example) if applicable.
                Dim Result() As HeaderValuePair = Personalia.Result
                ' Create a String array for storing the information to be passed to the parameterized query.
                Dim DataArr(11) As String
                ' The first five fields are text fields, so we can store these fields in the array without any special treatment,
                ' as the fields have already been validated. We do need to cast all of the values to their original data type, however,
                ' as all AudiopoLib classes are strongly typed. For text fields, we cast to String. DirectCast(obj, type) is faster
                ' than obj.ToString, and should be used when we are sure that the value is in fact a String type value.
                For i As Integer = 0 To 4
                    DataArr(i) = DirectCast(Result(i).Value, String)
                Next
                ' The 5th and 6th input fields are radio button type fields with Boolean values. In our database, Boolean values are
                ' either 0 or 1 (TinyInt(1) values), so we will not insert the String representation of the value directly ("True" or "False").
                ' 1 represents male and 0 represents female, meaning the leftmost radio button's value is True if male.
                'Also note that attempting to skip the conditional check by converting the Boolean value to an Integer and then converting to
                ' String might result in True being represented by -1 or similar (depending on the conversion method), so this is the safest approach.
                If DirectCast(Result(5).Value, Boolean) Then
                    DataArr(5) = "1"
                Else
                    DataArr(5) = "0"
                End If
                ' We do not check the value of the 6th input field, as its value is implied by the result of the previous field.
                ' The next 4 fields are text fields, so we can insert them directly.
                For i As Integer = 7 To 10
                    DataArr(i - 1) = DirectCast(Result(i).Value, String)
                Next
                ' We cast the values of the checkboxes to their original Boolean type and insert "1" for checked and "0" otherwise.
                If DirectCast(Result(11).Value, Boolean) Then
                    DataArr(10) = "1"
                Else
                    DataArr(10) = "0"
                End If
                If DirectCast(Result(12).Value, Boolean) Then
                    DataArr(11) = "1"
                Else
                    DataArr(11) = "0"
                End If
                ' Because this information will be used briefly outside this block, we assign the array to the FormResult field in the declarations.
                FormResult = DataArr
                ' Hide the visual elements specific to the first part (personal information) of the sign-up view.
                Personalia.Hide()
                PicDoktor.Hide()
                FormInfo.Hide()
                SendKnapp.Hide()
                NeiTakkKnapp.Hide()
                With AvbrytKnapp
                    .Hide()
                    ' Move the cancel button to make room for the "no thanks" button.
                    .Left = FormInfo.Left
                End With
                ' Where needed, apply changes to the elements while they're hidden. Another approach is to create controls with the appropriate
                ' styles and properties, and show/hide those controls instead of changing existing ones. This should be especially considered
                ' for controls whose fonts are changed, as font objects are apparently not immediately disposed of. This would involve separate
                ' handling of the Click event. For now, the font style is set to bold, and the UseCompatibleTextRendering is set to True, as
                ' the text looks irregularly compressed when drawn without it. The icon and location are also changed.
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
                ' Because the personal ID number is used as the account's username, get the first value of the previously created String array
                ' (the personal ID number/Norwegian social security number was requested in the first field of the personal information form)
                ' and assign it to the first field of the account creation form, which is a label and cannot be changed by the user.
                PasswordForm.Field(0, 0).Value = FormResult(0)
                ' Show the "decline" button to allow the user to register as a donor without without creating a user account. Show it.
                With NeiTakkKnapp
                    .Top = SendKnapp.Top
                    .Left = SendKnapp.Left - .Width - 10
                    .Show()
                End With
                ' Show the altered elements, as well as elements specific to the user account creation view.
                AvbrytKnapp.Show()
                SendKnapp.Show()
                PicOpprettKontoInfo.Show()
                PicDoktorPassord.Show()
                ' Show the account creation form and set the PasswordFormVisible field to True, so that clicking the submit button will result in the query being executed.
                PasswordForm.Show()
                PasswordFormVisible = True
            Else
                ' If an error occurred, it means the validation returned False, and the offending fields should now be displaying a red border.
                NotifManager.Display("Noe gikk galt. Vennligst forsikre deg om at skjemaet er fylt inn riktig.", NotificationPreset.OffRedAlert)
            End If
            ' If the account creation form is visible, it means the personal information form was validated, and the user is attempting to
            ' create the account. If the passwords are equal and within the allowed length, PasswordForm.Validate returns true.
            ' Known issue: The password fields do not display red borders if the minimum length requirement is not met, but the fields are equal.
        ElseIf PasswordForm.Validate Then
            ' If the passwords are equal and within the allowed length range, the personal information is sent to the database. The account is not created yet.
            DBC.Execute({"@fodselsnr", "@b_fornavn", "@b_etternavn", "@b_adresse", "@b_postnr", "@b_kjonn", "@b_telefon1", "@b_telefon2", "@b_telefon3", "@b_epost", "@send_epost", "@send_sms"}, FormResult)
        End If
    End Sub
    ''' <summary>
    ''' Handles the ListLoaded event of the DatabaseClient that inserts personal information.
    ''' </summary>
    ''' <param name="Sender">The DatabaseClient that finished </param>
    ''' <param name="e">Contains error information and an empty DataTable, as nothing has been selected.</param>
    Private Sub DBC_Finished(Sender As Object, e As DatabaseListEventArgs) Handles DBC.ListLoaded
        ' If an error occured, display a suitable message. This should not happen, and implies improper form validation.
        If e.ErrorOccurred Then
            NotifManager.Display("Noe gikk galt. Vennligst forsikre deg om at skjemaet er fylt inn riktig.", NotificationPreset.OffRedAlert)
        Else
            ' If the user chose to create an account, get the personal ID number and the chosen password from the second form and insert them into the
            ' user account table.
            If CreateLogin = True Then
                CreateLogin = False
                Dim Result() As HeaderValuePair = PasswordForm.Result
                PasswordResult = {FormResult(0), Result(1).Value.ToString}
                DBC.SQLQuery = "INSERT INTO Brukerkonto (b_fodselsnr, passord) VALUES (@fodselsnr, @passord);"
                DBC.Execute({"@fodselsnr", "@passord"}, {PasswordResult(0), PasswordResult(1)})
            Else
                ' Otherwise:
                PasswordForm.Hide()
                PicDoktorPassord.Hide()
                SendKnapp.Hide()
                AvbrytKnapp.Hide()
                NeiTakkKnapp.Hide()
                InfoLab.Hide()
                FormInfo.Hide()
                PicOpprettKontoInfo.Hide()
                PicSuccess.Show()
                NotifManager.Display("Du er nå registrert i vårt system.", NotificationPreset.GreenSuccess)
            End If
        End If
    End Sub
    ''' <summary>
    ''' Handles the ExecutionFailed event of the DatabaseClient (AudiopoLib) that inserts information. Implies a failed connection.
    ''' </summary>
    Private Sub DBC_Failed() Handles DBC.ExecutionFailed
        NotifManager.Display("Noe gikk galt. Vennligst forsikre deg om at du er koblet til internett.", NotificationPreset.OffRedAlert)
    End Sub
    Public Shadows Sub Show()
        ' An attempt at reducing flicker when showing the tab.
        FormPanel.Hide()
        MyBase.Show()
    End Sub
    Private Sub Me_VisibleChanged() Handles Me.VisibleChanged
        If Visible Then
            FormPanel.Show()
        End If
    End Sub
    ' Classes derived from Tab require a constructor that accepts a MultiTabWindow instance.
    Public Sub New(ParentWindow As MultiTabWindow)
        ' A call to MyBase.New(ParentWindow As MultiTabWindow) is required for classes derived from Tab.
        MyBase.New(ParentWindow)
        ' Suspend layout events while building the tab's contents.
        SuspendLayout()
        BackColor = Color.FromArgb(240, 240, 240)
        ' A BorderControl (AudiopoLib) that will contain forms and graphics related to account creation.
        With FormPanel
            .Hide()
            .Parent = Me
            ' The height and location of the TopBar is automatically set in its constructor
            .Location = New Point(30, TopBar.Bottom + 20)
            .Size = New Size(817, 480)
            .BackColor = Color.FromArgb(225, 225, 225)
        End With
        With FirstHeader
            .Size = New Size(817, 40)
            .Text = "Registrering"
            .BackColor = Color.FromArgb(183, 187, 191)
            .ForeColor = Color.White
        End With
#Region "Form"
        ' Build the first form (personal information questionnaire)
        With Personalia
            .NewRowHeight = 50
            ' We specify a width of 180. If no width is specified, the field fills the remaining available width, and a new row is created.
            .AddField(FormElementType.TextField, 180)
            With .Last
                .Header.Text = "Fødselsnummer* (11 siffer)"
                .Required = True
                .Numeric = True
                .MinLength = 11
                .MaxLength = 11
            End With
            ' Cast the FormField to the FormTextField class so we can access its TextBox and add an event handler for its TextChanged event,
            ' in which we will dynamically predict the sex of the user and automatically check the corresponding radio button.
            With DirectCast(.Last, FlatForm.FormTextField).TextField
                AddHandler .TextChanged, AddressOf CheckGender
            End With
            .AddField(FormElementType.TextField, 107)
            With .Last
                .Header.Text = "Fornavn*"
                .Required = True
                .MaxLength = 30
            End With
            ' This field will fill the remaining width, and the next field that's added will appear in a new row.
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
            ' Check the "male" radio button to bypass the need to check if the RadioButtonContext (AudiopoLib) has a checked radio button.
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
                ' Switch to dashed borders above
                .DrawBorder(FormField.ElementSide.Top) = False
                .DrawDashedSepararators(FormField.ElementSide.Top) = True
                ' Switch the header off (hide it)
                .SwitchHeader(False)
            End With
            ' Merge the field with the above field, so that they are neatly separated by a dim, dashed line.
            .MergeWithAbove(6, 0, 0, True)
            .Parent = FormPanel
            .Display()
            .Location = New Point(20, 60)
        End With
#End Region
#Region "Password Form"
        With PasswordForm
            .AddField(FormElementType.Label)
            With .Last
                .Header.Text = "Bruker-ID"
                .Value = "Fødselsnummer"
            End With
            .AddField(FormElementType.TextField)
            Dim FieldHeight As Integer
            With .Last
                .Header.Text = "Velg et passord"
                .DrawBorder(FormField.ElementSide.Bottom) = False
                .Required = True
                .MinLength = 8
                .MaxLength = 50
                ' Add handlers for the ValueChanged and ValidChanged events, so we can make sure that the red border is displayed for both
                ' password fields if it is displayed for one (happens when the values are unequal or illegal).
                AddHandler .ValueChanged, AddressOf PasswordChanged
                AddHandler .ValidChanged, AddressOf PasswordValidChanged
                FieldHeight = .Height - .Header.Bottom
            End With
            ' Cast the last added field to the FormTextField type (AudiopoLib) so that we can access its Placeholder property.
            With DirectCast(.Last, FlatForm.FormTextField)
                .PlaceHolder = "Minst 8 tegn"
                .TextField.UseSystemPasswordChar = True
            End With
            .AddField(FormElementType.TextField)
            .MergeWithAbove(2, 0)
            ' Add a "repeat password" field
            With .Last
                ' Handle the ValidChanged event with the same method as used for the "choose a password" field.
                AddHandler .ValidChanged, AddressOf PasswordValidChanged
                .DrawBorder(FormField.ElementSide.Top) = False
                .DrawDashedSepararators(FormField.ElementSide.Top) = True
                .SwitchHeader(False)
                .Height = FieldHeight
                .Required = True
                .RequireSpecificValue("")
            End With
            ' Cast the "repeat password" field to its original FormTextField type, so that we can access the Placeholder property,
            ' as there is no header in which to communicate the field's purpose.
            With DirectCast(.Last, FlatForm.FormTextField)
                .PlaceHolder = "Gjenta passordet"
                .TextField.UseSystemPasswordChar = True
            End With
            .Parent = FormPanel
            .Display()
        End With
#End Region
        TopBar.AddButton(My.Resources.HjemIcon, "Hjem", New Size(135, 36))
        SendKnapp.Location = New Point(Personalia.Right - SendKnapp.Width, Personalia.Bottom + 10)
        NeiTakkKnapp.Hide()
        With AvbrytKnapp
            .Location = New Point(SendKnapp.Left - .Width - 10, SendKnapp.Top)
            AddHandler .Click, AddressOf AvbrytKnapp_Klikk
        End With
        With PicDoktor
            .BackgroundImage = My.Resources.Doktor2
            .Size = .BackgroundImage.Size
            .Parent = FormPanel
            .Location = New Point(Personalia.Right + 20, Personalia.Bottom - .Height)
        End With
        With PicDoktorPassord
            .BackgroundImage = My.Resources.DoktorPassord
            .Size = .BackgroundImage.Size
            .Parent = FormPanel
            .Location = PicDoktor.Location
        End With
        With PasswordForm
            .Location = New Point((PicDoktorPassord.Left - .Width) \ 2, (PicDoktorPassord.Bottom - 210 - .Height) \ 2)
        End With
        With FormInfo
            .Parent = FormPanel
            .Location = New Point(Personalia.Left, Personalia.Bottom + 10)
            .AutoSize = False
            .Size = New Size(SendKnapp.Height, AvbrytKnapp.Left - .Left)
            .TextAlign = ContentAlignment.MiddleLeft
            .ForeColor = Color.FromArgb(80, 80, 80)
            .Text = "* markerer obligatoriske felt"
        End With
        With InfoLab
            .Parent = FormPanel
            .Location = New Point(PicDoktor.Left, PicDoktor.Bottom + 10)
            .Size = New Size(PicDoktor.Width, SendKnapp.Height)
            .Text = "Ved å registrere deg, samtykker du i at denne informasjonen blir lagret i våre systemer. Du kan når som helst slette disse opplysningene."
            .PanIn()
        End With
        With PicOpprettKontoInfo
            .Parent = FormPanel
            .Location = New Point(20, TopBar.Bottom)
            .BackgroundImage = My.Resources.OpprettKontoInfo
            .BackgroundImageLayout = ImageLayout.Center
            .Size = .BackgroundImage.Size
            .Hide()
        End With
        With PicSuccess
            .Hide()
            .Parent = FormPanel
            .Size = New Size(64, 64)
            .BackColor = Color.LimeGreen
            .Location = New Point((FormPanel.Width - .Width) \ 2, (FormPanel.Height - .Height) \ 2)
        End With
        DBC.SQLQuery = "INSERT INTO Blodgiver (b_fodselsnr, b_fornavn, b_etternavn, b_telefon1, b_telefon2, b_telefon3, b_epost, b_adresse, b_postnr, b_kjonn, send_epost, send_sms) VALUES (@fodselsnr, @b_fornavn, @b_etternavn, @b_telefon1, @b_telefon2, @b_telefon3, @b_epost, @b_adresse, @b_postnr, @b_kjonn, @send_epost, @send_sms);"
        FormPanel.Show()
        ResumeLayout()
    End Sub
    Private Sub CheckGender(Sender As Object, e As EventArgs)
        Dim Chars() As Char = DirectCast(Sender, TextBox).Text.ToCharArray
        If Chars.Length > 8 Then
            If (Convert.ToInt32(Chars(8)) Mod 2 = 1) Then
                Personalia.Field(2, 0).Value = True
            Else
                Personalia.Field(2, 1).Value = True
            End If
        End If
    End Sub
    Private Sub PasswordChanged(Sender As FormField, Value As Object)
        With PasswordForm.Field(1, 0)
            If .Validate Then
                PasswordForm.Field(2, 0).RequireSpecificValue(DirectCast(.Value, String))
            End If
        End With
    End Sub
    Private Sub PasswordValidChanged(Sender As FormField)
        With PasswordForm
            .Field(1, 0).IsValid = Sender.IsValid
            .Field(2, 0).IsValid = Sender.IsValid
        End With
    End Sub
    Private Sub AvbrytKnapp_Klikk(Sender As Object, e As EventArgs)
        ResetForm()
        Parent.Index = 1
    End Sub
    Public Overrides Sub ResetTab(Optional Arguments As Object = Nothing)
        MyBase.ResetTab(Arguments)
        ResetForm()
    End Sub
    Private Sub ResetForm()
        FormPanel.Hide()
        DBC.SQLQuery = "INSERT INTO Blodgiver (b_fodselsnr, b_fornavn, b_etternavn, b_telefon1, b_telefon2, b_telefon3, b_epost, b_adresse, b_postnr, b_kjonn, send_epost, send_sms) VALUES (@fodselsnr, @b_fornavn, @b_etternavn, @b_telefon1, @b_telefon2, @b_telefon3, @b_epost, @b_adresse, @b_postnr, @b_kjonn, @send_epost, @send_sms);"
        With PasswordForm
            .ClearAll(True)
            .Hide()
        End With
        With SendKnapp
            .BackColor = NeiTakkKnapp.BackColor
            .Label.Font = New Font(.Font, FontStyle.Regular)
            .Text = "Videre"
            .Top = Personalia.Bottom + 10
            .Left = Personalia.Right - .Width
        End With
        With NeiTakkKnapp
            .Hide()
        End With
        With AvbrytKnapp
            .Top = SendKnapp.Top
            .Left = SendKnapp.Left - .Width - 10
        End With
        PicDoktor.Show()
        PicDoktorPassord.Hide()
        PasswordForm.Hide()
        FormInfo.Show()
        PicOpprettKontoInfo.Hide()
        PicSuccess.Hide()
        FormResult = Nothing
        With Personalia
            .ClearAll()
            .Show()
            .Field(2, 0).Value = True
            .Field(5, 0).Value = True
            .Field(6, 0).Value = True
        End With
        PasswordFormVisible = False
    End Sub
    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            RemoveHandler PasswordForm.Field(1, 0).ValueChanged, AddressOf PasswordChanged
            RemoveHandler PasswordForm.Field(1, 0).ValidChanged, AddressOf PasswordValidChanged
            RemoveHandler PasswordForm.Field(2, 0).ValidChanged, AddressOf PasswordValidChanged
            RemoveHandler DirectCast(Personalia.Field(0, 0), FlatForm.FormTextField).TextField.TextChanged, AddressOf CheckGender
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
                .Location = New Point((Width - .Width) \ 2, (Height + TopBar.Bottom - Footer.Height - .Height) \ 2)
            End With
        End If
        ResumeLayout(True)
    End Sub
End Class

Public Class LoggInnNy
    Inherits Tab
    Private WithEvents Gear As New GearIcon
    Private LoginForm As New FlatForm(243, 100, 3, New FormFieldStyle(Color.FromArgb(245, 245, 245), Color.FromArgb(70, 70, 70), Color.White, Color.FromArgb(80, 80, 80), Color.White, Color.Black, {True, True, True, True}, 20))
    Private WithEvents TopBar As New TopBar(Me)
    Private FormPanel As New BorderControl(Color.FromArgb(210, 210, 210))
    Private PicSideInfo, PicInfoAbove, RightSide As New PictureBox
    Private FormInfo As New Label
    Private InfoLab As New InfoLabel
    Private WithEvents LoggInnKnapp As New TopBarButton(FormPanel, My.Resources.NesteIcon, "Logg inn", New Size(0, 36))
    Private WithEvents OpprettBrukerKnapp As New TopBarButton(FormPanel, My.Resources.RedigerProfilIcon, "Opprett bruker", New Size(0, 36))
    Private LayoutTool As New FormLayoutTools(Me)
    Private Footer As New Footer(Me)
    'Private WithEvents DBC As New DatabaseClient(Credentials.Server, Credentials.Database, Credentials.UserID, Credentials.Password)
    Private WithEvents UserLogin As New MySqlUserLogin(Credentials.Server, Credentials.Database, Credentials.UserID, Credentials.Password)
    Private FormHeader As New FullWidthControl(FormPanel)
    Private WithEvents NotifManager As New NotificationManager(FormHeader)
    Private LeftSide As New BorderControl(Color.FromArgb(210, 210, 210))
    Private PersonalNumber As String
    Private LoadingSurface As New PictureBox
    Private LG As New LoadingGraphics(Of PictureBox)(LoadingSurface)
    Private Sub NotificationOpened() Handles NotifManager.NotificationOpened
        Gear.Hide()
    End Sub
    Private Sub NotificationClosed() Handles NotifManager.NotificationClosed
        If NotifManager.IsReady Then
            Gear.Show()
        End If
    End Sub
    Private Sub Gear_Click() Handles Gear.Click
        ' TODO: Vis logintab
        Parent.Index = 5
    End Sub
    Private Sub LoginValid()
        Parent.Index = 2
        CurrentLogin = New UserInfo(PersonalNumber)
        Dashboard.Initiate()
        Egenerklæring.InitiateForm()
        LoginForm.ClearAll()
        For Each C As Control In FormPanel.Controls
            C.Show()
        Next
        LG.StopSpin()
    End Sub
    Private Sub LoginInvalid(ErrorOccurred As Boolean, ErrorMessage As String)
        For Each C As Control In FormPanel.Controls
            C.Show()
        Next
        LG.StopSpin()
        If ErrorOccurred Then
            NotifManager.Display("Noe gikk galt. Vennligst kontakt betjeningen.", NotificationPreset.OffRedAlert)
        Else
            NotifManager.Display("Brukernavnet eller passordet er feil.", NotificationPreset.OffRedAlert)
        End If
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
    Protected Overrides Sub OnDoubleClick(e As EventArgs)
        MyBase.OnDoubleClick(e)
    End Sub
    Public Sub New(Window As MultiTabWindow)
        MyBase.New(Window)
        With UserLogin
            .IfInvalid = AddressOf LoginInvalid
            .IfValid = AddressOf LoginValid
        End With
        DoubleBuffered = True
        SuspendLayout()
        BackColor = Color.FromArgb(240, 240, 240)
        With FormPanel
            .Hide()
            .Parent = Me
            .Top = TopBar.Bottom + 20
            .Left = 30
            .Width = 817
            .Height = 480
            .BackColor = Color.FromArgb(225, 225, 225)
        End With
        With FormHeader
            .Width = 817
            .Height = 40
            .Text = "Logg inn"
            .BackColor = Color.FromArgb(183, 187, 191)
            .ForeColor = Color.White
        End With
        With LeftSide
            .Parent = FormPanel
            .Size = New Size(FormPanel.Width \ 2, FormPanel.Height - FormHeader.Bottom - 1)
            .Top = FormHeader.Bottom
            .Left = 1
            .DrawBorder(FormField.ElementSide.Bottom) = False
            .DrawBorder(FormField.ElementSide.Top) = False
            .DrawBorder(FormField.ElementSide.Left) = False
        End With
        With RightSide
            .Parent = FormPanel
            .Location = New Point(LeftSide.Right, LeftSide.Top)
            .Size = New Size(FormPanel.Width - .Left - 1, FormPanel.Height - .Top - 1)
            .BackgroundImage = My.Resources.PicLoggInn
            .BackgroundImageLayout = ImageLayout.Center
        End With
#Region "LoginForm"
        With LoginForm
            .AddField(FormElementType.TextField)
            With .Last
                .Header.Text = "Fødselsnummer (11 siffer)"
                .Required = True
                .Numeric = True
                .MinLength = 11
                .MaxLength = 11
            End With
            With DirectCast(.Last, FlatForm.FormTextField)
                .PlaceHolder = "11 siffer"
            End With
            .AddField(FormElementType.TextField)
            With .Last
                .Header.Text = "Passord"
                .Required = True
                .MinLength = 6
                .MaxLength = 50
            End With
            With DirectCast(.Last, FlatForm.FormTextField)
                .PlaceHolder = "Passord"
                .TextField.UseSystemPasswordChar = True
            End With
            .Parent = LeftSide
            .Location = New Point(LeftSide.Width \ 2 - .Width \ 2, LeftSide.Height \ 2 - .Height \ 2)
            .Display()
        End With
#End Region
        With TopBar
            'AddHandler .Click, AddressOf 
        End With
        With OpprettBrukerKnapp
            .Parent = LeftSide
            .Left = LoginForm.Left
            .Top = LoginForm.Bottom + 10
        End With
        With LoggInnKnapp
            .Parent = LeftSide
            .Left = OpprettBrukerKnapp.Right + 10
            .Top = LoginForm.Bottom + 10
        End With
        With FormInfo
            .Parent = FormPanel
            .AutoSize = False
            .Height = LoggInnKnapp.Height
            .TextAlign = ContentAlignment.MiddleLeft
            .ForeColor = Color.FromArgb(80, 80, 80)
            .Text = "* markerer obligatoriske felt"
        End With
        With InfoLab
            .Parent = FormPanel
            .Height = LoggInnKnapp.Height
            .Width = PicSideInfo.Width
            .Text = "Ved å registrere deg, samtykker du i at denne informasjonen blir lagret i våre systemer. Du kan når som helst slette disse opplysningene."
        End With
        With PicInfoAbove
            .Parent = FormPanel
            .Top = TopBar.Bottom
            .Left = 20
            .BackgroundImage = My.Resources.OpprettKontoInfo
            .BackgroundImageLayout = ImageLayout.Center
            .Height = .BackgroundImage.Height
            .Width = .BackgroundImage.Width
            .Hide()
            '.MakeDashed(Color.Red)
        End With
        With Gear
            'a
            .Parent = FormHeader
            .Location = New Point(FormHeader.Width - .Width - 5, FormHeader.Height \ 2 - .Height \ 2)
        End With
        With LoadingSurface
            .Hide()
            .Size = New Size(50, 50)
            .Parent = FormPanel
            .Location = New Point((FormPanel.Width - .Width) \ 2, (FormPanel.Height + FormHeader.Bottom - .Height) \ 2)
        End With
        With LG
            .Stroke = 3
            .Pen.Color = Color.FromArgb(162, 25, 51)
        End With
        FormPanel.Show()
        ResumeLayout()
    End Sub
    Private Sub OpprettBruker_Click(sender As Object, e As EventArgs) Handles OpprettBrukerKnapp.Click
        Parent.Index = 0
    End Sub
    Private Sub LoggInn_Click(sender As Object, e As EventArgs) Handles LoggInnKnapp.Click
        For Each C As Control In FormPanel.Controls
            C.Hide()
        Next
        FormHeader.Show()
        LG.Spin(30, 10)
        PersonalNumber = LoginForm.Field(0, 0).Value.ToString
        ' TODO: Loading Graphics
        UserLogin.LoginAsync(LoginForm.Field(0, 0).Value.ToString, LoginForm.Field(1, 0).Value.ToString, "Brukerkonto", "b_fodselsnr", "passord")
    End Sub
    Private Sub PasswordChanged(Sender As FormField, Value As Object)
        LoginForm.Field(2, 0).RequireSpecificValue(LoginForm.Field(1, 0).Value.ToString)
    End Sub
    Private Sub PasswordValidChanged(Sender As FormField)
        With LoginForm
            .Field(1, 0).IsValid = Sender.IsValid
            .Field(2, 0).IsValid = Sender.IsValid
        End With
    End Sub
    Private Sub ResetForm()
        FormPanel.Hide()
        With LoginForm
            .ClearAll()
        End With
    End Sub
    Public Overrides Sub ResetTab(Optional Arguments As Object = Nothing)
        MyBase.ResetTab(Arguments)
        ResetForm()
    End Sub
    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            NotifManager.Dispose()
            LayoutTool.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub
    Protected Overrides Sub OnResize(e As EventArgs)
        SuspendLayout()
        MyBase.OnResize(e)
        If LayoutTool IsNot Nothing Then
            With FormPanel
                .Left = Width \ 2 - .Width \ 2
                .Top = TopBar.Bottom + (Height - TopBar.Bottom - Footer.Height) \ 2 - .Height \ 2
            End With
        End If
        ResumeLayout(True)
    End Sub
End Class

Public Class GearIcon
    Inherits PictureBox
    Private ImageSize As New Size(34, 34)
    Private SC As SynchronizationContext = SynchronizationContext.Current
    Private GearIcon As Bitmap = My.Resources.SettingsIcon
    Private varCurrentDegree As Integer = 0 ' TODO: Try Double
    Private varIsHovering As Boolean
    Private varIncrement As Integer = 5
    Private WithEvents SpinTimer As New Timers.Timer(1000 \ 30)
    Public Sub New()
        DoubleBuffered = True
        SpinTimer.AutoReset = False
        BackgroundImage = GearIcon
        BackgroundImageLayout = ImageLayout.Center
        Cursor = Cursors.Hand
        Size = New Size(34, 34)
    End Sub
    Protected Overrides Sub OnMouseEnter(e As EventArgs)
        MyBase.OnMouseEnter(e)
        varIsHovering = True
        varIncrement = 5
        SpinTimer.Start()
    End Sub
    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        varIsHovering = False
    End Sub
    Private Sub SpinTimer_Tick() Handles SpinTimer.Elapsed
        Dim returnBitmap As New Bitmap(ImageSize.Width, ImageSize.Height)
        Using g As Graphics = Graphics.FromImage(returnBitmap)
            With g
                .CompositingQuality = Drawing2D.CompositingQuality.HighQuality
                Dim OffsetSingle As Single = CSng(16.5)
                .TranslateTransform(OffsetSingle, OffsetSingle)
                .RotateTransform(varCurrentDegree)
                .TranslateTransform(-OffsetSingle, -OffsetSingle)
                .DrawImage(GearIcon, New Rectangle(Point.Empty, ImageSize))
            End With
        End Using
        SC.Post(AddressOf AdjustRotation, returnBitmap)
    End Sub
    Private Sub AdjustRotation(State As Object)
        Dim NewImage As Bitmap = DirectCast(State, Bitmap)
        varCurrentDegree += varIncrement
        If varCurrentDegree >= 360 Then varCurrentDegree = 0
        If varIsHovering Then
            BackgroundImage = NewImage
            SpinTimer.Start()
        ElseIf varCurrentDegree > 0 Then
            BackgroundImage = NewImage
            varIncrement += 2
            SpinTimer.Start()
        Else
            NewImage.Dispose()
            varIncrement = 5
            varCurrentDegree = 0
            BackgroundImage = GearIcon
        End If
    End Sub
    ' TODO: Add class for this in AudiopoLib
    Private Function RotateImage() As Bitmap
        Dim returnBitmap As New Bitmap(ImageSize.Width, ImageSize.Height)
        Using g As Graphics = Graphics.FromImage(returnBitmap)
            With g
                .CompositingQuality = Drawing2D.CompositingQuality.HighQuality
                .TranslateTransform(CSng(16.5), CSng(16.5))
                .RotateTransform(varCurrentDegree)
                .TranslateTransform(CSng(-16.5), CSng(-16.5))
                .DrawImage(GearIcon, New Rectangle(Point.Empty, ImageSize))
            End With
        End Using
        Return returnBitmap
    End Function
End Class

Public Class DashboardTab
    Inherits Tab
    Public WithEvents NotificationList As New UserNotificationContainer(Color.FromArgb(210, 210, 210))
    Private Header As New TopBar(Me)
    'Dim ScrollList As New Donasjoner(Me)
    Private WithEvents Beholder As New BlodBeholder(My.Resources.Tom_beholder, My.Resources.Full_beholder)
    Private WelcomeLabel As New InfoLabel(True, Direction.Right)

    Private OrganDonorInfo As New PictureBox
    Private IsLoaded As Boolean
    Private Messages As New MessageNotification(Header)
    Private WithEvents DBC, HentTimer_DBC, HentEgenerklæring_DBC, DBC_Delete As New DatabaseClient(Credentials.Server, Credentials.Database, Credentials.UserID, Credentials.Password)

    Private Sub DBC_HentTimer_Finished(Sender As Object, e As DatabaseListEventArgs) Handles HentTimer_DBC.ListLoaded
        TimeListe.Clear()
        If Not e.ErrorOccurred Then
            For Each Row As DataRow In e.Data.Rows
                Dim TimeID As Integer = DirectCast(Row.Item(0), Integer)
                Dim AnsattGodkjent As Boolean = DirectCast(Row.Item(1), Boolean)
                Dim Dato As Date = DirectCast(Row.Item(2), Date).Date.Add(DirectCast(Row.Item(3), TimeSpan))
                Dim AnsattID As Object = Row.Item(4)
                Dim Fødselsnummer As String = DirectCast(Row.Item(5), Int64).ToString
                Dim Fullført As Boolean = DirectCast(Row.Item(6), Boolean)
                Dim BlodgiverGodkjent As Boolean = DirectCast(Row.Item(7), Boolean)
                Dim NewTime As New StaffTimeliste.StaffTime(TimeID, Dato, AnsattGodkjent, Fødselsnummer, AnsattID, BlodgiverGodkjent)
                NewTime.Fullført = Fullført
                TimeListe.Add(NewTime)
                With NewTime
                    If Not .BlodgiverGodkjent AndAlso AnsattGodkjent Then
                        NotificationList.AddNotification("Du har mottatt en ny innkalling. Klikk her for nærmere opplysninger om dato og tid.", 0, AddressOf NyInnkalling, OffGreen, NewTime, DatabaseElementType.Time)
                    ElseIf Not .BlodgiverGodkjent AndAlso Not AnsattGodkjent AndAlso .AnsattID < 0 Then
                        NotificationList.AddNotification("Timeforespørselen din er til behandling.", 0, AddressOf CloseNotification, OffBlue, NewTime, DatabaseElementType.Time)
                    ElseIf Not AnsattGodkjent AndAlso .AnsattID >= 0 Then
                        NotificationList.AddNotification("Du har fått avslag på din timeforespørsel. Nærmere beskjed er sendt via epost.", 1, AddressOf DeleteRejected, OffRed, NewTime, DatabaseElementType.Time)
                    ElseIf AnsattGodkjent Then
                        If .DatoOgTid.Date.CompareTo(Date.Now.Date) > 0 Then
                            NotificationList.AddNotification("Du har en kommende time. Klikk her for nærmere informasjon.", 0, AddressOf NyInnkalling, OffBlue, NewTime, DatabaseElementType.Time)
                        End If
                    End If
                End With
            Next
            Timebestilling.SetAppointment()
            Dim AppointmentTodayID As Integer = -1
            For Each T As StaffTimeliste.StaffTime In TimeListe.Timer
                If T.DatoOgTid.Date = Date.Now.Date AndAlso T.BlodgiverGodkjent Then
                    AppointmentTodayID = T.TimeID
                    Exit For
                End If
            Next
            If AppointmentTodayID >= 0 Then
                With HentEgenerklæring_DBC
                    .SQLQuery = "SELECT * FROM Egenerklæring WHERE time_id = @id LIMIT 1;"
                    .Execute({"@id"}, {CStr(AppointmentTodayID)})
                End With
            Else
                With Dashboard.NotificationList
                    .Spin(False)
                End With
            End If
        Else
            ' TODO: Logg bruker ut med feilmelding
            Logout()
        End If
    End Sub
    Private Sub DeleteRejected(Sender As UserNotification, e As UserNotificationEventArgs)
        Dim Element As StaffTimeliste.StaffTime = DirectCast(Sender.RelatedElement, StaffTimeliste.StaffTime)
        With DBC_Delete
            .SQLQuery = "DELETE FROM Time WHERE time_id = @id;"
            .Execute({"@id"}, {Element.TimeID.ToString})
        End With
        Sender.Close()
    End Sub
    Private Sub NyInnkalling(Sender As UserNotification, e As UserNotificationEventArgs)
        Dim RelatedElement As StaffTimeliste.StaffTime = DirectCast(Sender.RelatedElement, StaffTimeliste.StaffTime)
        With Timebestilling
            Parent.Index = 4
            .Calendar.CurrentMonth = RelatedElement.DatoOgTid.Month
            .SelectDay(.Calendar.Day(RelatedElement.DatoOgTid.Date))
            Sender.Close()
        End With
    End Sub
    Private Sub DBC_Failed() Handles HentTimer_DBC.ExecutionFailed
        NotificationList.ShowMessage("Kunne ikke hente notifikasjoner og gjøremål. Trykk på ikonet over for å prøve på nytt. Hvis problemet fortsetter, vennligst logg ut og varsle personalet.", NotificationPreset.OffRedAlert)
    End Sub
    Private Sub Egenerklæring_Hentet(Sender As Object, e As DatabaseListEventArgs) Handles HentEgenerklæring_DBC.ListLoaded
        With e.Data.Rows
            If .Count > 0 Then
                With .Item(0)
                    Dim TimeID As Integer = DirectCast(.Item(0), Integer)
                    Dim SvarString As String = DirectCast(.Item(1), String)
                    Dim Land As String = DirectCast(.Item(2), String)
                    Dim Godkjent As Boolean = DirectCast(.Item(3), Boolean)
                    Dim AnsattSvar As String = Nothing
                    If Not IsDBNull(.Item(4)) Then
                        AnsattSvar = DirectCast(.Item(4), String)
                    End If
                    Dim NewElement As New Egenerklæringsliste.Egenerklæring(TimeID, SvarString, Land, Godkjent)
                    NewElement.AnsattSvar = AnsattSvar
                    CurrentLogin.FormInfo = NewElement
                    If AnsattSvar IsNot Nothing Then
                        Dashboard.NotificationList.AddNotification("Du har fått svar på din egenerklæring. Klikk her for mer informasjon.", 0, AddressOf FormAnswered, OffGreen, NewElement, DatabaseElementType.Egenerklæring)
                    Else
                        Dashboard.NotificationList.AddNotification("Din egenerklæring for dagens time er til behandling.", 1, AddressOf CloseNotification, OffGreen, NewElement, DatabaseElementType.Egenerklæring)
                    End If
                End With
            Else
                CurrentLogin.FormInfo = Nothing
                Dashboard.NotificationList.AddNotification("Du har ikke sendt egenerklæring for dagens time. Klikk her for å gå til skjemaet.", 2, AddressOf FormNotSent, OffRed, TimeListe.GetElementWhere(StaffTimeliste.TimeEgenskap.Dato, Date.Now.Date), DatabaseElementType.Time)
            End If
        End With
    End Sub
    Public Sub HentBrukerTimer() Handles NotificationList.Reloaded
        NotificationList.Clear()
        With HentTimer_DBC
            ' IF ERROR !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! UNCOMMENT
            '.SQLQuery = "SELECT * FROM Time WHERE b_fodselsnr = @nr AND NOT (a_id IS NOT NULL AND ansatt_godkjent = 0);"
            .SQLQuery = "SELECT * FROM Time WHERE b_fodselsnr = @nr;"
            .Execute({"@nr"}, {CurrentLogin.PersonalNumber})
        End With
    End Sub
    Private Sub TopBarButtonClick(Sender As TopBarButton, e As EventArgs)
        If Sender.IsLogout Then
            Logout()
        Else
            Select Case CInt(Sender.Tag)
                Case 0
                    Parent.Index = 4
                    With Timebestilling
                        .SelectDay(.Calendar.Day(Date.Now.Date))
                    End With
                Case 1
                    Windows.Index = 8
                    RedigerProfil.AutoFillOutForm(CurrentLogin.RelatedDonor)
            End Select
        End If
    End Sub
    Public Sub New(ParentWindow As MultiTabWindow)
        MyBase.New(ParentWindow)
        If Not IsLoaded Then
            With Header
                .AddButton(My.Resources.TimeBestillingIcon, "Bestill ny time", New Size(135, 36))
                .AddButton(My.Resources.RedigerProfilIcon, "Rediger profil", New Size(135, 36))
                .AddLogout("Logg ut", New Size(135, 36))
                AddHandler .ButtonClick, AddressOf TopBarButtonClick
            End With
            With Beholder
                .Parent = Me
                .Location = New Point(ClientSize.Width - .Width - 20, Header.Bottom + 20)
            End With
            With WelcomeLabel
                .ForeColor = Color.White
                .Parent = Header
                .Top = Header.Height \ 2 - .Height \ 2
                .Text = "Du er logget inn som..."
                .Height = Header.LogoutButton.Height - 3
            End With
            With NotificationList
                .Parent = Me
                .Location = New Point(20, Header.Bottom + 20)
            End With
            With OrganDonorInfo
                .BackgroundImage = My.Resources.infoTekstDashboard
                .Size = .BackgroundImage.Size
                .Parent = Me
            End With
            With Messages
                .Show()
                .Parent = Header
                .Left = ClientSize.Width - 500
                .Top = Header.Height \ 2 - .Height \ 2
                .BringToFront()
            End With
            IsLoaded = True
        End If
    End Sub
    Public Sub Initiate()
        With DBC
            .SQLQuery = "SELECT b_fornavn, b_etternavn FROM Blodgiver WHERE b_fodselsnr = @nr;"
            .Execute({"@nr"}, {CurrentLogin.PersonalNumber})
        End With
        Beholder.GetBlood()
    End Sub
    Private Sub DBC_Finished(Sender As Object, e As DatabaseListEventArgs) Handles DBC.ListLoaded
        If Not e.ErrorOccurred Then
            With e.Data
                CurrentLogin.FirstName = .Rows(0).Item(0).ToString
                CurrentLogin.LastName = .Rows(0).Item(1).ToString
                Header.RaiseNameSetEvent()
                WelcomeLabel.Text = "Du er logget inn som " & CurrentLogin.FirstName & " " & CurrentLogin.LastName
                WelcomeLabel.PanIn()
            End With
            NotificationList.Reload()
        Else
            NotificationList.ShowMessage("Kunne ikke hente informasjonen din, men du kan fortsatt bruke programmet.", NotificationPreset.OffRedAlert)
        End If
    End Sub
    'Public Shadows Sub Show()
    '    MyBase.Show()
    'End Sub
    Protected Overrides Sub OnResize(e As EventArgs)
        SuspendLayout()
        MyBase.OnResize(e)
        If IsLoaded Then
            With WelcomeLabel
                .Location = New Point(Width - 430, Header.LogoutButton.Top)
                .Size = New Size(300, Header.LogoutButton.Height - 3)
            End With
            With NotificationList
                .Top = (ClientSize.Height - .Height + Header.Bottom) \ 2
            End With
            With OrganDonorInfo
                .Location = New Point(ClientSize.Width - .Width - 20, NotificationList.Top)
            End With
            With Beholder
                .Location = New Point((NotificationList.Right + OrganDonorInfo.Left - .Width) \ 2, (ClientSize.Height - .Height + Header.Bottom) \ 2)
            End With
            With Messages
                .Left = ClientSize.Width - 500
                .Top = Header.Height \ 2 - .Height \ 2
            End With
        End If
        ResumeLayout(True)
    End Sub
    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            DBC.Dispose()
            HentEgenerklæring_DBC.Dispose()
            HentTimer_DBC.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class