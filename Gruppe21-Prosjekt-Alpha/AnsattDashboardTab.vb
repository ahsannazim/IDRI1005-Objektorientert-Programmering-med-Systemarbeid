Imports System.Threading
Imports AudiopoLib

Public Class AnsattDashboardTab
    Inherits Tab
    Private ViewContainer As New BorderControl(Color.FromArgb(220, 220, 220))
    Private RightContainer As New BorderControl(Color.FromArgb(220, 220, 220))
    Private ViewList As New MultiTabWindow(ViewContainer)
    Private RightViewList As New MultiTabWindow(RightContainer)
    Private T_View As New T_RightView(RightViewList)
    Private E_View As New E_RightView(RightViewList)
    Private RegistrerTappingView As New TappingView(RightViewList)
    Private LabView As New LabRapportView(RightViewList)

    Private Egenerklæringer As New Egenerklæringsliste
    Private Timer As New StaffTimeliste
    Public WithEvents T_NotificationList, E_NotificationList, Tapping_NotificationList, Lab_NotificationList As StaffNotificationContainer
    Private WithEvents Header As New TopBar(Me)
    Private Footer As New Footer(Me)
    'Dim ScrollList As New Donasjoner(Me)
    Private WelcomeLabel As New InfoLabel(True, Direction.Right)
    Private Messages As New MessageNotification(Header)
    Private WithEvents DBC_HentEgenerklæringer, DBC_HentTimer, DBC_HentKlareForTapping, DBC_HentLab As New DatabaseClient(Credentials.Server, Credentials.Database, Credentials.UserID, Credentials.Password)
    Private Sub NotificationList_Reloaded(Sender As Object, e As EventArgs) Handles T_NotificationList.Reloaded
        Timer.Clear()
        With DBC_HentTimer
            .SQLQuery = "SELECT * FROM Time WHERE (a_id = @aid AND ansatt_godkjent = 1 OR a_id IS NULL) AND fullført = 0 ORDER BY t_dato ASC LIMIT 10;"
            .Execute({"@aid"}, {CStr(CurrentStaff.ID)})
        End With
    End Sub
    Private Sub EgenerklæringList_Reloaded(Sender As Object, e As EventArgs) Handles E_NotificationList.Reloaded
        Egenerklæringer.Clear()
        With DBC_HentEgenerklæringer
            .SQLQuery = "SELECT A.* FROM Egenerklæring A INNER JOIN Time B ON A.time_id = B.time_id WHERE B.a_id = @aid AND A.svar IS NULL ORDER BY A.time_id DESC LIMIT 10;"
            .Execute({"@aid"}, {CStr(CurrentStaff.ID)})
        End With
    End Sub
    Private Sub TappingList_Reloaded(Sender As Object, e As EventArgs) Handles Tapping_NotificationList.Reloaded
        T_NotificationList.Reload()
        'With DBC_HentKlareForTapping
        '    .SQLQuery = "SELECT A* FROM Time A INNER JOIN Egenerklæring B ON A.time_id = B.time_id WHERE A.a_id = @aid AND B.godkjent = 1 AND A.t_dato = @dato AND fullført = 0 ORDER BY t_klokkeslett ASC LIMIT 10;"
        '    .Execute({"@aid", "@dato"}, {CStr(CurrentStaff.ID), Date.Now.ToString("yyyy-MM-dd")})
        'End With
    End Sub
    Private Sub LabList_Reloaded(Sender As Object, e As EventArgs) Handles Lab_NotificationList.Reloaded
        With DBC_HentLab
            .SQLQuery = "SELECT A.* FROM Time A INNER JOIN Blodtapping B ON A.time_id = B.time_id WHERE B.analysert = 0;"
            .Execute()
        End With
    End Sub
    Private Sub DBC_HentLab_Finished(Sender As Object, e As DatabaseListEventArgs) Handles DBC_HentLab.ListLoaded
        Lab_NotificationList.Spin(False)
        'If e.ErrorOccurred Then
        '    MsgBox("Error in lab: " & e.ErrorMessage)
        '    Lab_NotificationList.ShowMessage("Det oppsto en uventet feil." & vbNewLine & "Sjekk tilkoblingen.", NotificationPreset.OffRedAlert)
        'Else
        With e.Data
                If .Rows.Count > 0 Then
                    For Each Row As DataRow In .Rows
                        Dim TimeID As Integer = DirectCast(Row.Item(0), Integer)
                        Dim AnsattGodkjent As Boolean = DirectCast(Row.Item(1), Boolean)
                        Dim Dato As Date = DirectCast(Row.Item(2), Date).Date.Add(DirectCast(Row.Item(3), TimeSpan))
                        Dim AnsattID As Object = Row.Item(4)
                        Dim Fødselsnummer As String = DirectCast(Row.Item(5), Int64).ToString
                        Dim Fullført As Boolean = DirectCast(Row.Item(6), Boolean)
                        Dim BlodgiverGodkjent As Boolean = DirectCast(Row.Item(7), Boolean)
                        Dim NewTime As New StaffTimeliste.StaffTime(TimeID, Dato, AnsattGodkjent, Fødselsnummer, AnsattID, BlodgiverGodkjent)
                    NewTime.Fullført = Fullført
                    Lab_NotificationList.AddNotification("Blodmengde klar for behandling", 0, AddressOf SelectLab, OffBlue, NewTime, DatabaseElementType.Time)
                Next
                End If
            End With
        'End If
    End Sub
    Private Sub SelectLab(Sender As StaffNotification, e As StaffNotificationEventArgs)
        LabView.SelectNotification(Sender, e)
    End Sub
    Private Sub SelectTapping(Sender As StaffNotification, e As StaffNotificationEventArgs)
        RegistrerTappingView.SelectNotification(Sender, e)
    End Sub
    Private Sub TopBar_Click(Sender As TopBarButton, e As EventArgs) Handles Header.ButtonClick
        Logout(True)
    End Sub
    Public Overrides Sub ResetTab(Optional Arguments As Object = Nothing)
        MyBase.ResetTab(Arguments)
        Egenerklæringer.Clear()
        WelcomeLabel.Text = ""
        T_NotificationList.Clear()
        E_NotificationList.Clear()
        Tapping_NotificationList.Clear()

        ViewList.Index = 0
        RightViewList.Index = 0
    End Sub
    Public Sub GetData()
        T_NotificationList.Reload()
        E_NotificationList.Reload()
    End Sub
    Private Sub DBC_HentTimer_Ferdig(sender As Object, e As DatabaseListEventArgs) Handles DBC_HentTimer.ListLoaded
        T_NotificationList.Spin(False)
        If e.ErrorOccurred Then
            T_NotificationList.ShowMessage("Det oppsto en uventet feil." & vbNewLine & "Sjekk tilkoblingen.", NotificationPreset.OffRedAlert)
        Else
            With e.Data
                If .Rows.Count > 0 Then
                    For Each Row As DataRow In .Rows
                        Dim TimeID As Integer = DirectCast(Row.Item(0), Integer)
                        Dim AnsattGodkjent As Boolean = DirectCast(Row.Item(1), Boolean)
                        Dim Dato As Date = DirectCast(Row.Item(2), Date).Date.Add(DirectCast(Row.Item(3), TimeSpan))
                        Dim AnsattID As Object = Row.Item(4)
                        Dim Fødselsnummer As String = DirectCast(Row.Item(5), Int64).ToString
                        Dim Fullført As Boolean = DirectCast(Row.Item(6), Boolean)
                        Dim BlodgiverGodkjent As Boolean = DirectCast(Row.Item(7), Boolean)
                        Dim NewTime As New StaffTimeliste.StaffTime(TimeID, Dato, AnsattGodkjent, Fødselsnummer, AnsattID, BlodgiverGodkjent)
                        NewTime.Fullført = Fullført
                        Timer.Add(NewTime)
                        If NewTime.AnsattID <> CurrentStaff.ID Then
                            T_NotificationList.AddNotification("Timeforespørsel fra " & NewTime.Fødselsnummer.ToString, 0, AddressOf SelectTime, OffBlue, NewTime, DatabaseElementType.Time)
                        End If
                    Next
                End If
            End With
            Tapping_NotificationList.Clear()
            Dim MyAppointments As List(Of StaffTimeliste.StaffTime) = Timer.GetAllElementsWhere(StaffTimeliste.TimeEgenskap.AnsattID, CurrentStaff.ID)
            For Each T As StaffTimeliste.StaffTime In MyAppointments
                If T.BlodgiverGodkjent = True AndAlso T.DatoOgTid.Date = Date.Now.Date Then
                    Tapping_NotificationList.AddNotification("Klar for tapping: " & T.Fødselsnummer, 0, AddressOf SelectTapping, OffBlue, T, DatabaseElementType.Time)
                End If
            Next
            Tapping_NotificationList.Spin(False)
        End If
    End Sub
    Private Sub SelectTime(Sender As StaffNotification, e As StaffNotificationEventArgs)
        T_View.SelectTime(Sender, e)
    End Sub
    Private Sub DBC_HentEgenerklæringer_Finished(sender As Object, e As DatabaseListEventArgs) Handles DBC_HentEgenerklæringer.ListLoaded
        E_NotificationList.Spin(False)
        If e.ErrorOccurred Then
            E_NotificationList.ShowMessage("Det oppsto en uventet feil." & vbNewLine & "Sjekk tilkoblingen.", NotificationPreset.OffRedAlert)
        Else
            With e.Data
                If .Rows.Count > 0 Then
                    For Each Row As DataRow In .Rows
                        Dim TimeID As Integer = DirectCast(Row.Item(0), Integer)
                        Dim SvarString As String = DirectCast(Row.Item(1), String)
                        Dim Land As String = DirectCast(Row.Item(2), String)
                        Dim Godkjent As Boolean = DirectCast(Row.Item(3), Boolean)
                        Dim NewEgenerklæring As New Egenerklæringsliste.Egenerklæring(TimeID, SvarString, Land, Godkjent)
                        If Not IsDBNull(Row.Item(4)) Then
                            NewEgenerklæring.AnsattSvar = DirectCast(Row.Item(4), String)
                        End If
                        Egenerklæringer.Add(NewEgenerklæring)
                        E_NotificationList.AddNotification("Egenerklæring klar", 0, AddressOf SelectErklæring, OffBlue, NewEgenerklæring, DatabaseElementType.Egenerklæring)
                    Next
                End If
            End With
        End If
    End Sub
    Private Sub SelectErklæring(Sender As StaffNotification, e As StaffNotificationEventArgs)
        E_View.SelectErklæring(Sender, e)
    End Sub
    Public Sub New(ParentWindow As MultiTabWindow)
        MyBase.New(ParentWindow)
        With ViewContainer
            .Size = New Size(502, 502)
            .Left = 20
            .Parent = Me
        End With
        With RightContainer
            .Size = New Size(602, 502)
            .Left = ViewContainer.Right + 20
            .Parent = Me
        End With
        With ViewList
            .Size = New Size(500, 500)
            .Location = New Point(1, 1)
            .BackColor = Color.White
            .Show()
        End With
        With RightViewList
            .Size = New Size(600, 500)
            .Location = New Point(1, 1)
            .BackColor = Color.White
            .Show()
        End With
        With Header
            .AddLogout("Logg ut", New Size(0, 36))
        End With
        T_NotificationList = New StaffNotificationContainer(ViewList, "Timeforespørsler")
        E_NotificationList = New StaffNotificationContainer(ViewList, "Nye egenerklæringer")
        Tapping_NotificationList = New StaffNotificationContainer(ViewList, "Klare for tapping")
        Lab_NotificationList = New StaffNotificationContainer(ViewList, "Klare for analyse")
        RightViewList.Index = 0
        ViewList.Index = 0
    End Sub
    Public Sub SetRightIndex(ByVal Index As Integer)
        RightViewList.Index = Index
    End Sub
    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        If ViewContainer IsNot Nothing Then
            With ViewContainer
                .Top = (ClientSize.Height - .Height + Header.Bottom - Footer.Height) \ 2
                RightContainer.Top = .Top
            End With
        End If
    End Sub

    Private Class E_RightView
        Inherits Tab
        Private varSelectedNotification As StaffNotification
        Private WithEvents SendSvarKnapp As New FullWidthControl(Me, True, FullWidthControl.SnapType.Bottom)
        Private DonorInfoLabels(8), SkjemaSvar As Label
        Private Header As New FullWidthControl(Me)
        Private RedigerEgenerklæring, AutoSjekk As New BorderControl(Color.FromArgb(0, 100, 235))
        Private AcceptedCheckbox As New CheckBox
        Private SvarTextBox As New TextBox
        Private LoadingSurface As New PictureBox
        Private LG As LoadingGraphics(Of PictureBox)
        Private WithEvents DBC_Oppdater As New DatabaseClient(Credentials.Server, Credentials.Database, Credentials.UserID, Credentials.Password)
        Private WithEvents RedigerSkjemaLab, AutoSjekkLab As New Label
        Private IngentingValgtLab As New Label
        Private NotifManager As New NotificationManager(Header)
        Private PresetPics(4) As PictureBox
        Public Property SelectedNotification As StaffNotification
            Get
                Return varSelectedNotification
            End Get
            Set(value As StaffNotification)
                If varSelectedNotification IsNot Nothing Then
                    With varSelectedNotification
                        .BackColor = .DefaultColor
                        .IsSelected = False
                    End With
                End If
                varSelectedNotification = value
                If varSelectedNotification Is Nothing Then
                    SendSvarKnapp.Hide()
                    RedigerEgenerklæring.Hide()
                    AutoSjekk.Hide()
                    AcceptedCheckbox.Hide()
                    SvarTextBox.Hide()
                    For Each Lab As Label In DonorInfoLabels
                        Lab.Text = ""
                    Next
                    SkjemaSvar.Text = ""
                    IngentingValgtLab.Show()
                Else
                    With varSelectedNotification
                        .IsSelected = True
                        .BackColor = .PressColor
                    End With
                    SendSvarKnapp.Show()
                    RedigerEgenerklæring.Show()
                    AutoSjekk.Show()
                    AcceptedCheckbox.Show()
                    SvarTextBox.Show()
                    IngentingValgtLab.Hide()
                End If
            End Set
        End Property
        Public Sub New(ParentWindow As MultiTabWindow)
            MyBase.New(ParentWindow)
            With Header
                .Text = "Valgt egenerklæring"
            End With
            For i As Integer = 0 To DonorInfoLabels.Length - 1
                DonorInfoLabels(i) = New Label
                With DonorInfoLabels(i)
                    .Parent = Me
                    .AutoSize = False
                    .Size = New Size(300, 20)
                    .Location = New Point(20, Header.Bottom + 20 * (i + 1))
                End With
            Next
            SkjemaSvar = New Label
            With SkjemaSvar
                .Parent = Me
                .AutoSize = False
                .Size = New Size(500, 40)
                .Location = New Point(20, DonorInfoLabels(DonorInfoLabels.Count - 1).Bottom + 30)
            End With
            With AutoSjekk
                .Parent = Me
                .DrawBorder(FormField.ElementSide.Left) = False
                .DrawBorder(FormField.ElementSide.Top) = False
                .DrawBorder(FormField.ElementSide.Right) = False
                .DrawBorder(FormField.ElementSide.Bottom) = False
                .Size = TextRenderer.MeasureText("Utfør automatisk sjekk...", DefaultFont)
                .Height += 2
                .Location = New Point(20, SkjemaSvar.Bottom + 24)
            End With
            With AutoSjekkLab
                .AutoSize = False
                .Parent = AutoSjekk
                .Size = New Size(AutoSjekk.Width, AutoSjekk.Height - 1)
                .ForeColor = Color.FromArgb(0, 100, 235)
                .Text = "Utfør manuell sjekk..."
            End With
            With SendSvarKnapp
                .Size = New Size(130, 40)
                .Text = "Send svar"
                .BackColorNormal = Color.LimeGreen
                .BackColorSelected = ColorHelper.Multiply(.BackColorNormal, 0.7)
                .BackColorDisabled = Color.FromArgb(200, 200, 200)
                .ForeColor = Color.White
                .Left = 20
            End With
            With AcceptedCheckbox
                .Text = "Denne personen kan gi blod"
                .Parent = Me
                .Left = SendSvarKnapp.Right + 10
                .AutoSize = False
                .Width = 300
                .Height = 40
            End With
            With SvarTextBox
                .Parent = Me
                .Location = New Point(SendSvarKnapp.Left, SendSvarKnapp.Top - 60)
                .Multiline = True
                .WordWrap = True
                .Size = New Size(100, 50)
            End With
            With IngentingValgtLab
                .Hide()
                .ForeColor = Color.FromArgb(80, 80, 80)
                .Parent = Me
                .AutoSize = False
                .Size = New Size(100, 20)
                .TextAlign = ContentAlignment.MiddleCenter
                .Text = "Ingenting valgt"
            End With
            With LoadingSurface
                .Hide()
                .Parent = Me
                .Size = New Size(50, 50)
            End With
            For i As Integer = 0 To 4
                PresetPics(i) = New PictureBox
                With PresetPics(i)
                    .Size = New Size(15, 15)
                    .Parent = Me
                    .Tag = i
                    .Hide()
                    .Location = New Point(AutoSjekkLab.Left + 25 * i, AutoSjekkLab.Bottom + 10)
                    'AddHandler .Paint, AddressOf CirclePaint
                    'AddHandler .Click, AddressOf CheckString
                End With
            Next
            LG = New LoadingGraphics(Of PictureBox)(LoadingSurface)
            SelectedNotification = Nothing
        End Sub
        Private Sub CheckString(sender As Object, e As EventArgs)
            Dim PossibleChars() As Char = "01X".ToCharArray
            Dim Pattern() As String = {"X0001110000XXX00001111111", "X0001100100XX011101111111", "X00011001010X011101111111", "XXXXXXXXXXXXXXXXXXXXXXXXX", "0000000000000000000000000"}
            Dim CharArr() As Char = Pattern(CInt(DirectCast(sender, PictureBox).Tag)).ToCharArray
            Dim SvarArr() As Char = SkjemaSvar.Text.ToCharArray
            Dim Compatible As Boolean = True
            For i As Integer = 0 To 24
                Select Case CharArr(i)
                    Case PossibleChars(2)
                    Case Else
                        If SvarArr(i) <> CharArr(i) Then
                            Compatible = False
                            Exit For
                        End If
                End Select
            Next
            If Not Compatible Then
                MsgBox("Basert på det valgte mønsteret, kan ikke denne blodgiveren gi blod.")
            Else
                MsgBox("Basert på det valgte mønsteret, kan denne blodgiveren gi blod.")
            End If
        End Sub
        Private Sub CirclePaint(sender As Object, e As PaintEventArgs)
            Dim SenderX As PictureBox = DirectCast(sender, PictureBox)
            Dim Rect As Rectangle = SenderX.ClientRectangle
            Rect.Width -= 1
            Rect.Height -= 1
            Dim Colors() As Color = {Color.Green, Color.Blue, Color.Yellow, Color.Blue, Color.Black}
            Using NB As New SolidBrush(Colors(CInt(SenderX.Tag)))
                e.Graphics.FillEllipse(NB, Rect)
            End Using
        End Sub
        Private Sub DBC_Oppdater_Finished(sender As Object, e As DatabaseListEventArgs) Handles DBC_Oppdater.ListLoaded
            LG.StopSpin()
            If e.ErrorOccurred Then
                NotifManager.Display("Kunne ikke sende (uventet feil)", NotificationPreset.OffRedAlert, 3, FloatX.FillWidth, FloatY.FillHeight)
                IngentingValgtLab.Hide()
                AcceptedCheckbox.Show()
                SendSvarKnapp.Show()
                RedigerEgenerklæring.Show()
                AutoSjekk.Show()
                SvarTextBox.Show()
                SkjemaSvar.Show()
                For Each Lab As Label In DonorInfoLabels
                    Lab.Text = "Last inn på nytt..."
                Next
            Else
                AcceptedCheckbox.Hide()
                SendSvarKnapp.Hide()
                RedigerEgenerklæring.Hide()
                AutoSjekk.Hide()
                SvarTextBox.Hide()
                SkjemaSvar.Hide()
                IngentingValgtLab.Show()

                varSelectedNotification.Close()
                SelectedNotification = Nothing
                NotifManager.Display("Svaret er sendt", NotificationPreset.GreenSuccess, 3, FloatX.FillWidth, FloatY.FillHeight)
            End If
        End Sub
        Private Sub SendSvar_Click() Handles SendSvarKnapp.Click
            If AcceptedCheckbox.Checked Then
                If SvarTextBox.Text = "" Then SvarTextBox.Text = "Godkjent"
                DBC_Oppdater.SQLQuery = "UPDATE Egenerklæring SET svar = @svar, godkjent = 1 WHERE time_id = @timeid LIMIT 1;"
                With DirectCast(varSelectedNotification.RelatedElement, Egenerklæringsliste.Egenerklæring)
                    DBC_Oppdater.Execute({"@svar", "@timeid"}, {SvarTextBox.Text, .TimeID.ToString})
                End With
                ' TODO: Loading graphics
            Else
                If Not SvarTextBox.Text = "" Then
                    DBC_Oppdater.SQLQuery = "UPDATE Egenerklæring SET svar = @svar WHERE time_id = @timeid LIMIT 1;"
                    With DirectCast(varSelectedNotification.RelatedElement, Egenerklæringsliste.Egenerklæring)
                        DBC_Oppdater.Execute({"@svar", "@timeid"}, {SvarTextBox.Text, .TimeID.ToString})
                    End With
                    ' TODO: Slett time
                    ' TODO: Prevent changes while updating
                Else
                    SvarTextBox.Focus()
                    NotifManager.Display("Fyll ut mengde i liter", NotificationPreset.RedAlert,,, FloatY.FillHeight)
                End If
            End If
        End Sub
        Private Sub RedigerSkjemaLab_MouseEnter() Handles RedigerSkjemaLab.MouseEnter
            RedigerEgenerklæring.DrawBorder(FormField.ElementSide.Bottom) = True
        End Sub
        Private Sub RedigerSkjemaLab_MouseLeave() Handles RedigerSkjemaLab.MouseLeave
            RedigerEgenerklæring.DrawBorder(FormField.ElementSide.Bottom) = False
        End Sub
        Private Sub AutoSjekkLab_MouseEnter() Handles AutoSjekkLab.MouseEnter
            AutoSjekk.DrawBorder(FormField.ElementSide.Bottom) = True
        End Sub
        Private Sub AutoSjekkLab_MouseLeave() Handles AutoSjekkLab.MouseLeave
            AutoSjekk.DrawBorder(FormField.ElementSide.Bottom) = False
        End Sub

        Private Sub AutoSJekkLab_Click() Handles AutoSjekkLab.Click
            Dim TestString As String = InputBox("Skriv inn et mønster på 25 tegn (0, 1 eller X)", "Manuell sjekk")
            Dim PossibleChars() As Char = "01X".ToCharArray
            TestString.ToCharArray()
            Dim Compatible As Boolean = True
            Try
                Dim CharArr() As Char = TestString.ToCharArray()
                Dim SvarArr() As Char = SkjemaSvar.Text.ToCharArray
                For i As Integer = 0 To 24
                    Select Case CharArr(i)
                        Case PossibleChars(2)
                        Case Else
                            If SvarArr(i) <> CharArr(i) Then
                                Compatible = False
                                Exit For
                            End If
                    End Select
                Next
                If Not Compatible Then
                    MsgBox("Basert på det valgte mønsteret, kan ikke denne blodgiveren gi blod.")
                Else
                    MsgBox("Denne blodgiveren kan gi blod.")
                End If
            Catch
                MsgBox("Du må skrive inn 25 tegn.")
            End Try
        End Sub
        Public Sub SelectErklæring(Sender As StaffNotification, e As StaffNotificationEventArgs)
            SelectedNotification = Sender
            With Sender
                If .RelatedDonor IsNot Nothing Then
                    With .RelatedDonor
                        DonorInfoLabels(0).Text = "Fødselsnummer: " & .Fødselsnummer.ToString
                        DonorInfoLabels(1).Text = "Fornavn: " & .Fornavn
                        DonorInfoLabels(2).Text = "Etternavn: " & .Etternavn
                        DonorInfoLabels(3).Text = "Telefon: " & .Telefon(0)
                        DonorInfoLabels(4).Text = "Epostadresse: " & .Epost
                        DonorInfoLabels(5).Text = "Adresse: " & .Adresse
                        Dim PostnummerString As String = .Postnummer.ToString
                        If PostnummerString.Length < 4 Then
                            PostnummerString = "0" & PostnummerString
                        End If
                        DonorInfoLabels(6).Text = "Postnummer: " & PostnummerString
                        Dim KjønnString As String
                        If .Hankjønn Then
                            KjønnString = "Mann"
                        Else
                            KjønnString = "Kvinne"
                        End If
                        DonorInfoLabels(7).Text = "Kjønn: " & KjønnString
                        DonorInfoLabels(8).Text = "Blodtype: " & .Blodtype
                    End With
                    With DirectCast(.RelatedElement, Egenerklæringsliste.Egenerklæring)
                        SkjemaSvar.Text = "Sammendrag: " & .SkjemaString
                    End With
                Else
                    ' TODO: Loading graphics
                    LG.Spin(30, 10)
                End If
            End With
            Parent.Index = 1
        End Sub
        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                DBC_Oppdater.Dispose()
                LG.Dispose()
                NotifManager.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub
        Protected Overrides Sub OnResize(e As EventArgs)
            MyBase.OnResize(e)
            Header.Width = Width
            With SendSvarKnapp
                .Top = Height - .Height - 20
                .Left = 20
            End With
            With AcceptedCheckbox
                .Top = SendSvarKnapp.Top
            End With
            With IngentingValgtLab
                .Location = New Point((Width - .Width) \ 2, (Height + Header.Bottom - .Height) \ 2)
            End With
            With LoadingSurface
                .Location = New Point((Width - .Width) \ 2, (Height + Header.Bottom - .Height) \ 2)
            End With
        End Sub
    End Class

    Private Class T_RightView
        Inherits Tab
        Private varSelectedNotification As StaffNotification
        Private WithEvents GodkjennKnapp, AvslåKnapp As New FullWidthControl(Me, True, FullWidthControl.SnapType.Bottom)
        Private DonorInfoLabels(8), TimeInfoLabels(1) As Label
        Private Header As New FullWidthControl(Me)
        Private EndreDato, EndreKlokkeslett As New BorderControl(Color.FromArgb(0, 100, 235))
        Private WithEvents DBC_Oppdater As New DatabaseClient(Credentials.Server, Credentials.Database, Credentials.UserID, Credentials.Password)
        Private IngentingValgtLab As New Label
        Private WithEvents EndreDatoLab, EndreKlokkeslettLab As New Label
        Public Property SelectedNotification As StaffNotification
            Get
                Return varSelectedNotification
            End Get
            Set(value As StaffNotification)
                If varSelectedNotification IsNot Nothing Then
                    With varSelectedNotification
                        .BackColor = .DefaultColor
                        .IsSelected = False
                    End With
                End If
                varSelectedNotification = value
                If varSelectedNotification Is Nothing Then
                    GodkjennKnapp.Hide()
                    AvslåKnapp.Hide()
                    EndreDato.Hide()
                    EndreKlokkeslett.Hide()
                    For Each Lab As Label In DonorInfoLabels
                        Lab.Text = ""
                    Next
                    For Each Lab As Label In TimeInfoLabels
                        Lab.Text = ""
                    Next
                    IngentingValgtLab.Show()
                Else
                    GodkjennKnapp.Show()
                    AvslåKnapp.Show()
                    EndreDato.Show()
                    EndreKlokkeslett.Show()
                    IngentingValgtLab.Hide()
                    With varSelectedNotification
                        .IsSelected = True
                        .BackColor = .PressColor
                    End With
                End If
            End Set
        End Property
        Public Sub New(ParentWindow As MultiTabWindow)
            MyBase.New(ParentWindow)
            With Header
                .Text = "Valgt timeforespørsel"
            End With
            For i As Integer = 0 To DonorInfoLabels.Length - 1
                DonorInfoLabels(i) = New Label
                With DonorInfoLabels(i)
                    .Parent = Me
                    .AutoSize = False
                    .Size = New Size(300, 20)
                    .Location = New Point(20, Header.Bottom + 20 * (i + 1))
                End With
            Next
            For i As Integer = 0 To 1
                TimeInfoLabels(i) = New Label
                With TimeInfoLabels(i)
                    .Parent = Me
                    .AutoSize = False
                    .Size = New Size(300, 20)
                    .Location = New Point(320, Header.Bottom + 20 * (i + 1))
                End With
            Next
            With EndreDato
                .Parent = Me
                .DrawBorder(FormField.ElementSide.Left) = False
                .DrawBorder(FormField.ElementSide.Top) = False
                .DrawBorder(FormField.ElementSide.Right) = False
                .DrawBorder(FormField.ElementSide.Bottom) = False
                .Size = TextRenderer.MeasureText("Endre dato...", DefaultFont)
                .Location = New Point(320, Header.Bottom + 80)
            End With
            With EndreKlokkeslett
                .Parent = Me
                .DrawBorder(FormField.ElementSide.Left) = False
                .DrawBorder(FormField.ElementSide.Top) = False
                .DrawBorder(FormField.ElementSide.Right) = False
                .DrawBorder(FormField.ElementSide.Bottom) = False
                .Size = TextRenderer.MeasureText("Endre klokkeslett...", DefaultFont)
                .Location = New Point(320, Header.Bottom + 100)
            End With
            With EndreDatoLab
                .AutoSize = False
                .Parent = EndreDato
                .Size = New Size(EndreDato.Width, EndreDato.Height - 1)
                .ForeColor = Color.FromArgb(0, 100, 235)
                .Text = "Endre dato..."
            End With
            With EndreKlokkeslettLab
                .AutoSize = False
                .Parent = EndreKlokkeslett
                .Size = New Size(EndreKlokkeslett.Width, EndreKlokkeslett.Height - 1)
                .ForeColor = Color.FromArgb(0, 100, 235)
                .Text = "Endre klokkeslett..."
            End With
            With AvslåKnapp
                .Size = New Size(130, 40)
                .Text = "Avslå timeforespørsel"
                .BackColorNormal = Color.FromArgb(162, 25, 51)
                .BackColorSelected = ColorHelper.Multiply(.BackColorNormal, 0.7)
                .ForeColor = Color.White
                .Left = 20
            End With
            With GodkjennKnapp
                .Size = New Size(130, 40)
                .Text = "Send innkalling"
                .ForeColor = Color.White
                .BackColorNormal = Color.LimeGreen
                .BackColorSelected = ColorHelper.Multiply(Color.LimeGreen, 0.7)
                .Left = AvslåKnapp.Width + 40
            End With
            With IngentingValgtLab
                .Hide()
                .ForeColor = Color.FromArgb(80, 80, 80)
                .Parent = Me
                .AutoSize = False
                .Size = New Size(100, 20)
                .TextAlign = ContentAlignment.MiddleCenter
                .Text = "Ingenting valgt"
            End With
            SelectedNotification = Nothing
        End Sub
        Private Sub Avslå_Click() Handles AvslåKnapp.Click
            If MsgBox("Hvis det er kapasitet på en nærliggende dag, endre dato og tidspunkt i stedet for å avslå timeforespørselen. Blodgiveren vil bli automatisk informert.", MsgBoxStyle.OkCancel, "Forkast timeforespørsel") = MsgBoxResult.Ok Then
                DBC_Oppdater.SQLQuery = "UPDATE Time SET ansatt_godkjent = 0, a_id = @aid WHERE time_id = @timeid LIMIT 1;"
                With DirectCast(varSelectedNotification.RelatedElement, StaffTimeliste.StaffTime)
                    DBC_Oppdater.Execute({"@aid", "@timeid"}, {CurrentStaff.ID.ToString, .TimeID.ToString})
                End With
                ' TODO: Prevent changes while updating
            End If
        End Sub
        Private Sub DBC_Oppdater_Finished(sender As Object, e As DatabaseListEventArgs) Handles DBC_Oppdater.ListLoaded
            If e.ErrorOccurred Then

            Else
                varSelectedNotification.Close()
                SelectedNotification = Nothing
            End If
        End Sub
        Private Sub Godkjenn_Click() Handles GodkjennKnapp.Click
            DBC_Oppdater.SQLQuery = "UPDATE Time SET ansatt_godkjent = 1, a_id = @aid, t_dato = @dato, t_klokkeslett = @klokkeslett WHERE time_id = @timeid LIMIT 1;"
            With DirectCast(varSelectedNotification.RelatedElement, StaffTimeliste.StaffTime)
                DBC_Oppdater.Execute({"@aid", "@dato", "@klokkeslett", "@timeid"}, {CurrentStaff.ID.ToString, .DatoOgTid.ToString("yyyy-MM-dd"), .DatoOgTid.ToString("HH:mm"), .TimeID.ToString})
            End With
            ' TODO: Prevent changes while updating
        End Sub
        Private Sub EndreDatoLab_MouseEnter() Handles EndreDatoLab.MouseEnter
            EndreDato.DrawBorder(FormField.ElementSide.Bottom) = True
        End Sub
        Private Sub EndreDatoLab_MouseLeave() Handles EndreDatoLab.MouseLeave
            EndreDato.DrawBorder(FormField.ElementSide.Bottom) = False
        End Sub
        Private Sub EndreKlokkeslettLab_MouseEnter() Handles EndreKlokkeslettLab.MouseEnter
            EndreKlokkeslett.DrawBorder(FormField.ElementSide.Bottom) = True
        End Sub
        Private Sub EndreKlokkeslettLab_MouseLeave() Handles EndreKlokkeslettLab.MouseLeave
            EndreKlokkeslett.DrawBorder(FormField.ElementSide.Bottom) = False
        End Sub
        Private Sub EndreDato_Click() Handles EndreDatoLab.Click
            Try
                Dim NyDato As Date = Date.Parse(InputBox("Skriv inn ny dato (ÅÅÅÅ-MM-DD)", "Endre dato", DirectCast(varSelectedNotification.RelatedElement, StaffTimeliste.StaffTime).DatoOgTid.ToString("yyyy-MM-dd")))
                With DirectCast(varSelectedNotification.RelatedElement, StaffTimeliste.StaffTime)
                    .DatoOgTid = NyDato.Date.Add(.DatoOgTid.TimeOfDay)
                End With
                TimeInfoLabels(0).Text = "Dato: " & NyDato.ToShortDateString
            Catch
                MsgBox("Datoen ble skrevet inn feil")
            End Try
        End Sub
        Private Sub EndreKlokkeslettLab_Click() Handles EndreKlokkeslettLab.Click
            Try
                Dim NyttKlokkeslett As TimeSpan = TimeSpan.Parse(InputBox("Skriv inn nytt klokkeslett (HH:MM)", "Endre klokkeslett", DirectCast(varSelectedNotification.RelatedElement, StaffTimeliste.StaffTime).DatoOgTid.ToString("HH:mm")))
                With DirectCast(varSelectedNotification.RelatedElement, StaffTimeliste.StaffTime)
                    .DatoOgTid = .DatoOgTid.Date.Add(NyttKlokkeslett)
                    TimeInfoLabels(1).Text = "Klokkeslett: " & .DatoOgTid.ToString("HH:mm")
                End With
            Catch
                MsgBox("Datoen ble skrevet inn feil")
            End Try
        End Sub
        Public Sub SelectNone()
            SelectedNotification = Nothing
        End Sub
        Public Sub SelectTime(Sender As StaffNotification, e As StaffNotificationEventArgs)
            SelectedNotification = Sender
            With Sender
                If .RelatedDonor IsNot Nothing Then
                    With .RelatedDonor
                        DonorInfoLabels(0).Text = "Fødselsnummer: " & .Fødselsnummer.ToString
                        DonorInfoLabels(1).Text = "Fornavn: " & .Fornavn
                        DonorInfoLabels(2).Text = "Etternavn: " & .Etternavn
                        DonorInfoLabels(3).Text = "Telefon: " & .Telefon(0)
                        DonorInfoLabels(4).Text = "Epostadresse: " & .Epost
                        DonorInfoLabels(5).Text = "Adresse: " & .Adresse
                        Dim PostnummerString As String = .Postnummer.ToString
                        If PostnummerString.Length < 4 Then
                            PostnummerString = "0" & PostnummerString
                        End If
                        DonorInfoLabels(6).Text = "Postnummer: " & PostnummerString
                        Dim KjønnString As String
                        If .Hankjønn Then
                            KjønnString = "Mann"
                        Else
                            KjønnString = "Kvinne"
                        End If
                        DonorInfoLabels(7).Text = "Kjønn: " & KjønnString
                        DonorInfoLabels(8).Text = "Blodtype: " & .Blodtype
                    End With
                    With DirectCast(.RelatedElement, StaffTimeliste.StaffTime)
                        TimeInfoLabels(0).Text = "Dato: " & .DatoOgTid.ToShortDateString
                        TimeInfoLabels(1).Text = "Klokkeslett: " & .DatoOgTid.ToString("HH:mm")
                    End With
                Else

                End If
                Parent.Index = 0
            End With
        End Sub
        Protected Overrides Sub OnResize(e As EventArgs)
            MyBase.OnResize(e)
            Header.Width = Width
            With AvslåKnapp
                .Top = Height - .Height - 20
                .Left = 20
            End With
            With GodkjennKnapp
                .Top = Height - .Height - 20
            End With
            With IngentingValgtLab
                .Location = New Point((Width - .Width) \ 2, (Height + Header.Bottom - .Height) \ 2)
            End With
        End Sub
    End Class
End Class


Public Class TappingView
    Inherits Tab
    Private WithEvents Registrer As New FullWidthControl(Me, True, FullWidthControl.SnapType.Bottom)
    Private Header As New FullWidthControl(Me)

    Private DonorInfoLabels(8) As Label

    Private varSelectedNotification As StaffNotification

    Private NotifManager As NotificationManager
    Private RegistrerFravær, EndreBlodtype As New BorderControl(Color.FromArgb(0, 100, 235))
    Private WithEvents RegistrerFraværLab, EndreBlodtypeLab As New Label
    Private IngentingValgtLab As New Label
    Private TxtMengde As New TextBox
    Private LoadingSurface As New PictureBox
    Private LG As LoadingGraphics(Of PictureBox)
    Private WithEvents DBC_RegistrerTapping, DBC_RegistrerBlodtype, DBC_FullførTime As New DatabaseClient(Credentials.Server, Credentials.Database, Credentials.UserID, Credentials.Password)
    Public Sub SelectNotification(Sender As StaffNotification, e As StaffNotificationEventArgs)
        varSelectedNotification = Sender
        IngentingValgtLab.Hide()
        TxtMengde.Show()
        RegistrerFravær.Show()
        EndreBlodtype.Show()
        Registrer.Show()
        Parent.Index = 2
        With Sender
            If .RelatedDonor IsNot Nothing Then
                With .RelatedDonor
                    DonorInfoLabels(0).Text = "Fødselsnummer: " & .Fødselsnummer.ToString
                    DonorInfoLabels(1).Text = "Fornavn: " & .Fornavn
                    DonorInfoLabels(2).Text = "Etternavn: " & .Etternavn
                    DonorInfoLabels(3).Text = "Telefon: " & .Telefon(0)
                    DonorInfoLabels(4).Text = "Epostadresse: " & .Epost
                    DonorInfoLabels(5).Text = "Adresse: " & .Adresse
                    Dim PostnummerString As String = .Postnummer.ToString
                    If PostnummerString.Length < 4 Then
                        PostnummerString = "0" & PostnummerString
                    End If
                    DonorInfoLabels(6).Text = "Postnummer: " & PostnummerString
                    Dim KjønnString As String
                    If .Hankjønn Then
                        KjønnString = "Mann"
                    Else
                        KjønnString = "Kvinne"
                    End If
                    DonorInfoLabels(7).Text = "Kjønn: " & KjønnString
                    DonorInfoLabels(8).Text = "Blodtype: " & .Blodtype
                End With
                'With DirectCast(.RelatedElement, Egenerklæringsliste.Egenerklæring)
                '    SkjemaSvar.Text = "Sammendrag: " & .SkjemaString
                'End With
            Else

            End If
        End With
    End Sub
    Public Sub New(ParentWindow As MultiTabWindow)
        MyBase.New(ParentWindow)
        With Header
            .Text = "Registrer blodtapping"
        End With
        For i As Integer = 0 To DonorInfoLabels.Length - 1
            DonorInfoLabels(i) = New Label
            With DonorInfoLabels(i)
                .Parent = Me
                .AutoSize = False
                .Size = New Size(300, 20)
                .Location = New Point(20, Header.Bottom + 20 * (i + 1))
            End With
        Next
        With RegistrerFravær
            .Parent = Me
            .DrawBorder(FormField.ElementSide.Left) = False
            .DrawBorder(FormField.ElementSide.Top) = False
            .DrawBorder(FormField.ElementSide.Right) = False
            .DrawBorder(FormField.ElementSide.Bottom) = False
            .Size = TextRenderer.MeasureText("Registrer fravær...", DefaultFont)
            .Height += 2
            .Location = New Point(320, Header.Bottom + 50)
        End With
        With RegistrerFraværLab
            .AutoSize = False
            .Parent = RegistrerFravær
            .Size = New Size(RegistrerFravær.Width, RegistrerFravær.Height - 1)
            .Height += 2
            .ForeColor = Color.FromArgb(0, 100, 235)
            .Text = "Registrer fravær..."
        End With
        With EndreBlodtype
            .Parent = Me
            .DrawBorder(FormField.ElementSide.Left) = False
            .DrawBorder(FormField.ElementSide.Top) = False
            .DrawBorder(FormField.ElementSide.Right) = False
            .DrawBorder(FormField.ElementSide.Bottom) = False
            .Size = TextRenderer.MeasureText("Endre blodtype...", DefaultFont)
            .Location = New Point(320, Header.Bottom + 80)
        End With
        With EndreBlodtypeLab
            .AutoSize = False
            .Parent = EndreBlodtype
            .Size = New Size(EndreBlodtype.Width, EndreBlodtype.Height - 1)
            .ForeColor = Color.FromArgb(0, 100, 235)
            .Text = "Endre blodtype..."
        End With
        With Registrer
            .Size = New Size(130, 40)
            .Text = "Registrer"
            .BackColorNormal = Color.LimeGreen
            .BackColorSelected = ColorHelper.Multiply(.BackColorNormal, 0.7)
            .BackColorDisabled = Color.FromArgb(200, 200, 200)
            .ForeColor = Color.White
            .Left = 20
        End With
        With TxtMengde
            .Parent = Me
            .Location = New Point(200, 300)
            .Width = 100
        End With
        With IngentingValgtLab
            .Hide()
            .ForeColor = Color.FromArgb(80, 80, 80)
            .Parent = Me
            .AutoSize = False
            .Size = New Size(100, 20)
            .TextAlign = ContentAlignment.MiddleCenter
            .Text = "Ingenting valgt"
        End With
        With LoadingSurface
            .Hide()
            .Parent = Me
            .Size = New Size(50, 50)
        End With
        NotifManager = New NotificationManager(Header)
        LG = New LoadingGraphics(Of PictureBox)(LoadingSurface)
    End Sub
    Private Sub DBC_Registrer_Finished(sender As Object, e As DatabaseListEventArgs) Handles DBC_RegistrerTapping.ListLoaded
        LG.StopSpin()
        If e.ErrorOccurred Then
            NotifManager.Display("Kunne ikke registrere tapping.", NotificationPreset.OffRedAlert)
            TxtMengde.Show()
            RegistrerFravær.Show()
            EndreBlodtype.Show()
            Registrer.Show()
        Else
            NotifManager.Display("Tapping registrert.", NotificationPreset.GreenSuccess, 2, , FloatY.FillHeight)
            ' Oppdater fullført til 1
            DBC_FullførTime.SQLQuery = "UPDATE Time SET fullført = 1 WHERE time_id = @timeid;"
            DBC_FullførTime.Execute({"@timeid"}, {DirectCast(varSelectedNotification.RelatedElement, StaffTimeliste.StaffTime).TimeID.ToString})
            varSelectedNotification.Close()
            varSelectedNotification = Nothing
            TxtMengde.Hide()
            RegistrerFravær.Hide()
            EndreBlodtype.Hide()
            Registrer.Hide()
            For Each Lab As Label In DonorInfoLabels
                Lab.Text = ""
            Next
            IngentingValgtLab.Show()
        End If
    End Sub
    Private Sub FullførtRegistrert(Sender As Object, e As DatabaseListEventArgs) Handles DBC_FullførTime.ListLoaded
        NotifManager.Display("Timen er markert som fullført.", NotificationPreset.GreenSuccess,,, FloatY.FillHeight)
    End Sub
    Private Sub RegistrerTapping_Click() Handles Registrer.Click
        Dim ResultDouble As New Double
        If TxtMengde.Text = "" OrElse Not IsNumeric(TxtMengde.Text) OrElse Not Double.TryParse(TxtMengde.Text, ResultDouble) Then
            TxtMengde.Focus()
            Beep()
        Else
            With varSelectedNotification.RelatedDonor
                Dim UpdateBlood As Boolean
                If .Blodtype Is Nothing OrElse .Blodtype = "Ukjent" OrElse .Blodtype = "" Then
                    .Blodtype = InputBox("Blodtype:", "Blodtype ukjent", Nothing)
                    DonorInfoLabels(8).Text = "Blodtype: " & .Blodtype
                    UpdateBlood = True
                End If
                DBC_RegistrerTapping.SQLQuery = "INSERT INTO Blodtapping (ant_liter, time_id, blodtype) VALUES (@liter, @time_id, @blodtype);"
                With DirectCast(varSelectedNotification.RelatedElement, StaffTimeliste.StaffTime)
                    DBC_RegistrerTapping.Execute({"@liter", "@time_id", "@blodtype"}, {Double.Parse(TxtMengde.Text).ToString, .TimeID.ToString, varSelectedNotification.RelatedDonor.Blodtype})
                End With
                If UpdateBlood Then
                    DBC_RegistrerBlodtype.SQLQuery = "UPDATE Blodgiver SET blodtype = @blodtype WHERE b_fodselsnr = @nr;"
                    DBC_RegistrerBlodtype.Execute({"@blodtype", "@nr"}, { .Blodtype, .Fødselsnummer.ToString})
                End If
                TxtMengde.Hide()
                RegistrerFravær.Hide()
                EndreBlodtype.Hide()
                Registrer.Hide()
                For Each Lab As Label In DonorInfoLabels
                    Lab.Text = ""
                Next
                LG.Spin(30, 10)
            End With
        End If
    End Sub
    Private Sub RegistrerFraværLab_MouseEnter() Handles RegistrerFraværLab.MouseEnter
        RegistrerFravær.DrawBorder(FormField.ElementSide.Bottom) = True
    End Sub
    Private Sub RegistrerFraværLab_MouseLeave() Handles RegistrerFraværLab.MouseLeave
        RegistrerFravær.DrawBorder(FormField.ElementSide.Bottom) = False
    End Sub
    Private Sub EndreBlodtypeLab_MouseEnter() Handles EndreBlodtypeLab.MouseEnter
        EndreBlodtype.DrawBorder(FormField.ElementSide.Bottom) = True
    End Sub
    Private Sub EndreBlodtypeLab_MouseLeave() Handles EndreBlodtypeLab.MouseLeave
        EndreBlodtype.DrawBorder(FormField.ElementSide.Bottom) = False
    End Sub
    Private Sub RegistrerFravær_Click() Handles RegistrerFraværLab.Click
        DBC_RegistrerTapping.SQLQuery = "INSERT INTO Blodtapping (ant_liter, time_id, blodtype) VALUES (@liter, @time_id, @blodtype);"
        DBC_RegistrerTapping.Execute({"@liter", "@time_id", "@blodtype"}, {CStr(0), DirectCast(varSelectedNotification.RelatedElement, StaffTimeliste.StaffTime).TimeID.ToString, varSelectedNotification.RelatedDonor.Blodtype})
        NotifManager.Display("Blodtappingen registreres som fravær.", NotificationPreset.GreenSuccess,,, FloatY.FillHeight)
    End Sub
    Private Sub EndreBlodtype_Click() Handles EndreBlodtypeLab.Click
        With varSelectedNotification.RelatedDonor
            .Blodtype = InputBox("Blodtype:", "Blodtype ukjent", "O")
            DonorInfoLabels(8).Text = "Blodtype: " & .Blodtype
            DBC_RegistrerBlodtype.SQLQuery = "UPDATE Blodgiver SET blodtype = @blodtype WHERE b_fodselsnr = @nr;"
            DBC_RegistrerBlodtype.Execute({"@blodtype", "@nr"}, { .Blodtype, .Fødselsnummer.ToString})
        End With
    End Sub
    Private Sub DBC_Blodtype(Sender As Object, e As DatabaseListEventArgs) Handles DBC_RegistrerBlodtype.ListLoaded
        If e.ErrorOccurred Then
            NotifManager.Display("Blodtype kunne ikke registreres på grunn av en uventet feil.", NotificationPreset.OffRedAlert)
        Else
            NotifManager.Display("Blodtappingen ble registrert.", NotificationPreset.GreenSuccess)
        End If
    End Sub
    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        Header.Width = Width
        With Registrer
            .Top = Height - .Height - 20
            .Left = 20
        End With
        With IngentingValgtLab
            .Location = New Point((Width - .Width) \ 2, (Height + Header.Bottom - .Height) \ 2)
        End With
        With LoadingSurface
            .Location = New Point((Width - .Width) \ 2, (Height + Header.Bottom - .Height) \ 2)
        End With
    End Sub
End Class


Public Class LabRapportView
    Inherits Tab
    Private SC As SynchronizationContext = SynchronizationContext.Current
    Private WithEvents Registrer As New FullWidthControl(Me, True, FullWidthControl.SnapType.Bottom)
    Private Header As New FullWidthControl(Me)
    Private varSelectedNotification As StaffNotification
    Private NotifManager As NotificationManager
    Private IngentingValgtLab As New Label
    Private LoadingSurface As New PictureBox

    Private RBColumn, TromboColumn, PlasmaColumn As New PictureBox

    Private X As Integer = 0
    Private Xmax As Integer = 50
    Private EaseInOut As New EaseInOut
    Private WithEvents varEaseTimer As New Timers.Timer(1000 / 50)

    Private Verdier(2) As Double
    Private Blodtype As String
    Private varTotalMilliliters As Double
    Private LG As LoadingGraphics(Of PictureBox)
    Private varTappeId As Integer
    ' 0.55 = Plasma
    ' 0.45 = RB
    ' 0.01 = Trombocytter
    Private varReady As Boolean = True

    Private WithEvents DBC_HentBlodmengde, DBC_LagreRapport, DBC_LagreProdukter As New DatabaseClient(Credentials.Server, Credentials.Database, Credentials.UserID, Credentials.Password)
    Public Sub SelectNotification(Sender As StaffNotification, e As StaffNotificationEventArgs)
        If varReady Then
            Parent.Index = 3
            varReady = False
            LG.Spin(30, 10)
            varSelectedNotification = Sender
            IngentingValgtLab.Hide()
            X = 0
            DBC_HentBlodmengde.SQLQuery = "SELECT ant_liter, blodtype, tappe_id FROM Blodtapping WHERE time_id = @time_id;"
            DBC_HentBlodmengde.Execute({"@time_id"}, {DirectCast(Sender.RelatedElement, StaffTimeliste.StaffTime).TimeID.ToString})
        End If
    End Sub
    Private Sub DBC_Finished(Sender As Object, e As DatabaseListEventArgs) Handles DBC_HentBlodmengde.ListLoaded
        LG.StopSpin()

        If e.ErrorOccurred Then

        Else
            varTotalMilliliters = DirectCast(e.Data.Rows(0).Item(0), Double) * 1000
            Verdier(0) = (0.55 * varTotalMilliliters)
            Verdier(1) = (0.45 * varTotalMilliliters)
            Verdier(2) = (0.01 * varTotalMilliliters)
            varTappeId = DirectCast(e.Data.Rows(0).Item(2), Integer)
        End If
        varReady = True
        varEaseTimer.Start()
    End Sub
    Private Sub TimerTick() Handles varEaseTimer.Elapsed
        SC.Post(AddressOf TimerPost, Nothing)
    End Sub
    Private Sub TimerPost(State As Object)
        X += 1
        With RBColumn
            .Height = CInt(EaseInOut.GetY(0, Verdier(1), X, Xmax))
            .Location = New Point(Width \ 2 - .Width \ 2, Height - .Height - 90)
        End With
        With PlasmaColumn
            .Height = CInt(EaseInOut.GetY(0, Verdier(0), X, Xmax))
            .Location = New Point(RBColumn.Left - .Width - 10, Height - .Height - 90)
        End With
        With TromboColumn
            .Height = CInt(EaseInOut.GetY(0, Verdier(2), X, Xmax))
            .Location = New Point(RBColumn.Right + 10, Height - .Height - 90)
        End With
        If X < Xmax Then
            varEaseTimer.Start()
        End If
    End Sub
    Public Sub New(ParentWindow As MultiTabWindow)
        MyBase.New(ParentWindow)
        varEaseTimer.AutoReset = False
        With Header
            .Text = "Kontroller blodmengde"
        End With
        With Registrer
            .Size = New Size(130, 40)
            .Text = "Registrer"
            .BackColorNormal = Color.LimeGreen
            .BackColorSelected = ColorHelper.Multiply(.BackColorNormal, 0.7)
            .BackColorDisabled = Color.FromArgb(200, 200, 200)
            .ForeColor = Color.White
            .Left = 20
        End With
        With IngentingValgtLab
            .Hide()
            .ForeColor = Color.FromArgb(80, 80, 80)
            .Parent = Me
            .AutoSize = False
            .Size = New Size(100, 20)
            .TextAlign = ContentAlignment.MiddleCenter
            .Text = "Ingenting valgt"
        End With
        With LoadingSurface
            .Hide()
            .Parent = Me
            .Size = New Size(50, 50)
        End With

        With RBColumn
            .Size = New Size(80, 0)
            .BackColor = OffRed
            .Parent = Me
            .Location = New Point(Width \ 2 - .Width \ 2, Height - .Height - 90)
        End With
        With PlasmaColumn
            .Size = New Size(80, 0)
            .BackColor = OffGreen
            .Parent = Me
            .Show()
            .Location = New Point(RBColumn.Left - .Width - 10, Height - .Height - 90)
        End With
        With TromboColumn
            .Size = New Size(80, 0)
            .BackColor = OffBlue
            .Parent = Me
            .Location = New Point(RBColumn.Right + 10, Height - .Height - 90)
        End With



        NotifManager = New NotificationManager(Header)
        LG = New LoadingGraphics(Of PictureBox)(LoadingSurface)
        With LG
            .Stroke = 3
            .Pen.Color = Color.FromArgb(162, 25, 51)
        End With
    End Sub
    Private varProductStep As Integer = 0
    Private Sub RegistrerProdukter_Click() Handles Registrer.Click
        For Each C As Control In Me.Controls
            C.Hide()
        Next
        Header.Show()
        LG.Spin(30, 10)
        With DBC_LagreRapport
            .SQLQuery = "INSERT INTO Labrapport (tappe_id, plasma_mengde, blodplater, rode_blodlegemer) VALUES (@tappeid, @plasma, @blodplater, @rode);"
            .Execute({"@tappeid", "@plasma", "@blodplater", "@rode"}, {varTappeId.ToString, Verdier(0).ToString, Verdier(2).ToString, Verdier(1).ToString})
        End With

        With DBC_LagreProdukter
            .SQLQuery = "INSERT INTO Produkt (produkt_type, rapport_id, mengde) VALUES (@type, @rapportid, @mengde);"
            .Execute({"@type", "@rapportid", "@mengde"}, {"PL", "0", Verdier(0).ToString})
        End With
    End Sub
    Private Sub ProduktFerdig(Sender As Object, e As DatabaseListEventArgs) Handles DBC_LagreProdukter.ListLoaded
        Select Case varProductStep
            Case 0
                With DBC_LagreProdukter
                    .SQLQuery = "INSERT INTO Produkt (produkt_type, rapport_id, mengde) VALUES (@type, @rapportid, @mengde);"
                    .Execute({"@type", "@rapportid", "@mengde"}, {"RB", "0", Verdier(1).ToString})
                End With
                varProductStep += 1
            Case 1
                With DBC_LagreProdukter
                    .SQLQuery = "INSERT INTO Produkt (produkt_type, rapport_id, mengde) VALUES (@type, @rapportid, @mengde);"
                    .Execute({"@type", "@rapportid", "@mengde"}, {"TR", "0", Verdier(2).ToString})
                End With
                varProductStep += 1
            Case 2
                varProductStep += 1
                With DBC_LagreProdukter
                    .SQLQuery = "UPDATE Blodtapping SET analysert = 1 WHERE rapport_id = @rapportid;"
                    .Execute({"@type", "@rapportid", "@mengde"}, {"TR", "0", Verdier(2).ToString})
                End With
            Case Else
                LG.StopSpin()
                varProductStep = 0
                NotifManager.Display("Labrapporten er registrert, og produktene er lagt i systemet.", NotificationPreset.GreenSuccess,, FloatX.FillWidth, FloatY.FillHeight)
                varSelectedNotification.Close()
        End Select
    End Sub

    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        Header.Width = Width
        With Registrer
            .Top = Height - .Height - 20
            .Left = 20
        End With
        With IngentingValgtLab
            .Location = New Point((Width - .Width) \ 2, (Height + Header.Bottom - .Height) \ 2)
        End With
        With LoadingSurface
            .Location = New Point((Width - .Width) \ 2, (Height + Header.Bottom - .Height) \ 2)
        End With
        With RBColumn
            .Location = New Point(Width \ 2 - .Width \ 2, Height - .Height - 90)
        End With
        With PlasmaColumn
            .Location = New Point(RBColumn.Left - .Width - 10, Height - .Height - 90)
        End With
        With TromboColumn
            .Location = New Point(RBColumn.Right + 10, Height - .Height - 90)
        End With
    End Sub
End Class




Public Class Egenerklæringsliste
    Private Erklæringsliste As New List(Of Egenerklæring)
    Public Sub New()

    End Sub
    Public Sub Clear()
        Erklæringsliste.Clear()
    End Sub
    Public ReadOnly Property Count As Integer
        Get
            Return Erklæringsliste.Count
        End Get
    End Property
    Public Sub Add(ByRef Egenerklæring As Egenerklæring)
        Erklæringsliste.Add(Egenerklæring)
    End Sub
    Public ReadOnly Property TimeAtIndex(ByVal Index As Integer) As Egenerklæring
        Get
            Return Erklæringsliste(Index)
        End Get
    End Property
    Public Function GetAllElementsWhere(ByVal Egenskap As EgenerklæringEgenskap, ByVal Verdi As Object, Optional ByVal Condition As ComparisonOperator = ComparisonOperator.EqualTo, Optional ByVal ReturnIfConditionIs As Boolean = True) As List(Of Egenerklæring)
        Dim Ret As List(Of Egenerklæring)
        Try
            Ret = Erklæringsliste.FindAll(Function(Erklæring As Egenerklæring) As Boolean
                                              Select Case Egenskap
                                                  Case EgenerklæringEgenskap.TimeID
                                                      Return ((Erklæring.TimeID = DirectCast(Verdi, Integer)) = ReturnIfConditionIs)
                                                  Case EgenerklæringEgenskap.Godkjent
                                                      Return ((Erklæring.Godkjent = DirectCast(Verdi, Boolean)) = ReturnIfConditionIs)
                                                  Case EgenerklæringEgenskap.Land
                                                      Return ((Erklæring.Land = DirectCast(Verdi, String)) = ReturnIfConditionIs)
                                                  Case EgenerklæringEgenskap.Skjema
                                                      Select Case Condition
                                                          Case ComparisonOperator.EqualTo
                                                              Return ((Erklæring.SkjemaString = DirectCast(Verdi, String)) = ReturnIfConditionIs)
                                                          Case ComparisonOperator.IsLike
                                                              Return (Erklæring.SkjemaIsLike(DirectCast(Verdi, String)) = ReturnIfConditionIs)
                                                          Case Else
                                                              Return False
                                                      End Select
                                                  Case Else
                                                      Return (ReferenceEquals(Erklæring, Verdi) = ReturnIfConditionIs)
                                              End Select
                                          End Function)
        Catch ex As Exception

            Ret = Nothing
        End Try
        Return Ret
    End Function
    Public Function GetElementWhere(ByVal Egenskap As EgenerklæringEgenskap, ByVal Verdi As Object, Optional ByVal Condition As ComparisonOperator = ComparisonOperator.EqualTo, Optional ByVal ReturnIfConditionIs As Boolean = True) As Egenerklæring
        Dim Ret As Egenerklæring
        Try
            Ret = Erklæringsliste.Find(Function(Erklæring As Egenerklæring) As Boolean
                                           Select Case Egenskap
                                               Case EgenerklæringEgenskap.TimeID
                                                   Return ((Erklæring.TimeID = DirectCast(Verdi, Integer)) = ReturnIfConditionIs)
                                               Case EgenerklæringEgenskap.Godkjent
                                                   Return ((Erklæring.Godkjent = DirectCast(Verdi, Boolean)) = ReturnIfConditionIs)
                                               Case EgenerklæringEgenskap.Land
                                                   Return ((Erklæring.Land = DirectCast(Verdi, String)) = ReturnIfConditionIs)
                                               Case EgenerklæringEgenskap.Skjema
                                                   Select Case Condition
                                                       Case ComparisonOperator.EqualTo
                                                           Return ((Erklæring.SkjemaString = DirectCast(Verdi, String)) = ReturnIfConditionIs)
                                                       Case ComparisonOperator.IsLike
                                                           Return (Erklæring.SkjemaIsLike(DirectCast(Verdi, String)) = ReturnIfConditionIs)
                                                       Case Else
                                                           Return False
                                                   End Select
                                               Case Else
                                                   Return (ReferenceEquals(Erklæring, Verdi) = ReturnIfConditionIs)
                                           End Select
                                       End Function)
        Catch ex As Exception

            Ret = Nothing
        End Try
        Return Ret
    End Function
    Public Enum EgenerklæringEgenskap As Integer
        TimeID = 0
        Godkjent = 1
        Land = 2
        Skjema = 3
        AnsattSvar = 4
    End Enum
    Public Enum ComparisonOperator As Integer
        EqualTo = 0
        GreaterThan = 1
        LessThan = 2
        BooleanTrue = 3
        ReferencesEqual = 4
        IsLike = 5
    End Enum

    Public Class Egenerklæring
        Private varTimeID As Integer
        Private varLand, varSkjemaString, varAnsattSvar As String
        Private varGodkjent As Boolean

        ''' <summary>
        ''' X = Skip check
        ''' All other chars = Must match
        ''' </summary>
        ''' <param name="Pattern">A string consisting of 25 characters. Case sensitive.</param>
        ''' <returns>A boolean value indicating whether or not the string matches the pattern.</returns>
        Public Function SkjemaIsLike(Pattern As String) As Boolean
            Dim SkipChar As Char = "X".ToCharArray()(0)
            Dim SvarArr() As Char = varSkjemaString.ToCharArray()
            Dim PatternArr() As Char = Pattern.ToCharArray
            Dim IsCompatible As Boolean = True
            For i As Integer = 0 To 24
                If Not PatternArr(i) = SkipChar Then
                    If Not PatternArr(i) = SvarArr(i) Then
                        IsCompatible = False
                        Exit For
                    End If
                End If
            Next
            Return IsCompatible
        End Function
        Public Property AnsattSvar As String
            Get
                Return varAnsattSvar
            End Get
            Set(value As String)
                varAnsattSvar = value
            End Set
        End Property
        Public Property TimeID As Integer
            Get
                Return varTimeID
            End Get
            Set(value As Integer)
                varTimeID = value
            End Set
        End Property
        Public Property Land As String
            Get
                Return varLand
            End Get
            Set(value As String)
                varLand = Land
            End Set
        End Property
        Public Property SkjemaString As String
            Get
                Return varSkjemaString
            End Get
            Set(value As String)
                varSkjemaString = value
            End Set
        End Property
        Public Property Godkjent As Boolean
            Get
                Return varGodkjent
            End Get
            Set(value As Boolean)
                varGodkjent = value
            End Set
        End Property
        Public Sub New(TimeID As Integer, SkjemaString As String, Land As String, Godkjent As Boolean)
            varTimeID = TimeID
            varSkjemaString = SkjemaString
            varLand = Land
            varGodkjent = Godkjent
        End Sub
    End Class
End Class
Public Class StaffTimeliste
    Private Timeliste As New List(Of StaffTime)
    Public Sub New()

    End Sub
    Public Sub Clear()
        Timeliste.Clear()
    End Sub
    Public Sub RemoveAt(ByVal Index As Integer)
        Timeliste.RemoveAt(Index)
    End Sub
    Public Function GetIndexOf(Time As StaffTime) As Integer
        Dim MatchAt As Integer = -1
        Dim iLast As Integer = Timeliste.Count - 1
        For i As Integer = 0 To iLast
            If ReferenceEquals(Timeliste(i), Time) Then
                MatchAt = i
                Exit For
            End If
        Next
        Return MatchAt
    End Function
    Public ReadOnly Property Count As Integer
        Get
            Return Timeliste.Count
        End Get
    End Property
    Public ReadOnly Property Timer As List(Of StaffTime)
        Get
            Return Timeliste
        End Get
    End Property
    Public ReadOnly Property TimeAtIndex(ByVal Index As Integer) As StaffTime
        Get
            Return Timeliste(Index)
        End Get
    End Property
    Public Function GetAllElementsWhere(ByVal Egenskap As TimeEgenskap, ByVal Verdi As Object, Optional ByVal Condition As ComparisonOperator = ComparisonOperator.EqualTo, Optional ByVal ReturnIfConditionIs As Boolean = True) As List(Of StaffTime)
        Dim Ret As List(Of StaffTime)
        Try
            Ret = Timeliste.FindAll(Function(Time As StaffTime) As Boolean
                                        Select Case Egenskap
                                            Case TimeEgenskap.TimeID
                                                Select Case Condition
                                                    Case ComparisonOperator.EqualTo
                                                        Return ((Time.TimeID = DirectCast(Verdi, Integer)) = ReturnIfConditionIs)
                                                    Case Else
                                                        Return False
                                                End Select
                                            Case TimeEgenskap.Dato
                                                Select Case Condition
                                                    Case ComparisonOperator.EqualTo
                                                        Return ((Time.DatoOgTid.Date = DirectCast(Verdi, Date).Date) = ReturnIfConditionIs)
                                                    Case ComparisonOperator.LessThan
                                                        Return ((Time.DatoOgTid.Date.CompareTo(DirectCast(Verdi, Date).Date) < 0) = ReturnIfConditionIs)
                                                    Case ComparisonOperator.GreaterThan
                                                        Return ((Time.DatoOgTid.Date.CompareTo(DirectCast(Verdi, Date).Date) > 0) = ReturnIfConditionIs)
                                                    Case Else
                                                        Return False
                                                End Select
                                            Case TimeEgenskap.Tid
                                                Select Case Condition
                                                    Case ComparisonOperator.EqualTo
                                                        Return ((Time.DatoOgTid.TimeOfDay = DirectCast(Verdi, TimeSpan)) = ReturnIfConditionIs)
                                                    Case ComparisonOperator.GreaterThan
                                                        Return ((Time.DatoOgTid.TimeOfDay.CompareTo(DirectCast(Verdi, TimeSpan)) > 0) = ReturnIfConditionIs)
                                                    Case ComparisonOperator.LessThan
                                                        Return ((Time.DatoOgTid.TimeOfDay.CompareTo(DirectCast(Verdi, TimeSpan)) < 0) = ReturnIfConditionIs)
                                                    Case Else
                                                        Return False
                                                End Select
                                            Case TimeEgenskap.DatoOgTid
                                                Select Case Condition
                                                    Case ComparisonOperator.EqualTo
                                                        Return ((Time.DatoOgTid = DirectCast(Verdi, Date)) = ReturnIfConditionIs)
                                                    Case ComparisonOperator.LessThan
                                                        Return ((Time.DatoOgTid.CompareTo(DirectCast(Verdi, Date)) < 0) = ReturnIfConditionIs)
                                                    Case ComparisonOperator.GreaterThan
                                                        Return ((Time.DatoOgTid.CompareTo(DirectCast(Verdi, Date)) > 0) = ReturnIfConditionIs)
                                                    Case Else
                                                        Return False
                                                End Select
                                            Case TimeEgenskap.Fødselsnummer
                                                Return ((Time.Fødselsnummer = DirectCast(Verdi, String)) = ReturnIfConditionIs)
                                            Case TimeEgenskap.Godkjent
                                                Return ((Time.Godkjent = DirectCast(Verdi, Boolean)) = ReturnIfConditionIs)
                                            Case TimeEgenskap.AnsattID
                                                If IsDBNull(Verdi) Then
                                                    Return ((Time.AnsattID < 0) = ReturnIfConditionIs)
                                                Else
                                                    Return ((Time.AnsattID = DirectCast(Verdi, Integer)) = ReturnIfConditionIs)
                                                End If
                                            Case Else
                                                Return (ReferenceEquals(Time, Verdi) = ReturnIfConditionIs)
                                        End Select
                                    End Function)
        Catch ex As Exception

            Ret = Nothing
        End Try
        Return Ret
    End Function
    Public Function GetElementWhere(ByVal Egenskap As TimeEgenskap, ByVal Verdi As Object, Optional ByVal Condition As ComparisonOperator = ComparisonOperator.EqualTo, Optional ByVal ReturnIfConditionIs As Boolean = True) As StaffTime
        Dim Ret As StaffTime
        Try
            Ret = Timeliste.Find(Function(Time As StaffTime) As Boolean
                                     Select Case Egenskap
                                         Case TimeEgenskap.TimeID
                                             Select Case Condition
                                                 Case ComparisonOperator.EqualTo
                                                     Return ((Time.TimeID = DirectCast(Verdi, Integer)) = ReturnIfConditionIs)
                                                 Case Else
                                                     Return False
                                             End Select
                                         Case TimeEgenskap.Dato
                                             Select Case Condition
                                                 Case ComparisonOperator.EqualTo
                                                     Return ((Time.DatoOgTid.Date = DirectCast(Verdi, Date).Date) = ReturnIfConditionIs)
                                                 Case ComparisonOperator.LessThan
                                                     Return ((Time.DatoOgTid.Date.CompareTo(DirectCast(Verdi, Date).Date) < 0) = ReturnIfConditionIs)
                                                 Case ComparisonOperator.GreaterThan
                                                     Return ((Time.DatoOgTid.Date.CompareTo(DirectCast(Verdi, Date).Date) > 0) = ReturnIfConditionIs)
                                                 Case Else
                                                     Return False
                                             End Select
                                         Case TimeEgenskap.Tid
                                             Select Case Condition
                                                 Case ComparisonOperator.EqualTo
                                                     Return ((Time.DatoOgTid.TimeOfDay = DirectCast(Verdi, TimeSpan)) = ReturnIfConditionIs)
                                                 Case ComparisonOperator.GreaterThan
                                                     Return ((Time.DatoOgTid.TimeOfDay.CompareTo(DirectCast(Verdi, TimeSpan)) > 0) = ReturnIfConditionIs)
                                                 Case ComparisonOperator.LessThan
                                                     Return ((Time.DatoOgTid.TimeOfDay.CompareTo(DirectCast(Verdi, TimeSpan)) < 0) = ReturnIfConditionIs)
                                                 Case Else
                                                     Return False
                                             End Select
                                         Case TimeEgenskap.DatoOgTid
                                             Select Case Condition
                                                 Case ComparisonOperator.EqualTo
                                                     Return ((Time.DatoOgTid = DirectCast(Verdi, Date)) = ReturnIfConditionIs)
                                                 Case ComparisonOperator.LessThan
                                                     Return ((Time.DatoOgTid.CompareTo(DirectCast(Verdi, Date)) < 0) = ReturnIfConditionIs)
                                                 Case ComparisonOperator.GreaterThan
                                                     Return ((Time.DatoOgTid.CompareTo(DirectCast(Verdi, Date)) > 0) = ReturnIfConditionIs)
                                                 Case Else
                                                     Return False
                                             End Select
                                         Case TimeEgenskap.Fødselsnummer
                                             Return ((Time.Fødselsnummer = DirectCast(Verdi, String)) = ReturnIfConditionIs)
                                         Case TimeEgenskap.Godkjent
                                             Return ((Time.Godkjent = DirectCast(Verdi, Boolean)) = ReturnIfConditionIs)
                                         Case TimeEgenskap.AnsattID
                                             If IsDBNull(Verdi) Then
                                                 Return ((Time.AnsattID < 0) = ReturnIfConditionIs)
                                             Else
                                                 Return ((Time.AnsattID = DirectCast(Verdi, Integer)) = ReturnIfConditionIs)
                                             End If
                                         Case Else
                                             Return (ReferenceEquals(Time, Verdi) = ReturnIfConditionIs)
                                     End Select
                                 End Function)
        Catch ex As Exception

            Ret = Nothing
        End Try
        Return Ret
    End Function
    Public Enum TimeEgenskap As Integer
        TimeID = 0
        Dato = 1
        Tid = 2
        DatoOgTid = 3
        Fødselsnummer = 4
        Godkjent = 5
        AnsattID = 6
        Reference = 7
    End Enum
    Public Enum ComparisonOperator As Integer
        EqualTo = 0
        GreaterThan = 1
        LessThan = 2
        BooleanTrue = 3
        ReferencesEqual = 4
    End Enum
    Public Sub Add(ByRef Time As StaffTime)
        Timeliste.Add(Time)
    End Sub
    Public Class StaffTime
        Private varTimeID As Integer
        Private varDatoOgTid As Date
        Private varFødselsnummer As String
        Private varGodkjent As Boolean
        Private varAnsattID As Integer = -1
        Private varBlodgiverGodkjent As Boolean
        Private varFullført As Boolean
        Public Property Fullført As Boolean
            Get
                Return varFullført
            End Get
            Set(value As Boolean)
                varFullført = value
            End Set
        End Property
        Public Property BlodgiverGodkjent As Boolean
            Get
                Return varBlodgiverGodkjent
            End Get
            Set(value As Boolean)
                varBlodgiverGodkjent = value
            End Set
        End Property
        Public Property AnsattID As Integer
            Get
                Return varAnsattID
            End Get
            Set(value As Integer)
                varAnsattID = value
            End Set
        End Property
        Public Property TimeID As Integer
            Get
                Return varTimeID
            End Get
            Set(value As Integer)
                varTimeID = value
            End Set
        End Property
        Public Property DatoOgTid As Date
            Get
                Return varDatoOgTid
            End Get
            Set(value As Date)
                varDatoOgTid = value
            End Set
        End Property
        Public Property Godkjent As Boolean
            Get
                Return varGodkjent
            End Get
            Set(value As Boolean)
                varGodkjent = value
            End Set
        End Property
        Public Property Fødselsnummer As String
            Get
                Return varFødselsnummer
            End Get
            Set(value As String)
                varFødselsnummer = value
            End Set
        End Property
        Public Sub New(TimeID As Integer, DatoOgTid As Date, Godkjent As Boolean, Fødselsnummer As String, AnsattID As Object, BlodgiverGodkjent As Boolean)
            varBlodgiverGodkjent = BlodgiverGodkjent
            varTimeID = TimeID
            varDatoOgTid = DatoOgTid
            varGodkjent = Godkjent
            varFødselsnummer = Fødselsnummer
            If IsDBNull(AnsattID) Then
                varAnsattID = -1
            Else
                varAnsattID = DirectCast(AnsattID, Integer)
            End If
        End Sub
    End Class
End Class