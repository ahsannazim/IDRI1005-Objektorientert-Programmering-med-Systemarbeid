Imports System.Text.RegularExpressions
Imports AudiopoLib

Module Globals
    Public CurrentLogin As UserInfo
    Public CurrentStaff As StaffInfo

    Public Windows As MultiTabWindow
    Public TimeListe As New StaffTimeliste

    ' Regex
    Public RegExEmail As New Regex("([\w-+]+(?:\.[\w-+]+)*@(?:[\w-]+\.)+[a-zA-Z]{2,7})")


    Public PersonaliaTest As Personopplysninger
    Public MainWindow As Main
    Public LoggInnTab As LoggInnNy
    Public Dashboard As DashboardTab
    Public Egenerklæring As EgenerklæringTab
    Public Timebestilling As TimebestillingTab
    Public AnsattLoggInn As AnsattLoggInnTab
    Public OpprettAnsatt As OpprettAnsattTab
    Public Credentials As DatabaseCredentials
    Public AnsattDashboard As AnsattDashboardTab
    Public RedigerProfil As RedigerProfilTab
    Public TimerHentet As Boolean

    Public OffGreen As Color = Color.FromArgb(94, 166, 21)
    Public OffBlue As Color = Color.FromArgb(47, 111, 149)
    Public OffRed As Color = Color.FromArgb(162, 23, 27)


    Public Sub Logout(Optional ByVal StaffLogout As Boolean = False)
        TimeListe.Clear()
        TimerHentet = False
        Windows.ResetAll()
        If Not StaffLogout Then
            Windows.Index = 1
        Else
            Windows.Index = 5
        End If
        ' TODO: Erase all traces of user data
    End Sub


    Public Sub CloseNotification(Sender As UserNotification, e As UserNotificationEventArgs)
        Sender.close
    End Sub
    Public Sub FormAnswered(Sender As UserNotification, e As UserNotificationEventArgs)
        MsgBox(DirectCast(Sender.RelatedElement, Egenerklæringsliste.Egenerklæring).AnsattSvar)
        Sender.Close()
    End Sub
    Public Sub FormNotSent(Sender As UserNotification, e As UserNotificationEventArgs)
        Egenerklæring.SelectTime(Sender, e)
    End Sub
End Module

'Public Class DatabaseTimeListe
'    Private varTimer As New List(Of StaffTimeliste.StaffTime)
'    Public Sub New()

'    End Sub
'    Public ReadOnly Property Count As Integer
'        Get
'            Return varTimer.Count
'        End Get
'    End Property
'    Public Sub Clear()
'        varTimer.Clear()
'    End Sub
'    Public Sub Add(ByRef Element As DatabaseTimeElement)
'        varTimer.Add(Element)
'    End Sub
'    Public ReadOnly Property Time(ByVal Index As Integer) As DatabaseTimeElement
'        Get
'            Return varTimer(Index)
'        End Get
'    End Property
'    Public ReadOnly Property TimeListe As List(Of StaffTimeliste.StaffTime)
'        Get
'            Return varTimer
'        End Get
'    End Property
'End Class

'Public Class DatabaseTimeElement
'    Private varTimeID As Integer
'    Private varDatoOgTid As Date
'    Public ReadOnly Property TimeID As Integer
'        Get
'            Return varTimeID
'        End Get
'    End Property
'    Public ReadOnly Property DatoOgTid As Date
'        Get
'            Return varDatoOgTid
'        End Get
'    End Property
'    Public Sub New(ByVal TimeID As Integer, ByVal DatoOgTid As Date)
'        varTimeID = TimeID
'        varDatoOgTid = DatoOgTid
'    End Sub
'End Class

Public Class StaffInfo
    Private varID As Integer
    Private varUsername, varFirstName, varLastName As String
    Public Property ID As Integer
        Get
            Return varID
        End Get
        Set(value As Integer)
            varID = value
        End Set
    End Property
    Public Property Username As String
        Get
            Return varUsername
        End Get
        Set(value As String)
            varUsername = value
        End Set
    End Property
    Public Property FirstName As String
        Get
            Return varFirstName
        End Get
        Set(value As String)
            varFirstName = value
        End Set
    End Property
    Public Property LastName As String
        Get
            Return varLastName
        End Get
        Set(value As String)
            varLastName = value
        End Set
    End Property
    Public ReadOnly Property FullName As String
        Get
            Return varFirstName & " " & varLastName
        End Get
    End Property
    Public Sub New(Username As String)
        varUsername = Username
        ID = -1
        varFirstName = "Undefined"
        varLastName = "Undefined"
    End Sub
    Public Sub EraseInfo()
        varID = Nothing
        varUsername = Nothing
        varFirstName = Nothing
        varLastName = Nothing
    End Sub
End Class

Public Class UserInfo
    Private Number, varFirstName, varLastName As String
    Private varFormInfo As Egenerklæringsliste.Egenerklæring
    Private varLatestAppointment As Date
    Private varRelatedDonor As Donor
    Private WithEvents DBC_GetDonorInfo As New DatabaseClient(Credentials.Server, Credentials.Database, Credentials.UserID, Credentials.Password)
    Private WithEvents DBC_GetLastAppointment As New DatabaseClient(Credentials.Server, Credentials.Database, Credentials.UserID, Credentials.Password)
    Public Property RelatedDonor As Donor
        Get
            Return varRelatedDonor
        End Get
        Set(value As Donor)
            varRelatedDonor = value
        End Set
    End Property
    Public Property LatestAppointment As Date
        Get
            Return varLatestAppointment
        End Get
        Set(value As Date)
            varLatestAppointment = value
        End Set
    End Property
    Private Sub DBC_GetLast_Finished(Sender As Object, e As DatabaseListEventArgs) Handles DBC_GetLastAppointment.ListLoaded
        If e.ErrorOccurred Then
            varLatestAppointment = Nothing
        ElseIf e.Data.Rows.Count > 0 Then
            varLatestAppointment = DirectCast(e.Data.Rows(0).Item(0), Date).Date
        End If
        DBC_GetLastAppointment.Dispose()
    End Sub
    Private Sub DBC_GetDonor_Finished(Sender As Object, e As DatabaseListEventArgs) Handles DBC_GetDonorInfo.ListLoaded
        If e.ErrorOccurred Then
            Logout()
        Else
            With e.Data.Rows(0)
                If IsDBNull(.Item(4)) Then
                    .Item(4) = "Ikke registrert"
                End If
                If IsDBNull(.Item(5)) Then
                    .Item(5) = "Ikke registrert"
                End If
                If IsDBNull(.Item(12)) Then
                    .Item(12) = "Ukjent"
                End If
                Dim Fødselsnummer As Long = DirectCast(.Item(0), Long)
                Dim Navn() As String = {DirectCast(.Item(1), String), DirectCast(.Item(2), String)}
                Dim Telefon() As String = {DirectCast(.Item(3), String), DirectCast(.Item(4), String), DirectCast(.Item(5), String)}
                Dim EpostOgAdresse() As String = {DirectCast(.Item(6), String), DirectCast(.Item(7), String)}
                Dim Postnummer As Integer = DirectCast(.Item(8), Integer)
                Dim KjønnEpostSMS() As Boolean = {DirectCast(.Item(9), Boolean), DirectCast(.Item(10), Boolean), DirectCast(.Item(11), Boolean)}
                Dim Blodtype As String = DirectCast(.Item(12), String)
                varRelatedDonor = New Donor(Fødselsnummer, Navn(0), Navn(1), Telefon, EpostOgAdresse(0), EpostOgAdresse(1), Blodtype, KjønnEpostSMS(1), KjønnEpostSMS(2), KjønnEpostSMS(0), Postnummer)
            End With
        End If
        DBC_GetDonorInfo.Dispose()
    End Sub
    Public Property FormInfo As Egenerklæringsliste.Egenerklæring
        Get
            Return varFormInfo
        End Get
        Set(value As Egenerklæringsliste.Egenerklæring)
            varFormInfo = value
        End Set
    End Property
    Public Sub EraseInfo()
        Number = Nothing
        varFirstName = Nothing
        varLastName = Nothing
        varFormInfo = Nothing
        varLatestAppointment = Nothing
        varRelatedDonor = Nothing
    End Sub
    Public ReadOnly Property FullName As String
        Get
            Return varFirstName & " " & varLastName
        End Get
    End Property
    Public Property FirstName As String
        Get
            Return varFirstName
        End Get
        Set(value As String)
            varFirstName = value
        End Set
    End Property
    Public Property LastName As String
        Get
            Return varLastName
        End Get
        Set(value As String)
            varLastName = value
        End Set
    End Property
    Public Sub New(PersonalNumber As String)
        Number = PersonalNumber
        varFirstName = "Undefined"
        varLastName = "Undefined"
        DBC_GetLastAppointment.SQLQuery = "SELECT t_dato FROM Time WHERE fullført = 1 AND b_fodselsnr = @nr ORDER BY t_dato DESC LIMIT 1;"
        DBC_GetLastAppointment.Execute({"@nr"}, {PersonalNumber})

        DBC_GetDonorInfo.SQLQuery = "SELECT * FROM Blodgiver WHERE b_fodselsnr = @id LIMIT 1;"
        DBC_GetDonorInfo.Execute({"@id"}, {PersonalNumber})
    End Sub
    Public Property PersonalNumber As String
        Get
            Return Number
        End Get
        Set(value As String)
            Number = value
        End Set
    End Property
    Public Function IsMale() As Boolean
        Dim Chars() As Char = Number.ToCharArray
        Return (Convert.ToInt32(Chars(8)) Mod 2 = 1)
    End Function
End Class
