Option Strict On
Option Explicit On
Option Infer Off

Imports System.IO
Imports System.Security.Cryptography
Imports AudiopoLib
Imports Microsoft.VisualBasic.FileIO

Public Class SetupTjener
    Dim LoginHelper As DatabaseSetup
    Dim LoadingGraphics As LoadingGraphics(Of PictureBox)
    Dim WithEvents LayoutTool As FormLayoutTools
    Dim WithEvents FWButton, FerdigKnapp, GodtaKnapp, OKOmrådeKnapp, TilbakeKnapp As FullWidthControl
    Dim NotifManager As NotificationManager
    Dim Heh As ResourceDeployer
    Dim Filbane As String = SpecialDirectories.MyDocuments & "\Blodbank\"
    Private Sub FerdigKnapp_Click() Handles FerdigKnapp.Click
        Me.Hide()
        If CheckKjør.Checked Then

        End If
        FreeResources()
        End
    End Sub
    Private Sub GodtaKnapp_Click() Handles GodtaKnapp.Click
        GroupAvtale.Hide()
        LayoutTool.CenterOnForm(GroupFilbane)
        GroupFilbane.Show()
        TilbakeKnapp.Show()
    End Sub
    Private Sub OKOmrådeKnapp_Click() Handles OKOmrådeKnapp.Click
        Dim ErrorOccurred As Boolean = False
        Try
            My.Computer.FileSystem.CreateDirectory(txtFilbane.Text)
        Catch
            ErrorOccurred = True
            txtFilbane.Clear()
            txtFilbane.Focus()
            NotifManager.Display("Ugyldig område", NotificationPreset.RedAlert)
        End Try
        If Not ErrorOccurred Then
            Filbane = txtFilbane.Text
            GroupFilbane.Hide()
            LayoutTool.CenterOnForm(GroupLoggInn)
            LayoutTool.CenterSurface(PicLoadingSurface, GroupLoggInn)
            GroupLoggInn.Show()
        End If
    End Sub
    Private Sub FreeResources()
        LoginHelper.Dispose()
        LoadingGraphics.Dispose()
        LayoutTool.Dispose()
        FWButton.Dispose()
        FerdigKnapp.Dispose()
        NotifManager.Dispose()
    End Sub
    Private Sub WriteToDisk()
        Heh = New ResourceDeployer
        Heh.WhenFinished = AddressOf FinishUp
        Heh.AddResourceName("kek", "txt")
        Heh.AddResourceName("Auditory", "mp3")
        Heh.DeployAll(Filbane & "\Blodbank\")
    End Sub
    Private Sub FinishUp(ByVal ErrorOccurred As Boolean)
        If Not ErrorOccurred Then
            LoadingGraphics.StopSpin()
            NotifManager.Display("The service is ready to be used.", New NotificationAppearance(Color.LimeGreen, Color.White,,,,, 0))
            GroupLoggInn.Hide()
            GroupFerdig.Show()
            TilbakeKnapp.Hide()
            LayoutTool.Refresh()
        Else
            MsgBox("Error occurred in FinishUp")
        End If
    End Sub
    Private Sub ActionFinished(Valid As Boolean, ErrorOccurred As Boolean)
        If Valid Then
            If Not ErrorOccurred Then
                WriteToDisk()
            Else
                LoadingGraphics.StopSpin()
                FWButton.Enabled = True
                For Each C As Control In GroupLoggInn.Controls
                    C.Show()
                Next
                NotifManager.Display("An error occurred. Please try again later.", NotificationPreset.RedAlert)
            End If
        Else
            LoadingGraphics.StopSpin()
            FWButton.Enabled = True
            For Each C As Control In GroupLoggInn.Controls
                C.Show()
            Next
            If Not ErrorOccurred Then
                NotifManager.Display("Please check your credentials.", NotificationPreset.RedAlert)
            Else
                NotifManager.Display("An unexpected error has occurred.", NotificationPreset.RedAlert)
            End If
        End If
    End Sub

    Private Sub cmdBlaGjennom_Click(sender As Object, e As EventArgs) Handles cmdBlaGjennom.Click
        Dim FolderBrower As New FolderBrowserDialog
        FolderBrower.RootFolder = Environment.SpecialFolder.MyComputer
        If FolderBrower.ShowDialog() = DialogResult.OK Then
            txtFilbane.Text = FolderBrower.SelectedPath
        End If
        FolderBrower.Dispose()
    End Sub

    Private Sub Setup_Tjener_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Hide()
        txtFilbane.Text = Filbane
        LoadingGraphics = New LoadingGraphics(Of PictureBox)(PicLoadingSurface)
        LayoutTool = New FormLayoutTools(Me)
        LoginHelper = New DatabaseSetup(AddressOf ActionFinished)
        NotifManager = New NotificationManager(Me)
        NotifManager.AssignedLayoutManager = LayoutTool
        With LoadingGraphics
            .Stroke = 2
            .Pen.Color = Color.LimeGreen
        End With
        With LayoutTool
            .IncludeFormTitle = True
            .CenterOnForm(GroupAvtale)
        End With
        FWButton = New FullWidthControl(GroupLoggInn, True, FullWidthControl.SnapType.Bottom)
        FerdigKnapp = New FullWidthControl(GroupFerdig, True, FullWidthControl.SnapType.Bottom)
        OKOmrådeKnapp = New FullWidthControl(GroupFilbane, True, FullWidthControl.SnapType.Bottom)
        GodtaKnapp = New FullWidthControl(GroupAvtale, True, FullWidthControl.SnapType.Bottom)
        TilbakeKnapp = New FullWidthControl(PanelBack, True)
        TilbakeKnapp.Hide()
        FWButton.Text = "Sjekk"
        FerdigKnapp.Text = "Ferdig"
        OKOmrådeKnapp.Text = "Videre"
        GodtaKnapp.Text = "Jeg/vi godtar vilkårene"
        Dim GroupHeaderInformasjon As New FullWidthControl(GroupLoggInn, False, FullWidthControl.SnapType.Top)
        Dim GroupHeaderAvtale As New FullWidthControl(GroupAvtale, False, FullWidthControl.SnapType.Top)
        Dim GroupHeaderFilbane As New FullWidthControl(GroupFilbane, False, FullWidthControl.SnapType.Top)
        Dim GroupHeaderFerdig As New FullWidthControl(GroupFerdig, False, FullWidthControl.SnapType.Top)
        With GroupHeaderInformasjon
            .Height = 20
            .BackColor = Color.FromArgb(230, 230, 230)
            .ForeColor = Color.FromArgb(100, 100, 100)
            .TextAlign = ContentAlignment.MiddleLeft
            .Padding = New Padding(5, 0, 0, 0)
            .Text = "Installasjon"
        End With
        With GroupHeaderAvtale
            .Height = 20
            .BackColor = Color.FromArgb(230, 230, 230)
            .ForeColor = Color.FromArgb(100, 100, 100)
            .TextAlign = ContentAlignment.MiddleLeft
            .Padding = New Padding(5, 0, 0, 0)
            .Text = "Vilkår for bruk"
        End With
        With GroupHeaderFilbane
            .Height = 20
            .BackColor = Color.FromArgb(230, 230, 230)
            .ForeColor = Color.FromArgb(100, 100, 100)
            .TextAlign = ContentAlignment.MiddleLeft
            .Padding = New Padding(5, 0, 0, 0)
            .Text = "Installasjon"
        End With
        With GroupHeaderFerdig
            .Height = 20
            .BackColor = Color.FromArgb(230, 230, 230)
            .ForeColor = Color.FromArgb(100, 100, 100)
            .TextAlign = ContentAlignment.MiddleLeft
            .Padding = New Padding(5, 0, 0, 0)
            .Text = "Installasjon"
        End With
        LayoutTool.CenterSurfaceH(checkAutoTabeller, GroupLoggInn)
        LayoutTool.CenterSurfaceH(CheckKjør, GroupFerdig)
        GroupFerdig.Hide()
        GroupFilbane.Hide()
        GroupLoggInn.Hide()
        Me.Show()
    End Sub
    Private Sub FWButton_Click(sender As Object, e As EventArgs) Handles FWButton.Click
        If txtTjener.Text = "" OrElse txtDatabase.Text = "" OrElse txtBrukernavn.Text = "" OrElse txtPassord.Text = "" OrElse txtKey.Text = "" OrElse txtRepeatKey.Text = "" Then
            NotifManager.Display("Alle felt må fylles ut", NotificationPreset.RedAlert)
        Else
            If txtKey.Text = txtRepeatKey.Text Then
                For Each C As Control In GroupLoggInn.Controls
                    If Not C.GetType = GetType(FullWidthControl) Then
                        C.Hide()
                    End If
                Next
                FWButton.Hide()
                FWButton.Enabled = False
                LayoutTool.Refresh()
                LayoutTool.CenterSurface(PicLoadingSurface, GroupLoggInn)
                LoadingGraphics.Spin(30, 10)
                LoginHelper.Path = Filbane
                LoginHelper.CheckCredentials(txtTjener.Text, txtDatabase.Text, txtBrukernavn.Text, txtPassord.Text, txtKey.Text)
                txtBrukernavn.Focus()
            Else
                txtKey.Clear()
                txtRepeatKey.Clear()
                txtKey.Focus()
                NotifManager.Display("Nøkkelfeltene er forskjellige", NotificationPreset.RedAlert)
            End If
        End If
    End Sub
    Private Sub LayoutTool_Refreshed() Handles LayoutTool.Refreshed, Me.Resize
        If LayoutTool IsNot Nothing Then
            If GroupAvtale.Visible Then
                LayoutTool.CenterOnForm(GroupAvtale)
            ElseIf GroupFilbane.Visible Then
                LayoutTool.CenterOnForm(GroupFilbane)
                If PicLoadingSurface.Visible = True Then
                    LayoutTool.CenterSurface(PicLoadingSurface, GroupFilbane, 0, 5)
                End If
            ElseIf GroupLoggInn.Visible Then
                LayoutTool.CenterOnForm(GroupLoggInn)
                If PicLoadingSurface.Visible = True Then
                    LayoutTool.CenterSurface(PicLoadingSurface, GroupLoggInn, 0, 5)
                End If
            ElseIf GroupFerdig.Visible Then
                LayoutTool.CenterOnForm(GroupFerdig)
            End If
            If PanelBack.Visible Then
                LayoutTool.CenterSurfaceH(PanelBack, Me)
            End If
        End If
    End Sub
    Private Sub GoBack() Handles TilbakeKnapp.Click
        MsgBox("GOING BACK")
        If GroupFilbane.Visible = True Then
            GroupFilbane.Hide()
            LayoutTool.CenterOnForm(GroupAvtale)
            TilbakeKnapp.Hide()
            GroupAvtale.Show()
        ElseIf GroupLoggInn.Visible = True Then
            GroupLoggInn.Hide()
            LayoutTool.CenterOnForm(GroupFilbane)
            GroupFilbane.Show()
            TilbakeKnapp.Show()
        End If
    End Sub
End Class

Public Class ResourceDeployer
    Private ResList, ExtList As List(Of String)
    Private TS As ThreadStarter
    Private IgnoreDup As Boolean = False
    Private ThreadsStarted, ThreadFinished As UInteger
    Private WhenFinishedAction As Action(Of Boolean)
    ''' <summary>
    ''' Gets or sets the delegate to be invoked when all the files have been to disk. This method must accept a boolean value that indicates whether or not the write was successfully completed.
    ''' Example: AddressOf mySub
    ''' </summary>
    Public Property WhenFinished As Action(Of Boolean)
        Get
            Return WhenFinishedAction
        End Get
        Set(Action As Action(Of Boolean))
            WhenFinishedAction = Action
        End Set
    End Property
    Public Property IgnoreDuplicateEntries As Boolean
        Get
            Return IgnoreDup
        End Get
        Set(value As Boolean)
            IgnoreDup = value
        End Set
    End Property
    Public Sub New(Optional ByVal Capacity As Integer = 20)
        ResList = New List(Of String)(Capacity)
        ExtList = New List(Of String)(Capacity)
    End Sub
    Private Function GetIndex(ByVal Name As String) As Integer
        For i As Integer = 0 To ResList.Count - 1
            If ResList(i) = Name Then
                Return i
            End If
        Next
        Return -1
    End Function
    Public Overloads Sub AddResourceName(ByVal Name As String, ByVal Extension As String)
        If IgnoreDup = False OrElse ResList.Count = 0 OrElse GetIndex(Name) = -1 Then
            ResList.Add(Name)
            ExtList.Add(Extension)
        End If
    End Sub
    Public Overloads Sub AddResourceName(ByVal Names() As String, ByVal Extensions() As String)
        If Names.Length = Extensions.Length Then
            For i As Integer = 0 To Names.GetLength(0) - 1
                If IgnoreDup = False OrElse ResList.Count = 0 OrElse GetIndex(Names(i)) = -1 Then
                    ResList.Add(Names(i))
                    ExtList.Add(Extensions(i))
                End If
            Next
        Else
            Throw New Exception("This overload of AddResourceName requires two one-dimensional string arrays of equal length.")
        End If
    End Sub
    Public Overloads Sub RemoveResourceName(ByVal Name As String)
        Dim i As Integer = GetIndex(Name)
        If i >= 0 Then
            ResList.RemoveAt(i)
            ExtList.RemoveAt(i)
        End If
    End Sub
    Public Overloads Sub RemoveResourceName(ByVal Names() As String)
        For i As Integer = 0 To Names.GetLength(0) - 1
            Dim n As Integer = GetIndex(Names(i))
            If n >= 0 Then
                ResList.RemoveAt(n)
                ExtList.RemoveAt(n)
            End If
        Next
    End Sub
    Public Sub DeployAll(ByVal Path As String, Optional ByVal Extension As String = "txt")
        Dim Data As Object() = {ResList, ExtList, Path}
        TS = New ThreadStarter(AddressOf DoWork)
        TS.WhenFinished = AddressOf WhenFinishedA
        TS.Start(Data)
    End Sub
    Private Function DoWork(ByVal RList As Object) As Object
        Dim Data As Object() = DirectCast(RList, Object())
        Dim RL As List(Of String) = DirectCast(Data(0), List(Of String))
        Dim XL As List(Of String) = DirectCast(Data(1), List(Of String))
        Dim Path As String = DirectCast(Data(2), String)
        Debug.Print("COUNT: " & RL.Count)
        Debug.Print("PATH: " & Path)
        Dim ErrorOccurred As Boolean = False
        Try
            For i As Integer = 0 To RL.Count - 1
                Dim ResourceContents As Object = My.Resources.ResourceManager.GetObject(RL(i))
                If ResourceContents.GetType = GetType(String) Then
                    Debug.Print("WRITING TRUE")
                    My.Computer.FileSystem.WriteAllBytes(Path & "\" & RL(i) & "." & XL(i), System.Text.Encoding.Unicode.GetBytes(CType(ResourceContents, String)), False)
                    ResourceContents = ""
                Else
                    Debug.Print("WRITING FALSE")
                    My.Computer.FileSystem.WriteAllBytes(Path & "\" & RL(i) & "." & XL(i), CType(ResourceContents, Byte()), False)
                    ResourceContents = Nothing
                End If
            Next
        Catch
            Debug.Print("CATCH")
            ErrorOccurred = True
        Finally
            RL.Clear()
        End Try
        Return ErrorOccurred
    End Function
    Private Sub WhenFinishedA(ErrorOccurred As Object, e As ThreadStarterEventArgs)
        ResList.Clear()
        TS.Dispose()
        WhenFinishedAction.Invoke(DirectCast(ErrorOccurred, Boolean))
    End Sub
End Class

Public NotInheritable Class DatabaseSetup
    Implements IDisposable
    Private Login As MySqlAdminLogin
    Private TS As ThreadStarter
    Private ServerString, DatabaseString, UIDString, PasswordString, KeyString As String
    Private ActionWhenFinished As ValidityAction = Nothing
    Private SavePath As String = "Default"
    Private AppName As String = "ApplicationName"
    Public Delegate Sub ValidityAction(ByVal Valid As Boolean, ByVal ErrorOccurred As Boolean)
    Public Property ApplicationName As String
        Get
            Return AppName
        End Get
        Set(value As String)
            AppName = value
        End Set
    End Property
    Public Property Path As String
        Get
            Return SavePath
        End Get
        Set(value As String)
            SavePath = value
        End Set
    End Property
    Public Property WhenFinishedAction As ValidityAction
        Get
            Return ActionWhenFinished
        End Get
        Set(Action As ValidityAction)
            ActionWhenFinished = Action
        End Set
    End Property
    Public Sub CheckCredentials(ByVal Server As String, ByVal Database As String, ByVal UID As String, ByVal Password As String, ByVal Key As String)
        ServerString = Server
        DatabaseString = Database
        UIDString = UID
        PasswordString = Password
        KeyString = Key
        Login = New MySqlAdminLogin(Server, Database, AddressOf WhenFinished)
        Login.AutoDispose = True
        ' Denne fryser GUI i over 0.3 sekunder. Årsaken må finnes (enten i DatabaseClient.New eller DatabaseClient.ValidateConnection 
        Login.LoginAsync(UID, Password)
    End Sub
    Public Sub New(ActionIfValid As ValidityAction)
        WhenFinishedAction = ActionIfValid
    End Sub
    Private Sub CheckValidity(Args As Object)
        Dim ArgsArr() As String = DirectCast(Args, String())
    End Sub
    Private Sub WhenFinished(Valid As Boolean)
        If Valid Then
            TS = New ThreadStarter(AddressOf SaveCredentials)
            TS.WhenFinished = AddressOf WhenFinishedSave
            Dim Args As Object = {ServerString, DatabaseString, UIDString, PasswordString, KeyString, SavePath}
            TS.Start(Args)
        Else
            ActionWhenFinished.Invoke(False, False)
        End If
    End Sub
    Private Sub WhenFinishedSave(Success As Object, e As ThreadStarterEventArgs)
        'If IfValidAction IsNot Nothing Then
        ActionWhenFinished.Invoke(True, DirectCast(Success, Boolean))
        'End If
    End Sub
    Private Function SaveCredentials(Info As Object) As Boolean
        Dim InfoArr() As String = DirectCast(Info, String())
        Dim CompleteString As String = InfoArr(0) & "%" & InfoArr(1) & "%" & InfoArr(2) & "%" & InfoArr(3)
        Dim Key As String = InfoArr(4)
        Dim Path As String = InfoArr(5)
        Dim Ret As Boolean = False
        Dim wrapper As New EncryptedReadWrite(Key)
        Try
            Dim cipherText As String = wrapper.EncryptData(CompleteString)
            If Path = "Default" OrElse Path = "" Then
                Path = SpecialDirectories.MyDocuments
            End If
            Path &= "\Blodbank"
            My.Computer.FileSystem.CreateDirectory(Path)
            My.Computer.FileSystem.WriteAllText(Path & "\auth.txt", cipherText, False)
            wrapper.Dispose()
            Ret = False
            MsgBox(Path)
        Catch
            Ret = True
        Finally
            wrapper.Dispose()
        End Try
        Return Ret
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean
    Protected Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                TS.Dispose()
                Login.Dispose()
            End If
        End If
        disposedValue = True
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
    End Sub
#End Region
End Class
Public NotInheritable Class EncryptedReadWrite
    Implements IDisposable
    Private TripleDes As New TripleDESCryptoServiceProvider
    Private Function TruncateHash(ByVal key As String, ByVal length As Integer) As Byte()
        Dim sha1 As New SHA1CryptoServiceProvider
        Dim ret() As Byte
        Try
            Dim keyBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(key)
            'Dim hashList As New List(Of Byte)(sha1.ComputeHash(keyBytes))
            ret = sha1.ComputeHash(keyBytes)
            ' Truncate or pad the hash. ?????
            ReDim Preserve ret(length - 1)
        Catch
            ret = Nothing
        Finally
            sha1.Dispose()
        End Try
        Return ret
    End Function
    Sub New(ByVal key As String)
        TripleDes.Key = TruncateHash(key, TripleDes.KeySize \ 8)
        TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
    End Sub
    Public Function EncryptData(ByVal plaintext As String) As String
        ' Convert the plaintext string to a byte array.
        Dim plaintextBytes() As Byte = Text.Encoding.Unicode.GetBytes(plaintext)

        ' Create the stream.
        Dim ms As New IO.MemoryStream
        ' Create the encoder to write to the stream.
        Dim encStream As New CryptoStream(ms, TripleDes.CreateEncryptor(), CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
        encStream.FlushFinalBlock()

        ' Convert the encrypted stream to a printable string.
        Return Convert.ToBase64String(ms.ToArray)
    End Function
    Public Function DecryptData(ByVal encryptedtext As String) As String
        ' Create the stream.
        Dim ms As New IO.MemoryStream
        Dim decStream As New CryptoStream(ms, TripleDes.CreateDecryptor(), CryptoStreamMode.Write)
        Try
            ' Convert the encrypted text string to a byte array.
            Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedtext)
            ' Create the decoder to write to the stream.
            ' Use the crypto stream to write the byte array to the stream.
            decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
            decStream.FlushFinalBlock()
            ' Convert the plaintext stream to a string.
            Return Text.Encoding.Unicode.GetString(ms.ToArray)
        Catch
            Return "Feil"
        Finally
            ms.Dispose()
            decStream.Dispose()
        End Try
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                TripleDes.Dispose()
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class

' WIP; brukes ikke i dette prosjektet (brukes når passord skal leses)
Public Class CredentialsManager
    Private DefPath As String
    Private TS As ThreadStarter
    Private AppName As String = "ApplicationName"
    Public Property ApplicationName As String
        Get
            Return AppName
        End Get
        Set(value As String)
            AppName = value
        End Set
    End Property
    Public Sub New(DefaultPath As String)
        If Not DefaultPath = "Default" Then
            DefPath = DefaultPath
        Else
            DefPath = SpecialDirectories.MyDocuments & "\" & AppName
        End If
        TS = New ThreadStarter(AddressOf NewAsync)
        TS.WhenFinished = AddressOf NewFinished
        TS.Start(DefPath)
    End Sub
    Private Sub NewFinished(ErrorOccurred As Object, e As ThreadStarterEventArgs)
        Dim Result As Boolean = DirectCast(ErrorOccurred, Boolean)
        If Result = True Then
            Throw New Exception("The specified path (" & DefPath & ") is inaccessible.")
        End If
        TS.Dispose()
    End Sub
    Private Function NewAsync(P As Object) As Boolean
        Dim Ret As Boolean = False
        Dim Path As String = DirectCast(P, String)
        Try
            System.IO.Directory.CreateDirectory(Path)
            If Not File.Exists(Path & "\auth.txt") Then
                File.Create(Path & "\auth.txt")
            End If
        Catch ex As Exception
            Ret = True
        End Try
        Return Ret
    End Function
    Public Sub Decode(ByVal Password As String)
        Dim cipherText As String = My.Computer.FileSystem.ReadAllText(DefPath & "\test.txt")
        Dim wrapper As New EncryptedReadWrite(Password)
        Try
            Dim plainText As String = wrapper.DecryptData(cipherText)
        Catch ex As System.Security.Cryptography.CryptographicException
            MsgBox("The data could not be decrypted with the password.")
        Finally
            Try
                wrapper.Dispose()
            Catch
            End Try
        End Try
    End Sub
End Class