Option Strict On
Option Explicit On
Option Infer Off

Imports AudiopoLib
Public Class AdminUnlockWrapper
    Inherits ContainerControl
    Private InnerContainer As New Control
    Private HeaderControl As FullWidthControl
    Private Lab As New Label
    Private TB As New TextBox
    Private LoggInn As FullWidthControl
    Private TSL As ThreadStarterLight
    Private LG As LoadingGraphics(Of PictureBox)
    Private LoadingSurface As New PictureBox
    Private ValidityChecker As MySqlAdminLogin
    Private CPath, FName As String
    Public Event IncorrectKey(Sender As Object, e As EventArgs)
    Public Event CorrectKey(Sender As Object, ConnectionSucceeded As Boolean)
    Public Sub New(CredPath As String, FileName As String)
        Hide()
        DoubleBuffered = True
        Size = New Size(250, 120)
        BackColor = Color.Gray
        With InnerContainer
            .Parent = Me
            .Width = Width - 2
            .Height = Height - 2
            .Location = New Point(1, 1)
            .BackColor = Color.White
            .TabStop = False
        End With
        CPath = CredPath
        FName = FileName
        HeaderControl = New FullWidthControl(InnerContainer)
        With HeaderControl
            .Text = "Lås opp (administrator)"
            .Height = 20
            .TextAlign = ContentAlignment.MiddleLeft
            .BackColor = Color.FromArgb(80, 80, 80)
            .TabStop = False
        End With
        With Lab
            .Parent = InnerContainer
            .Top = HeaderControl.Bottom + 8
            .Width = Width - 20
            .Left = 10
            .Text = "Skriv inn nøkkel for å låse opp systemet"
            .TabStop = False
        End With
        With TB
            .TabIndex = 0
            .Parent = InnerContainer
            .Width = Width - 20
            .Left = 10
            .Top = Lab.Bottom
            .UseSystemPasswordChar = True
            .Focus()
        End With
        Height = TB.Bottom + 40
        InnerContainer.Height = Height - 2
        LoggInn = New FullWidthControl(InnerContainer, True, FullWidthControl.SnapType.Bottom)
        With LoggInn
            .Text = "Lås opp"
            .TabIndex = 1
        End With
        AddHandler LoggInn.Click, AddressOf LoggInn_Click
        With LoadingSurface
            .TabStop = False
            .Hide()
            .Size = New Size(40, 40)
            .Location = New Point(Width \ 2 - .Width \ 2, Height \ 2 - .Height \ 2)
            .SendToBack()
            .Parent = InnerContainer
        End With
        LG = New LoadingGraphics(Of PictureBox)(LoadingSurface)
        With LG
            .Pen.Color = Color.FromArgb(230, 50, 80)
            .Stroke = 3
        End With
        Show()
    End Sub
    Private Sub LoggInn_Click(Sender As Object, e As EventArgs)
        SuspendLayout()
        LoggInn.Enabled = False
        If TSL IsNot Nothing Then
            TSL.Dispose()
        End If
        TSL = New ThreadStarterLight(AddressOf CredManager_Decode)
        TSL.WhenFinished = AddressOf CredManager_Finished
        TSL.Start(New String() {CPath, FName, TB.Text})
        For Each C As Control In InnerContainer.Controls
            C.Hide()
        Next
        LG.Spin(30, 10)
        ResumeLayout()
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
            Credentials = New DatabaseCredentials(Data(0), Data(1), Data(2), Data(3))
            If ValidityChecker IsNot Nothing Then
                ValidityChecker.Dispose()
            End If
            ValidityChecker = New MySqlAdminLogin(Data(0), Data(1))
            ValidityChecker.WhenFinished = AddressOf OnCheckFinished
            ValidityChecker.LoginAsync(Data(2), Data(3))
        Else
            Credentials = Nothing
            LoggInn.Enabled = True
            LG.StopSpin()
            LoadingSurface.SendToBack()
            TB.Clear()
            For Each C As Control In InnerContainer.Controls
                C.Show()
            Next
            TB.Focus()
            RaiseEvent IncorrectKey(Me, EventArgs.Empty)
        End If
    End Sub
    Private Sub OnCheckFinished(Valid As Boolean)
        LoggInn.Enabled = True
        LG.StopSpin()
        LoadingSurface.SendToBack()
        TB.Clear()
        If Not Valid Then
            For Each C As Control In InnerContainer.Controls
                C.Show()
            Next
        End If
        'If Valid Then
        '    TB.Hide()
        'End If
        RaiseEvent CorrectKey(Me, Valid)
    End Sub
    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            'HeaderControl.Dispose()
            'Lab.Dispose()
            'TB.Dispose()
            'LoggInn.Dispose()
            If TSL IsNot Nothing Then
                TSL.Dispose()
            End If
            LG.Dispose()
            'LoadingSurface.Dispose()
            If ValidityChecker IsNot Nothing Then
                ValidityChecker.Dispose()
            End If
            'InnerContainer.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class