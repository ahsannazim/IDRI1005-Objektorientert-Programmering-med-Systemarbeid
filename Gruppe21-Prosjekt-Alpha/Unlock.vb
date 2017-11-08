Imports AudiopoLib

Public Class Unlock
    Private WithEvents AdminUnlock As New AdminUnlockWrapper(Application.StartupPath & "\cred\", "auth.txt")
    Private NotifManager As New NotificationManager(Me)
    Private LayoutHelper As New FormLayoutTools(Me)
    Private Sub Unlock_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With AdminUnlock
            .Parent = Me
        End With
        LayoutHelper.CenterOnForm(AdminUnlock)
    End Sub
    Private Sub IncorrentKey(Sender As Object, e As EventArgs) Handles AdminUnlock.IncorrectKey
        NotifManager.Display("Nøkkelen var feil.", NotificationPreset.RedAlert)
    End Sub
    Private Sub CorrectKey(Sender As Object, ConnectionSucceeded As Boolean) Handles AdminUnlock.CorrectKey
        If ConnectionSucceeded Then
            AdminUnlock.Hide()
            Hide()
            Splash.Show()
        Else
            NotifManager.Display("Nøkkelen er riktig, men databasen kan ikke nås. Er du koblet til internett?", NotificationPreset.RedAlert)
        End If
    End Sub
    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        SuspendLayout()
        LayoutHelper.CenterOnForm(AdminUnlock)
        ResumeLayout(True)
    End Sub
End Class