Public Class HemoGlobeLogo
    Inherits Control
    Public Sub New()
        BackgroundImage = My.Resources.NyLogo
        Size = BackgroundImage.Size
        Location = New Point(0, 10)
        BackgroundImageLayout = ImageLayout.Zoom
    End Sub
End Class