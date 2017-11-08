Public Class Footer
    Inherits Label
    Private varParentControl As Control
    Private CopyrightLabel As New Label
    Public Shadows Property Parent As Control
        Get
            Return MyBase.Parent
        End Get
        Set(value As Control)
            If varParentControl IsNot Nothing Then
                RemoveHandler varParentControl.Resize, AddressOf OnParentResize
            End If
            varParentControl = value
            AddHandler varParentControl.Resize, AddressOf OnParentResize
            MyBase.Parent = value
        End Set
    End Property
    Public Sub New(ParentControl As Control)
        DoubleBuffered = True
        Parent = ParentControl
        BackColor = Color.FromArgb(54, 68, 78)
        Height = 40
        Location = New Point(0, ParentControl.ClientSize.Height - Height)
        ForeColor = Color.FromArgb(169, 180, 188)
        Padding = New Padding(20, 0, 0, 0)
        TextAlign = ContentAlignment.MiddleLeft
        Text = "HemoGlobe AS        Organisasjonsnummer: 915831508/917638209        Kundeservice (tlf): 80015367        Kundeservice (epost): kontakt@hemoglobe.no"
        With CopyrightLabel
            .Parent = Me
            .AutoSize = False
            .Size = New Size(400, Height - 2)
            .BackColor = BackColor
            .ForeColor = ForeColor
            .Location = New Point(Width - .Width, 1)
            .TextAlign = ContentAlignment.MiddleRight
            .Padding = New Padding(0, 0, 20, 0)
            .Text = "Copyright Magnus Bakke og Andreas Ore Larssen | 2017"
        End With
    End Sub
    Private Sub OnParentResize(Sender As Object, e As EventArgs)
        With varParentControl
            Location = New Point(0, .ClientSize.Height - Height)
            Width = .Width
        End With
        CopyrightLabel.Location = New Point(Width - CopyrightLabel.Width, 1)
    End Sub
End Class
