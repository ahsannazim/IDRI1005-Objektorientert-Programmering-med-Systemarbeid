Option Strict On
Option Explicit On
Option Infer Off

Public Class GlobalControl
    Private Shared OnClickAction As Action(Of LocalDeployment)
    Private Shared varDefaultBitmap, varHoverBitmap As Bitmap
    Private Shared DeploymentList As List(Of PictureBox)
    Private Shared varDefaultPosition As New Point(0, 0)
    Public Shared Event Click(Sender As LocalDeployment)
    Public Shared Event MouseEnter(Sender As LocalDeployment)
    Public Shared Event MouseLeave(Sender As LocalDeployment)
    Public Shared Property DefaultPosition As Point
        Get
            Return varDefaultPosition
        End Get
        Set(value As Point)
            varDefaultPosition = value
        End Set
    End Property
    Public Shared Property ClickAction As Action(Of LocalDeployment)
        Get
            Return OnClickAction
        End Get
        Set(Action As Action(Of LocalDeployment))
            OnClickAction = Action
        End Set
    End Property
    Public Shared Property DefaultBitmap As Bitmap
        Get
            Return varDefaultBitmap
        End Get
        Set(value As Bitmap)
            varDefaultBitmap = value
        End Set
    End Property
    Public Shared Property HoverBitmap As Bitmap
        Get
            Return varHoverBitmap
        End Get
        Set(value As Bitmap)
            varHoverBitmap = value
        End Set
    End Property
    Protected Shared Sub OnDeploymentMouseEnter(Sender As LocalDeployment)
        RaiseEvent MouseEnter(Sender)
    End Sub
    Protected Shared Sub OnDeploymentMouseLeave(Sender As LocalDeployment)
        RaiseEvent MouseLeave(Sender)
    End Sub
    Protected Shared Sub OnDeploymentClick(Sender As LocalDeployment)
        RaiseEvent Click(Sender)
    End Sub
    Public Sub Deploy(Parent As Control)
        Dim NewLocal As New LocalDeployment
        With NewLocal
            .BackgroundImageLayout = ImageLayout.Zoom
            .BackgroundImage = varDefaultBitmap
            .Location = varDefaultPosition
            .Size = varDefaultBitmap.Size
            .Parent = Parent
            .Show()
        End With
    End Sub
    Public Class LocalDeployment
        Inherits Control
        Protected Friend Sub New()
            With Me
                .Hide()
            End With
        End Sub
        Protected Overrides Sub OnMouseEnter(e As EventArgs)
            MyBase.OnMouseEnter(e)
            OnDeploymentMouseEnter(Me)
            BackgroundImage = varHoverBitmap
        End Sub
        Protected Overrides Sub OnMouseLeave(e As EventArgs)
            MyBase.OnMouseLeave(e)
            OnDeploymentMouseLeave(Me)
            BackgroundImage = varDefaultBitmap
        End Sub
        Protected Overrides Sub OnClick(e As EventArgs)
            MyBase.OnClick(e)
            OnDeploymentClick(Me)
        End Sub
    End Class
End Class