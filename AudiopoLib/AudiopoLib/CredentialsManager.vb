Option Strict On
Option Explicit On
Option Infer Off

Imports System.IO
Imports System.Windows.Forms

Public Class CredentialsManager
    Private DefPath As String
    Private TSEncode, TSDecode As ThreadStarterLight ' Replace with light version
    Public Event DecodeFinished(ByVal ClearText As String, ByVal Valid As Boolean)
    Public Event EncodeFinished()
    Public Property DefaultPath As String
        Get
            Return DefPath
        End Get
        Set(value As String)
            DefPath = value
        End Set
    End Property
    Public Sub New(Optional DefaultPath As String = "Default")
        TSEncode = New ThreadStarterLight(AddressOf AsyncEncode)
        TSDecode = New ThreadStarterLight(AddressOf AsyncDecode)
        TSEncode.WhenFinished = AddressOf OnEncodeFinished
        TSDecode.WhenFinished = AddressOf OnDecodeFinished
        If Not DefaultPath = "Default" Then
            DefPath = DefaultPath
        Else
            DefPath = Application.StartupPath & "\test\"
        End If
        If (Not System.IO.Directory.Exists(DefPath)) Then
            System.IO.Directory.CreateDirectory(DefPath)
        End If
        If Not File.Exists(DefPath & "\test.txt") Then
            File.Create(DefPath & "\test.txt")
        End If
    End Sub
    Private Sub AsyncEncode(Data As Object)
        Dim StringArr() As String = DirectCast(Data, String())
        Dim Value As String = StringArr(0)
        Dim Password As String = StringArr(1)
        Dim DPath As String = StringArr(2)

        Dim wrapper As New EncryptedReadWrite(Password)
        Dim cipherText As String = wrapper.EncryptData(Value)

        MsgBox("The cipher text is: " & cipherText)
        My.Computer.FileSystem.WriteAllText(DPath & "\test.txt", cipherText, False)
    End Sub
    Private Function AsyncDecode(Data As Object) As Object
        Dim StringArr() As String = DirectCast(Data, String())
        Dim Password As String = StringArr(0)
        Dim DPath As String = StringArr(1)

        Dim cipherText As String = My.Computer.FileSystem.ReadAllText(DPath & "\test.txt")
        Dim wrapper As New EncryptedReadWrite(Password)

        ' DecryptData throws if the wrong password is used.
        Dim plainText As String
        Dim ErrorOccurred As Boolean
        Try
            plainText = wrapper.DecryptData(cipherText)
            MsgBox("The plain text is: " & plainText)
        Catch ex As System.Security.Cryptography.CryptographicException
            plainText = Nothing
            ErrorOccurred = True
            MsgBox("The data could not be decrypted with the password.")
        Finally
            wrapper.Dispose()
        End Try
        Dim Ret() As Object = {plainText, ErrorOccurred}
        Return Ret
    End Function
    Public Sub Encode(ByVal Value As String, ByVal Password As String)
        Dim Args() As String = {Value, Password, DefPath}
        TSEncode.Start(Args)
    End Sub
    Public Sub Decode(ByVal Password As String)
        Dim Args() As String = {Password, DefPath}
        TSEncode.Start(Args)
    End Sub
    Private Sub OnEncodeFinished(Result As Object)
        RaiseEvent EncodeFinished()
    End Sub
    Private Sub OnDecodeFinished(Result As Object)
        Dim ResultArr() As Object = DirectCast(Result, String())
        Dim ClearText As String = DirectCast(ResultArr(0), String)
        Dim ErrorOccurred As Boolean = DirectCast(ResultArr(1), Boolean)
        RaiseEvent DecodeFinished(ClearText, ErrorOccurred)
    End Sub
End Class
