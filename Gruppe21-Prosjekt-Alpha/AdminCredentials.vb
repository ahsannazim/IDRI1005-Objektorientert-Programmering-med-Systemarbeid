Option Strict On
Option Explicit On
Option Infer Off

Imports AudiopoLib
Public Class AdminCredentials
    ''' <summary>
    ''' Encodes the specified text with the specified key, then saves the encoded text to the specified location.
    ''' </summary>
    ''' <param name="PlainText">The text to be encoded.</param>
    ''' <param name="Key">The key with which the encoded text can be decoded.</param>
    ''' <param name="PathIncludingBackSlash">The absolute path to the folder in which the decoded text will be saved.</param>
    ''' <param name="FileNameIncludingExtension">The name of the file (including extension) in which the decoded text will be saved</param>
    ''' <returns>A boolean value indicating whether or not the encoding was successful.</returns>
    Public Function Encode(PlainText As String, Key As String, PathIncludingBackSlash As String, FileNameIncludingExtension As String) As Boolean
        Try
            Dim wrapper As New CryptoServiceProvider(Key)
            Dim cipherText As String = wrapper.EncryptData(PlainText)
            My.Computer.FileSystem.WriteAllText(PathIncludingBackSlash & FileNameIncludingExtension, cipherText, False)
            'Wrapper.dispose
            Return True
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Decodes the text in the specified file using the specified key.
    ''' </summary>
    ''' <param name="PathIncludingBackSlash">The absolute path to the folder in which the decoded text will be saved.</param>
    ''' <param name="FileNameIncludingExtension">The name of the file (including extension) in which the decoded text will be saved</param>
    ''' <param name="Key">The key that was used to encode the text.</param>
    ''' <returns>If the correct key was provided: returns the decoded text. If the wrong key was provided: returns nothing.</returns>
    Public Function Decode(PathIncludingBackSlash As String, FileNameIncludingExtension As String, Key As String) As String
        Dim cipherText As String = My.Computer.FileSystem.ReadAllText(PathIncludingBackSlash & FileNameIncludingExtension)
        Dim wrapper As New CryptoServiceProvider(Key)
        ' DecryptData throws if the wrong password is used.
        Try
            Dim plainText As String = wrapper.DecryptData(cipherText)
            Return plainText
        Catch CryptoEx As System.Security.Cryptography.CryptographicException
            Return Nothing
        Catch DirEx As System.IO.DirectoryNotFoundException
            Throw
            'Finally
            '    wrapper.Dispose()
            'TODO: Find out if this is necessary
        End Try
    End Function
End Class
