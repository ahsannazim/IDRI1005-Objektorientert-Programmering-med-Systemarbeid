Public Class DatabaseListEventArgs
    Inherits EventArgs
    Public Data As DataTable
    Public ErrorMessage As String = ""
    Public ErrorOccurred As Boolean
    Public Sub New()
    End Sub
    Public Sub New(D As DataTable, ByVal Err As Boolean, ErrorMessage As String)
        Data = D.Copy
        D.Dispose()
        ErrorOccurred = Err
    End Sub
End Class