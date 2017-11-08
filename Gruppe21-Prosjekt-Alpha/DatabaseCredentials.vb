Public Class DatabaseCredentials
    Private DB_Server, DB_Name, DB_UID, DB_Pwd As String
    Public ReadOnly Property Server As String
        Get
            Return DB_Server
        End Get
    End Property
    Public ReadOnly Property Database As String
        Get
            Return DB_Name
        End Get
    End Property
    Public ReadOnly Property UserID As String
        Get
            Return DB_UID
        End Get
    End Property
    Public ReadOnly Property Password As String
        Get
            Return DB_Pwd
        End Get
    End Property
    Public Sub New(Server As String, Database As String, UserID As String, Password As String)
        DB_Server = Server
        DB_Name = Database
        DB_UID = UserID
        DB_Pwd = Password
    End Sub
End Class