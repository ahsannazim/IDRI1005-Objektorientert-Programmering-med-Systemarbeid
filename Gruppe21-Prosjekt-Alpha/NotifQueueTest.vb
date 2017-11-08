

'Public MustInherit Class NotificationQueue
'    Private ElementList As New List(Of NotificationQueueElement)

'    Public MustOverride Function ElementAtIndex(ByVal Index As )

'    Public Sub FindTest(Of TIn)()


'    Public Enum ElementProperty As Integer
'        ID = 0
'    End Enum
'End Class

'Public MustInherit Class NotificationQueueElement
'    Private varID As Object
'    Public Sub New()

'    End Sub
'End Class

'Public Class Egenerklæringsliste
'    Inherits NotificationQueue
'    Public Shadows Enum ElementProperty
'        TimeID = 0
'        Dato = 1
'        Tid = 2
'        DatoOgTid = 3
'        Fødselsnummer = 4
'        Godkjent = 5
'        AnsattID = 6
'        Reference = 7
'    End Enum
'    Private Erklæringsliste As New List(Of Egenerklæring)
'    Public Sub New()

'    End Sub
'    Public Sub Add(ByRef Egenerklæring As Egenerklæring)
'        Erklæringsliste.Add(Egenerklæring)
'    End Sub

'    Public ReadOnly Property TimeAtIndex(ByVal Index As Integer) As Egenerklæring
'        Get
'            Return Erklæringsliste(Index)
'        End Get
'    End Property
'    Public Function GetAllElementsWhere(ByVal Egenskap As TimeEgenskap, ByVal Verdi As Object, Optional ByVal Condition As ComparisonOperator = ComparisonOperator.EqualTo, Optional ByVal ReturnIfConditionIs As Boolean = True) As List(Of StaffTime)
'        Dim Ret As List(Of StaffTime)
'        Try
'            Ret = TimeListe.FindAll(Function(Time As StaffTime) As Boolean
'                                        Select Case Egenskap
'                                            Case TimeEgenskap.TimeID
'                                                Select Case Condition
'                                                    Case ComparisonOperator.EqualTo
'                                                        Return ((Time.TimeID = DirectCast(Verdi, Integer)) = ReturnIfConditionIs)
'                                                    Case Else
'                                                        Return False
'                                                End Select
'                                            Case TimeEgenskap.Dato
'                                                Select Case Condition
'                                                    Case ComparisonOperator.EqualTo
'                                                        Return ((Time.DatoOgTid.Date = DirectCast(Verdi, Date).Date) = ReturnIfConditionIs)
'                                                    Case ComparisonOperator.LessThan
'                                                        Return ((Time.DatoOgTid.Date.CompareTo(DirectCast(Verdi, Date).Date) < 0) = ReturnIfConditionIs)
'                                                    Case ComparisonOperator.GreaterThan
'                                                        Return ((Time.DatoOgTid.Date.CompareTo(DirectCast(Verdi, Date).Date) > 0) = ReturnIfConditionIs)
'                                                    Case Else
'                                                        Return False
'                                                End Select
'                                            Case TimeEgenskap.Tid
'                                                Select Case Condition
'                                                    Case ComparisonOperator.EqualTo
'                                                        Return ((Time.DatoOgTid.TimeOfDay = DirectCast(Verdi, TimeSpan)) = ReturnIfConditionIs)
'                                                    Case ComparisonOperator.GreaterThan
'                                                        Return ((Time.DatoOgTid.TimeOfDay.CompareTo(DirectCast(Verdi, TimeSpan)) > 0) = ReturnIfConditionIs)
'                                                    Case ComparisonOperator.LessThan
'                                                        Return ((Time.DatoOgTid.TimeOfDay.CompareTo(DirectCast(Verdi, TimeSpan)) < 0) = ReturnIfConditionIs)
'                                                    Case Else
'                                                        Return False
'                                                End Select
'                                            Case TimeEgenskap.DatoOgTid
'                                                Select Case Condition
'                                                    Case ComparisonOperator.EqualTo
'                                                        Return ((Time.DatoOgTid = DirectCast(Verdi, Date)) = ReturnIfConditionIs)
'                                                    Case ComparisonOperator.LessThan
'                                                        Return ((Time.DatoOgTid.CompareTo(DirectCast(Verdi, Date)) < 0) = ReturnIfConditionIs)
'                                                    Case ComparisonOperator.GreaterThan
'                                                        Return ((Time.DatoOgTid.CompareTo(DirectCast(Verdi, Date)) > 0) = ReturnIfConditionIs)
'                                                    Case Else
'                                                        Return False
'                                                End Select
'                                            Case TimeEgenskap.Fødselsnummer
'                                                Return ((Time.Fødselsnummer = DirectCast(Verdi, String)) = ReturnIfConditionIs)
'                                            Case TimeEgenskap.Godkjent
'                                                Return ((Time.Godkjent = DirectCast(Verdi, Boolean)) = ReturnIfConditionIs)
'                                            Case TimeEgenskap.AnsattID
'                                                If IsDBNull(Verdi) Then
'                                                    Return ((Time.AnsattID < 0) = ReturnIfConditionIs)
'                                                Else
'                                                    Return ((Time.AnsattID = DirectCast(Verdi, Integer)) = ReturnIfConditionIs)
'                                                End If
'                                            Case Else
'                                                Return (ReferenceEquals(Time, Verdi) = ReturnIfConditionIs)
'                                        End Select
'                                    End Function)
'        Catch ex As Exception
'            MsgBox(ex.Message)
'            Ret = Nothing
'        End Try
'        Return Ret
'    End Function
'    Public Function GetElementWhere(ByVal Egenskap As TimeEgenskap, ByVal Verdi As Object, Optional ByVal Condition As ComparisonOperator = ComparisonOperator.EqualTo, Optional ByVal ReturnIfConditionIs As Boolean = True) As StaffTime
'        Dim Ret As StaffTime
'        Try
'            Ret = TimeListe.Find(Function(Time As StaffTime) As Boolean
'                                     Select Case Egenskap
'                                         Case TimeEgenskap.TimeID
'                                             Select Case Condition
'                                                 Case ComparisonOperator.EqualTo
'                                                     Return ((Time.TimeID = DirectCast(Verdi, Integer)) = ReturnIfConditionIs)
'                                                 Case Else
'                                                     Return False
'                                             End Select
'                                         Case TimeEgenskap.Dato
'                                             Select Case Condition
'                                                 Case ComparisonOperator.EqualTo
'                                                     Return ((Time.DatoOgTid.Date = DirectCast(Verdi, Date).Date) = ReturnIfConditionIs)
'                                                 Case ComparisonOperator.LessThan
'                                                     Return ((Time.DatoOgTid.Date.CompareTo(DirectCast(Verdi, Date).Date) < 0) = ReturnIfConditionIs)
'                                                 Case ComparisonOperator.GreaterThan
'                                                     Return ((Time.DatoOgTid.Date.CompareTo(DirectCast(Verdi, Date).Date) > 0) = ReturnIfConditionIs)
'                                                 Case Else
'                                                     Return False
'                                             End Select
'                                         Case TimeEgenskap.Tid
'                                             Select Case Condition
'                                                 Case ComparisonOperator.EqualTo
'                                                     Return ((Time.DatoOgTid.TimeOfDay = DirectCast(Verdi, TimeSpan)) = ReturnIfConditionIs)
'                                                 Case ComparisonOperator.GreaterThan
'                                                     Return ((Time.DatoOgTid.TimeOfDay.CompareTo(DirectCast(Verdi, TimeSpan)) > 0) = ReturnIfConditionIs)
'                                                 Case ComparisonOperator.LessThan
'                                                     Return ((Time.DatoOgTid.TimeOfDay.CompareTo(DirectCast(Verdi, TimeSpan)) < 0) = ReturnIfConditionIs)
'                                                 Case Else
'                                                     Return False
'                                             End Select
'                                         Case TimeEgenskap.DatoOgTid
'                                             Select Case Condition
'                                                 Case ComparisonOperator.EqualTo
'                                                     Return ((Time.DatoOgTid = DirectCast(Verdi, Date)) = ReturnIfConditionIs)
'                                                 Case ComparisonOperator.LessThan
'                                                     Return ((Time.DatoOgTid.CompareTo(DirectCast(Verdi, Date)) < 0) = ReturnIfConditionIs)
'                                                 Case ComparisonOperator.GreaterThan
'                                                     Return ((Time.DatoOgTid.CompareTo(DirectCast(Verdi, Date)) > 0) = ReturnIfConditionIs)
'                                                 Case Else
'                                                     Return False
'                                             End Select
'                                         Case TimeEgenskap.Fødselsnummer
'                                             Return ((Time.Fødselsnummer = DirectCast(Verdi, String)) = ReturnIfConditionIs)
'                                         Case TimeEgenskap.Godkjent
'                                             Return ((Time.Godkjent = DirectCast(Verdi, Boolean)) = ReturnIfConditionIs)
'                                         Case TimeEgenskap.AnsattID
'                                             If IsDBNull(Verdi) Then
'                                                 Return ((Time.AnsattID < 0) = ReturnIfConditionIs)
'                                             Else
'                                                 Return ((Time.AnsattID = DirectCast(Verdi, Integer)) = ReturnIfConditionIs)
'                                             End If
'                                         Case Else
'                                             Return (ReferenceEquals(Time, Verdi) = ReturnIfConditionIs)
'                                     End Select
'                                 End Function)
'        Catch ex As Exception
'            MsgBox(ex.Message)
'            Ret = Nothing
'        End Try
'        Return Ret
'    End Function
'    Public Enum TimeEgenskap As Integer
'        TimeID = 0
'        Dato = 1
'        Tid = 2
'        DatoOgTid = 3
'        Fødselsnummer = 4
'        Godkjent = 5
'        AnsattID = 6
'        Reference = 7
'    End Enum
'    Public Enum ComparisonOperator As Integer
'        EqualTo = 0
'        GreaterThan = 1
'        LessThan = 2
'        BooleanTrue = 3
'        ReferencesEqual = 4
'    End Enum

'    Public Class Egenerklæring
'        Private varTimeID As Integer
'        Private varLand, varSkjemaString, varAnsattSvar As String
'        Private varGodkjent As Boolean
'        Public Property AnsattSvar As String
'            Get
'                Return varAnsattSvar
'            End Get
'            Set(value As String)
'                varAnsattSvar = value
'            End Set
'        End Property
'        Public Property TimeID As Integer
'            Get
'                Return varTimeID
'            End Get
'            Set(value As Integer)
'                varTimeID = value
'            End Set
'        End Property
'        Public Property Land As String
'            Get
'                Return varLand
'            End Get
'            Set(value As String)
'                varLand = Land
'            End Set
'        End Property
'        Public Property SkjemaString As String
'            Get
'                Return varSkjemaString
'            End Get
'            Set(value As String)
'                varSkjemaString = value
'            End Set
'        End Property
'        Public Property Godkjent As Boolean
'            Get
'                Return varGodkjent
'            End Get
'            Set(value As Boolean)
'                varGodkjent = value
'            End Set
'        End Property
'        Public Sub New(TimeID As Integer, SkjemaString As String, Land As String, Godkjent As Boolean)
'            varTimeID = TimeID
'            varSkjemaString = SkjemaString
'            varLand = Land
'            varGodkjent = Godkjent
'        End Sub
'    End Class
'End Class
'Public Class StaffTimeliste
'    Private Timeliste As New List(Of StaffTime)
'    Public Sub New()

'    End Sub
'    Public ReadOnly Property Timer As List(Of StaffTime)
'        Get
'            Return Timeliste
'        End Get
'    End Property
'    Public ReadOnly Property TimeAtIndex(ByVal Index As Integer) As StaffTime
'        Get
'            Return Timeliste(Index)
'        End Get
'    End Property
'    Public Function GetAllElementsWhere(ByVal Egenskap As TimeEgenskap, ByVal Verdi As Object, Optional ByVal Condition As ComparisonOperator = ComparisonOperator.EqualTo, Optional ByVal ReturnIfConditionIs As Boolean = True) As List(Of StaffTime)
'        Dim Ret As List(Of StaffTime)
'        Try
'            Ret = Timeliste.FindAll(Function(Time As StaffTime) As Boolean
'                                        Select Case Egenskap
'                                            Case TimeEgenskap.TimeID
'                                                Select Case Condition
'                                                    Case ComparisonOperator.EqualTo
'                                                        Return ((Time.TimeID = DirectCast(Verdi, Integer)) = ReturnIfConditionIs)
'                                                    Case Else
'                                                        Return False
'                                                End Select
'                                            Case TimeEgenskap.Dato
'                                                Select Case Condition
'                                                    Case ComparisonOperator.EqualTo
'                                                        Return ((Time.DatoOgTid.Date = DirectCast(Verdi, Date).Date) = ReturnIfConditionIs)
'                                                    Case ComparisonOperator.LessThan
'                                                        Return ((Time.DatoOgTid.Date.CompareTo(DirectCast(Verdi, Date).Date) < 0) = ReturnIfConditionIs)
'                                                    Case ComparisonOperator.GreaterThan
'                                                        Return ((Time.DatoOgTid.Date.CompareTo(DirectCast(Verdi, Date).Date) > 0) = ReturnIfConditionIs)
'                                                    Case Else
'                                                        Return False
'                                                End Select
'                                            Case TimeEgenskap.Tid
'                                                Select Case Condition
'                                                    Case ComparisonOperator.EqualTo
'                                                        Return ((Time.DatoOgTid.TimeOfDay = DirectCast(Verdi, TimeSpan)) = ReturnIfConditionIs)
'                                                    Case ComparisonOperator.GreaterThan
'                                                        Return ((Time.DatoOgTid.TimeOfDay.CompareTo(DirectCast(Verdi, TimeSpan)) > 0) = ReturnIfConditionIs)
'                                                    Case ComparisonOperator.LessThan
'                                                        Return ((Time.DatoOgTid.TimeOfDay.CompareTo(DirectCast(Verdi, TimeSpan)) < 0) = ReturnIfConditionIs)
'                                                    Case Else
'                                                        Return False
'                                                End Select
'                                            Case TimeEgenskap.DatoOgTid
'                                                Select Case Condition
'                                                    Case ComparisonOperator.EqualTo
'                                                        Return ((Time.DatoOgTid = DirectCast(Verdi, Date)) = ReturnIfConditionIs)
'                                                    Case ComparisonOperator.LessThan
'                                                        Return ((Time.DatoOgTid.CompareTo(DirectCast(Verdi, Date)) < 0) = ReturnIfConditionIs)
'                                                    Case ComparisonOperator.GreaterThan
'                                                        Return ((Time.DatoOgTid.CompareTo(DirectCast(Verdi, Date)) > 0) = ReturnIfConditionIs)
'                                                    Case Else
'                                                        Return False
'                                                End Select
'                                            Case TimeEgenskap.Fødselsnummer
'                                                Return ((Time.Fødselsnummer = DirectCast(Verdi, String)) = ReturnIfConditionIs)
'                                            Case TimeEgenskap.Godkjent
'                                                Return ((Time.Godkjent = DirectCast(Verdi, Boolean)) = ReturnIfConditionIs)
'                                            Case TimeEgenskap.AnsattID
'                                                If IsDBNull(Verdi) Then
'                                                    Return ((Time.AnsattID < 0) = ReturnIfConditionIs)
'                                                Else
'                                                    Return ((Time.AnsattID = DirectCast(Verdi, Integer)) = ReturnIfConditionIs)
'                                                End If
'                                            Case Else
'                                                Return (ReferenceEquals(Time, Verdi) = ReturnIfConditionIs)
'                                        End Select
'                                    End Function)
'        Catch ex As Exception
'            MsgBox(ex.Message)
'            Ret = Nothing
'        End Try
'        Return Ret
'    End Function
'    Public Function GetElementWhere(ByVal Egenskap As TimeEgenskap, ByVal Verdi As Object, Optional ByVal Condition As ComparisonOperator = ComparisonOperator.EqualTo, Optional ByVal ReturnIfConditionIs As Boolean = True) As StaffTime
'        Dim Ret As StaffTime
'        Try
'            Ret = Timeliste.Find(Function(Time As StaffTime) As Boolean
'                                     Select Case Egenskap
'                                         Case TimeEgenskap.TimeID
'                                             Select Case Condition
'                                                 Case ComparisonOperator.EqualTo
'                                                     Return ((Time.TimeID = DirectCast(Verdi, Integer)) = ReturnIfConditionIs)
'                                                 Case Else
'                                                     Return False
'                                             End Select
'                                         Case TimeEgenskap.Dato
'                                             Select Case Condition
'                                                 Case ComparisonOperator.EqualTo
'                                                     Return ((Time.DatoOgTid.Date = DirectCast(Verdi, Date).Date) = ReturnIfConditionIs)
'                                                 Case ComparisonOperator.LessThan
'                                                     Return ((Time.DatoOgTid.Date.CompareTo(DirectCast(Verdi, Date).Date) < 0) = ReturnIfConditionIs)
'                                                 Case ComparisonOperator.GreaterThan
'                                                     Return ((Time.DatoOgTid.Date.CompareTo(DirectCast(Verdi, Date).Date) > 0) = ReturnIfConditionIs)
'                                                 Case Else
'                                                     Return False
'                                             End Select
'                                         Case TimeEgenskap.Tid
'                                             Select Case Condition
'                                                 Case ComparisonOperator.EqualTo
'                                                     Return ((Time.DatoOgTid.TimeOfDay = DirectCast(Verdi, TimeSpan)) = ReturnIfConditionIs)
'                                                 Case ComparisonOperator.GreaterThan
'                                                     Return ((Time.DatoOgTid.TimeOfDay.CompareTo(DirectCast(Verdi, TimeSpan)) > 0) = ReturnIfConditionIs)
'                                                 Case ComparisonOperator.LessThan
'                                                     Return ((Time.DatoOgTid.TimeOfDay.CompareTo(DirectCast(Verdi, TimeSpan)) < 0) = ReturnIfConditionIs)
'                                                 Case Else
'                                                     Return False
'                                             End Select
'                                         Case TimeEgenskap.DatoOgTid
'                                             Select Case Condition
'                                                 Case ComparisonOperator.EqualTo
'                                                     Return ((Time.DatoOgTid = DirectCast(Verdi, Date)) = ReturnIfConditionIs)
'                                                 Case ComparisonOperator.LessThan
'                                                     Return ((Time.DatoOgTid.CompareTo(DirectCast(Verdi, Date)) < 0) = ReturnIfConditionIs)
'                                                 Case ComparisonOperator.GreaterThan
'                                                     Return ((Time.DatoOgTid.CompareTo(DirectCast(Verdi, Date)) > 0) = ReturnIfConditionIs)
'                                                 Case Else
'                                                     Return False
'                                             End Select
'                                         Case TimeEgenskap.Fødselsnummer
'                                             Return ((Time.Fødselsnummer = DirectCast(Verdi, String)) = ReturnIfConditionIs)
'                                         Case TimeEgenskap.Godkjent
'                                             Return ((Time.Godkjent = DirectCast(Verdi, Boolean)) = ReturnIfConditionIs)
'                                         Case TimeEgenskap.AnsattID
'                                             If IsDBNull(Verdi) Then
'                                                 Return ((Time.AnsattID < 0) = ReturnIfConditionIs)
'                                             Else
'                                                 Return ((Time.AnsattID = DirectCast(Verdi, Integer)) = ReturnIfConditionIs)
'                                             End If
'                                         Case Else
'                                             Return (ReferenceEquals(Time, Verdi) = ReturnIfConditionIs)
'                                     End Select
'                                 End Function)
'        Catch ex As Exception
'            MsgBox(ex.Message)
'            Ret = Nothing
'        End Try
'        Return Ret
'    End Function
'    Public Enum TimeEgenskap As Integer
'        TimeID = 0
'        Dato = 1
'        Tid = 2
'        DatoOgTid = 3
'        Fødselsnummer = 4
'        Godkjent = 5
'        AnsattID = 6
'        Reference = 7
'    End Enum
'    Public Enum ComparisonOperator As Integer
'        EqualTo = 0
'        GreaterThan = 1
'        LessThan = 2
'        BooleanTrue = 3
'        ReferencesEqual = 4
'    End Enum
'    Public Sub Add(ByRef Time As StaffTime)
'        Timeliste.Add(Time)
'    End Sub
'    Public Class StaffTime
'        Private varTimeID As Integer
'        Private varDatoOgTid As Date
'        Private varFødselsnummer As String
'        Private varGodkjent As Boolean
'        Private varAnsattID As Integer = -1
'        Public Property AnsattID As Integer
'            Get
'                Return varAnsattID
'            End Get
'            Set(value As Integer)
'                varAnsattID = value
'            End Set
'        End Property
'        Public Property TimeID As Integer
'            Get
'                Return varTimeID
'            End Get
'            Set(value As Integer)
'                varTimeID = value
'            End Set
'        End Property
'        Public Property DatoOgTid As Date
'            Get
'                Return varDatoOgTid
'            End Get
'            Set(value As Date)
'                varDatoOgTid = value
'            End Set
'        End Property
'        Public Property Godkjent As Boolean
'            Get
'                Return varGodkjent
'            End Get
'            Set(value As Boolean)
'                varGodkjent = value
'            End Set
'        End Property
'        Public Property Fødselsnummer As String
'            Get
'                Return varFødselsnummer
'            End Get
'            Set(value As String)
'                varFødselsnummer = value
'            End Set
'        End Property
'        Public Sub New(TimeID As Integer, DatoOgTid As Date, Godkjent As Boolean, Fødselsnummer As String, AnsattID As Object)
'            varTimeID = TimeID
'            varDatoOgTid = DatoOgTid
'            varGodkjent = Godkjent
'            varFødselsnummer = Fødselsnummer
'            If IsDBNull(AnsattID) Then
'                varAnsattID = -1
'            Else
'                varAnsattID = DirectCast(AnsattID, Integer)
'            End If
'        End Sub
'    End Class
'End Class