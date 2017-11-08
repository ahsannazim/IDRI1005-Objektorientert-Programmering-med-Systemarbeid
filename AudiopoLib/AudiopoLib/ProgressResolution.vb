Option Strict On
Option Explicit On
Option Infer Off

' Kan gjøres om til enum

Public Class ProgressResolution
    Private M As Integer = 1
    Protected Friend Property Smoothness As Integer
        Get
            Return M
        End Get
        Set(value As Integer)
            M = value
        End Set
    End Property
    Public Shared ReadOnly Property Smooth As ProgressResolution
        Get
            Return New ProgressResolution(1)
        End Get
    End Property
    Public Shared ReadOnly Property Half As ProgressResolution
        Get
            Return New ProgressResolution(2)
        End Get
    End Property
    Public Shared ReadOnly Property Quarter As ProgressResolution
        Get
            Return New ProgressResolution(4)
        End Get
    End Property
    Private Sub New(Optional ByVal Multiplier As Integer = 1)
        M = Multiplier
    End Sub
End Class