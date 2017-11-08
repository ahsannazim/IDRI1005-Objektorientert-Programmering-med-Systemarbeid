Option Strict On
Option Explicit On
Option Infer Off
Imports System.Drawing

Public NotInheritable Class ColorHelper
    Public Shared Function Multiply(Prototype As Color, ByVal Ratio As Double) As Color
        Dim RGB() As Integer = {Prototype.R, Prototype.G, Prototype.B}
        For i As Integer = 0 To 2
            RGB(i) = CInt(RGB(i) * (Ratio))
            If RGB(i) > 255 Then
                RGB(i) = 255
            ElseIf RGB(i) < 0 Then
                RGB(i) = 0
            End If
        Next
        Return Color.FromArgb(RGB(0), RGB(1), RGB(2))
    End Function
    Public Overloads Shared Function Add(Prototype As Color, ByVal Amount As Integer) As Color
        Dim RGB() As Integer = {Prototype.R, Prototype.G, Prototype.B}
        For i As Integer = 0 To 2
            RGB(i) += Amount
            If RGB(i) > 255 Then
                RGB(i) = 255
            ElseIf RGB(i) < 0 Then
                RGB(i) = 0
            End If
        Next
        Return Color.FromArgb(RGB(0), RGB(1), RGB(2))
    End Function
    Public Overloads Shared Function Add(Prototype As Color, AddColor As Color) As Color
        Dim RGB() As Integer = {Prototype.R, Prototype.G, Prototype.B}
        RGB(0) += AddColor.R
        RGB(1) += AddColor.G
        RGB(2) += AddColor.B
        For i As Integer = 0 To 2
            If RGB(i) > 255 Then
                RGB(i) = 255
            ElseIf RGB(i) < 0 Then
                RGB(i) = 0
            End If
        Next
        Return Color.FromArgb(RGB(0), RGB(1), RGB(2))
    End Function
    Public Shared Function Mix(FirstColor As Color, SecondColor As Color, Optional Ratio As Double = 0.5) As Color
        Dim RGB(2) As Integer
        Dim RatioSecond As Double = 1 - Ratio
        RGB(0) = CInt((FirstColor.R * Ratio + SecondColor.R * RatioSecond) / 2)
        RGB(1) = CInt((FirstColor.G * Ratio + SecondColor.G * RatioSecond) / 2)
        RGB(2) = CInt((FirstColor.B * Ratio + SecondColor.B * RatioSecond) / 2)
        For i As Integer = 0 To 2
            If RGB(i) > 255 Then
                RGB(i) = 255
            ElseIf RGB(i) < 0 Then
                RGB(i) = 0
            End If
        Next
        Return Color.FromArgb(RGB(0), RGB(1), RGB(2))
    End Function
    Public Shared Function FillRemainingRGB(Prototype As Color, Optional Power As Double = 0.5) As Color
        Dim Remaining() As Integer = {255 - Prototype.R, 255 - Prototype.G, 255 - Prototype.B}
        Dim Colors() As Integer = {Prototype.R + CInt(Power * Remaining(0)), Prototype.G + CInt(Power * Remaining(1)), Prototype.B + CInt(Power * Remaining(2))}
        For i As Integer = 0 To 2
            If Colors(i) > 255 Then
                Colors(i) = 255
            ElseIf Colors(i) < 0 Then
                Colors(i) = 0
            End If
        Next
        Dim Ret As Color = Color.FromArgb(Colors(0), Colors(1), Colors(2))
        Return Ret
    End Function
End Class
