
''' <summary>
''' Provides functions to be used in conjunction with a timer to achieve an easing animation.
''' </summary>
Public Class EaseInOut
    ''' <summary>
    ''' Given an initial value, a final value, the current timer tick and the final timer tick, returns the value to assign to the eased property or number.
    ''' </summary>
    ''' <param name="InitialY">The initial to ease in from.</param>
    ''' <param name="GoalY">The final value to ease out to.</param>
    ''' <param name="CurrentX">The number of timer ticks that have elapsed.</param>
    ''' <param name="LastX">The number of timer ticks that have elapsed + the number of timer ticks that will elapse.</param>
    ''' <returns>A double with the eased value.</returns>
    Public Shared Function GetY(InitialY As Double, GoalY As Double, ByVal CurrentX As Integer, ByVal LastX As Integer) As Double
        Dim DoubleX As Double = CurrentX / LastX
        Dim DiffY As Double = GoalY - InitialY
        Dim DoubleXSquared As Double = DoubleX * DoubleX
        Dim Multiplier As Double = DoubleXSquared / (DoubleXSquared + (1 - DoubleX) * (1 - DoubleX))
        Return InitialY + DiffY * Multiplier
    End Function
End Class