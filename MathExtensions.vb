Imports System.Runtime.CompilerServices

Public Module MathExtensions

    <Extension()>
    Function LimitInRange(Value As Double, Min As Double, Max As Double) As Double
        Select Case Value
            Case <= Min
                Return Min
            Case >= Max
                Return Max
            Case Else
                If IsNanOrInfinity(Value) Then
                    Return 0.0
                Else
                    Return Value
                End If
        End Select
    End Function

    <Extension()>
    Function IsNanOrInfinity(Value As Double) As Boolean
        Return Double.IsNaN(Value) OrElse Double.IsInfinity(Value)
    End Function

End Module