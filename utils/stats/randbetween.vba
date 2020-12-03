Option Explicit

Public Function RandBetween(minimum As Long, maximum As Long) As Integer
' Generate a random number that falls withn a range, inclusive.

    RandBetween = CInt(1 + Rnd() * (maximum - minimum))

End Function
