Option Explicit

Function NORMDISTS(x As Integer, UVal As Integer) As Double
Dim Val As Double
Dim i As Integer
Dim AvgU As Double
Dim StdDevU As Double
Dim Offset As Double

If UVal <> 0 Then
    AvgU = UVal / 2 ' Gauß
    
    StdDevU = 0
    For i = 0 To UVal
        StdDevU = StdDevU + (i - AvgU) ^ 2 ' SUM ((i - average) squared)
    Next i
    
    StdDevU = Sqr(StdDevU / UVal)     ' variance devided by UVal
                                      ' standard deviation = square-root of variance
                                      
    Val = Application.WorksheetFunction.NormDist(CDbl(x), AvgU, StdDevU, False)
    
    Offset = 0
    For i = 0 To UVal
        Offset = Offset + Application.WorksheetFunction.NormDist(i, AvgU, StdDevU, False)
    Next i
    Offset = 1 - Offset
    
    Val = Val + Offset / UVal
Else
    Val = 0
End If
    
NORMDISTS = Val

End Function
