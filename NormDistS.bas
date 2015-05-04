Option Explicit

Function NORMDISTS(x As Integer, UVal As Integer) As Double
Dim Val As Double
Dim i As Integer
Dim AvgU As Double
Dim StdDevU As Double

If UVal <> 0 Then
   AvgU = (UVal + 1) / 2 ' Gauß
    
    StdDevU = 0
    
    For i = 1 To UVal
        StdDevU = StdDevU + (i - AvgU) ^ 2 ' SUM i - average squared
    Next i
    
    StdDevU = (StdDevU / UVal) ^ 0.5  ' variance = SUM i - average squared; devided by UVal
                                      ' standard deviation = square-root of variance -> to the power of 0.5
                                      
    Val = Application.WorksheetFunction.NormDist(CDbl(x), AvgU, StdDevU, False)
Else
    Val = 0
End If
    
NORMDISTS = Val

End Function
