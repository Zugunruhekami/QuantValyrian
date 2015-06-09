Public Sub CommandButton1_Click()
ReDim DiscFactor(1 To 6)
DiscFactor(1) = 1
For i = 2 To 6
    Dim SwapRate As Double: SwapRate = Cells(20 + i, 3).Value
    Dim TimeDiff As Double: TimeDiff = 1
    Dim RollingSum As Double
    RollingSum = DiscFactor(1) * TimeDiff
    For j = 1 To i - 1
        RollingSum = RollingSum + DiscFactor(j) * TimeDiff
    Next j
    DiscFactor(i) = (1 - SwapRate * RollingSum) / (1 + SwapRate * TimeDiff)
    Cells(20 + i, 9).Value = DiscFactor(i)
Next i

End Sub
