Attribute VB_Name = "General"
Public Sub delay(mSecs As Double)
Do
    cur = Timer
    DoEvents
Loop Until Timer - cur > mSecs / 1000
End Sub

