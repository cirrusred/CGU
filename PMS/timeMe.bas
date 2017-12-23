Attribute VB_Name = "Module1"
Sub timeMe()
Dim startTimer As Date, endTimer As Date, totalTimer As Date
Dim hours As Integer, minutes As Integer, seconds As Integer


startTimer = Now
endTimer = Now

totalTimer = (endTimer - startTimer)
    

hours = Left(totalTimer, 2) - 12
minutes = Mid(totalTimer, 4, 2)
seconds = Right(Mid(totalTimer, 7, 2), 2)

If hours <> 0 Then
    minutes = minutes + (hours * 60)
End If

MsgBox "Total time: " & minutes & "mins, " & seconds & "secs."
End Sub
