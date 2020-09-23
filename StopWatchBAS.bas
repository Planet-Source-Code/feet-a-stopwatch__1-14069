Attribute VB_Name = "StopWatch"
Dim Hours As Integer
Dim Minutes As Integer
Dim Seconds As Integer
Dim Days As Integer
Dim AddMinutes As Boolean
Dim AddHours As Boolean
Sub StopWatch(WhatLabel As Label)
'to work in a label
'just put in a timer with an interval of about 950:
'call stopwatch(label that time will be displayed on)

If Seconds = 60 Then
AddMinutes = True
addSeconds = 0
Else
Seconds = Seconds + 1
End If
'see's if the amount of seconds is 60 so it can go to MinutesAdd
'if it's not it will add 1 second to the seconds
If AddMinutes = True Then
If Minutes = 60 Then
AddHours = True
Minutes = 0
Else
Minutes = Minutes + 1
End If
End If
'see's if the amount of minutes is 60 so it can go to HoursAdd
'if it's not it will add one minute to the minutes
If AddHours = True Then
If Hours = 24 Then
Hours = 0
Days = Days + 1
MsgBox Days & " day(s) have gone by since u started the stopwatch.", vbInformation, "Days That Have Gone By"
Else
Hours = Hours + 1
End If
End If
'see's if the amount of minutes is 24 so it can go to DaysAdd
'if it's not it will add one hour to the hours
'if Days is not zero, it will pop up a message box saying how many days have gone by since u started the stopwatch

WhatLabel.Caption = Format(Hours, "00") & ":" & Format(Minutes, "00") & ":" & Format(Seconds, "00")
End Sub
Public Sub Pause(duration As Long)
'i did not write this pause sub!!!!
'All credits for this sub goto Dos
       Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= duration
        DoEvents
    Loop
End Sub

Sub Restart(WhatLabel As Label)
WhatLabel.Caption = "00:00:00"
Seconds = 0
Minutes = 0
Hours = 0
Days = 0
'makes ur label no seconds, minutes, days or hours
End Sub
