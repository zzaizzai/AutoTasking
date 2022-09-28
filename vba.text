Public Function NextWeekEnd(FromDate As Date, AfterDays As Integer)
    TheDay = FromDate + AfterDays
    If Weekday(TheDay) = 1 Then
        TheDay = TheDay + 1
    ElseIf Weekday(TheDay) = 7 Then
        TheDay = TheDay + 2
    Else
        TheDay = TheDay
    End If
    NextWeekEnd = TheDay
End Function

Public Function NextThusedat(FromDate As Date, AfterDays As Integer)
    TheDay = FromDate + AfterDays
    For i = 1 To 10
        If Weekday(TheDay + i) = 3 Then
            Exit For
        End If
    Next i
    NextThusedat = TheDay + i
End Function
