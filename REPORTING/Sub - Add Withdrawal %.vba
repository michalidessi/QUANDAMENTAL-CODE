Sub AddWithDrawalPercent()

Dim withdr_date As Date
withdr_date = InputBox("Please input the withdrawal date in following format DD/MM/YYY:")
Dim withdr_percent As Long
withdr_percent = InputBox("Please input the withdrawal in %: ")
Dim next_ann_date As String
Dim percent As Double
Dim withdr_amount As Double
Dim fee_date As Date

'fee_date = 1 / 6 / 2020


percent = withdr_percent / 100


Dim lastrow As Integer

lastrow = Range("L" & Rows.Count).End(xlUp).Row

next_ann_date = Range("AA" & lastrow).Value
Range("AA" & lastrow).Value = withdr_date 'set withdr date


withdr_amount = Range("AJ" & lastrow) * percent




If Range("AD" & lastrow).Value > 0.05 And withdr_date > CDate("1/6/2020") Then


Range("AK" & lastrow).Value = withdr_amount * 0.95
Range("AM" & lastrow).Value = (Range("AJ" & lastrow) * 0.95) - withdr_amount 'set next year starting investment
Range("AN" & lastrow).Value = Range("AJ" & lastrow) * 0.05

Else

Range("AK" & lastrow).Value = withdr_amount
Range("AM" & lastrow).Value = Range("AJ" & lastrow) - withdr_amount
Range("AN" & lastrow).Value = 0

End If


Sheets("NewSection").Rows("2:2").Copy

Range("A" & lastrow + 1).Select

ActiveSheet.Paste

'Funds invested on: withdrawal date


Range("I" & lastrow + 1).Value = withdr_date 'funds invested on (next period)
Range("AA" & lastrow + 1).Value = next_ann_date ' set next ann date
Rows(lastrow).Interior.Color = RGB(242, 220, 219)


'Range("K" & lastrow + 1).Value = withdr_amount

End Sub
