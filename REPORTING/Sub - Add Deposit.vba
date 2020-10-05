Sub AddDeposit(ByVal dep_amount As Double, ByVal dep_date As Date)

Dim rp_ws As Worksheet
Set rp_ws = rp_wb.Sheets("Sales Data NEW")

Dim lastrow As Integer
lastrow = Range("L" & Rows.Count).End(xlUp).Row

Dim mng_fee_pct As Double
mng_fee_pct = CDbl(rp_ws.Range("V19").Value)

Dim next_ann_date As Date
next_ann_date = Range("AA" & lastrow).Value

Dim current_period_start_date As Date
current_period_start_date = Range("I" & lastrow).Value

Dim date_diff As Long
date_diff = DateDiff("d", dep_date, next_ann_date) - 1

Dim mng_fee_prorata As Double
mng_fee_prorata = (date_diff / 365 * mng_fee_pct)

Dim mng_fee As Double
mng_fee = mng_fee_prorata * dep_amount

Range("AA" & lastrow).Value = dep_date
Range("AL" & lastrow).Value = dep_amount
Range("AI" & lastrow).Value = 0 ' set success fee as 0
Range("T" & lastrow).Value = CDbl(Range("T" & lastrow).Value) + mng_fee
Range("AM" & lastrow).Value = Range("AJ" & lastrow) + dep_amount

Sheets("NewSection").Range("I11:Z11").Copy
rp_ws.Range("I" & lastrow + 1).Select
rp_ws.Paste

Range("AA" & lastrow + 1).Value = next_ann_date ' set next ann date

Sheets("NewSection").Range("AB11:AN11").Copy
rp_ws.Range("AB" & lastrow + 1).PasteSpecial

rp_ws.Rows(lastrow).Interior.Color = RGB(226, 239, 218)

End Sub
