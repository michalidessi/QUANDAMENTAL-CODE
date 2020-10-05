Option Explicit

Sub AddWithDrawal(ByVal withdr_amount As Double, ByVal withdr_date As Date)

  Dim rp_wb As Workbook
  Set rp_wb = ThisWorkbook
  rp_wb.Activate

  Dim rp_ws As Worksheet
  Set rp_ws = rp_wb.Sheets("Sales Data NEW")
  rp_ws.Activate

  Dim lastrow As Integer
  lastrow = Range("L" & Rows.Count).End(xlUp).Row

  Dim next_ann_date As Date
  next_ann_date = Range("AA" & lastrow).Value

  Dim current_period_start_date As Date
  current_period_start_date = Range("I" & lastrow).Value

  Range("AA" & lastrow).Value = withdr_date 'set withdr date
  Range("AK" & lastrow).Value = withdr_amount

    If Range("AD" & lastrow).Value > 0.05 And withdr_date > CDate("1/6/2020") Then
      Range("AN" & lastrow).Value = Range("AJ" & lastrow) * 0.05
    Else
      Range("AN" & lastrow).Value = 0
    End If

  Sheets("NewSection").Range("I2:Z2").Copy
  rp_ws.Range("I" & lastrow + 1).Select
  rp_ws.Paste

  Range("I" & lastrow + 1).Value = withdr_date 'funds invested on (next period)
  Range("AA" & lastrow + 1).Value = next_ann_date ' set next ann date

  Sheets("NewSection").Range("AB14:AM14").Copy
  rp_ws.Range("AB" & lastrow + 1).PasteSpecial
  rp_ws.Rows(lastrow).Interior.Color = RGB(242, 220, 219)

End Sub
