VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub WBdry()

    Dim id As Integer
    Dim first_name As String
    Dim last_name As String
    Dim phone As Integer
    Dim date_brought As Date
    Dim description As String
    Dim cell As Range
    Dim btn As Button
    Dim ws As Worksheet
    Dim rng As Range
    Dim item As Range
    Dim btn2 As Button
    
    
    Set ws1 = ThisWorkbook.Sheets("Sheet1")
    Set cell = ws1.Range("P1")
    'Set btn = ws1.Buttons.Add(cell.Left, cell.Top, cell.Width, cell.Height)
    'btn.Text = "Button"
    'btn.OnAction = "ThisWorkbook.ShowMessage"
    Set rng = ws1.Range("N1:N100")
    For Each item In rng
        Set btn = ws1.Buttons.Add(item.Left, item.Top, item.Width * 2, item.Height)
        btn.Text = "Button"
        btn.OnAction = "ThisWorkbook.ShowMessage"
    Next item
    Set btn2 = ws1.Buttons.Add(cell.Left, cell.Top, cell.Width * 4, cell.Height)
    btn2.Text = "Aggregate Button"
    btn2.OnAction = "ThisWorkbook.Aggregate_em"
    
End Sub
Sub ShowMessage()
    MsgBox ("Message sent!")
End Sub
Sub Aggregate_em()
    Dim Sum As Double
    Dim rngAvg As Range
    Dim rngAvgE As Range
    Dim i As Long
    Dim lastRow As Long
    Dim Sum_paid As Double
    Dim Count As Integer
    
    Sum_paid = 0
    Count = 0
    MsgBox ("Button Pressed!")
    Set ws1 = ThisWorkbook.Sheets("Sheet1")
    Sum = Application.WorksheetFunction.Sum(ws1.Range("H1:H100"))
    
    'Set rngAvg = ws1.Range("H1:H" & lastRow)
    lastRow = ws1.Cells(ws1.Rows.Count, "G").End(xlUp).Row
    
    For i = 2 To lastRow
        If ws1.Cells(i, 7).Value = "Paid" Or ws1.Cells(i, 7).Value = "paid" Or ws1.Cells(i, 10).Value = "Yes" Or ws1.Cells(i, 10).Value = "yes" Or ws1.Cells(i, 10).Value = "Y" Or ws1.Cells(i, 10).Value = "y" Then
            Sum_paid = Sum_paid + ws1.Cells(i, 8).Value
            Count = Count + 1
        End If
    Next i
    'Avg = Application.WorksheetFunction.Average()
     ws1.Range("P2").Value = "Total Sum"
    ws1.Range("Q2").Value = "Sum of Paid Items"
    ws1.Range("R2").Value = "Average of Paid Items"
    
    ws1.Range("P3").Value = Sum
    ws1.Range("Q3").Value = Sum_paid
    ws1.Range("R3").Value = Sum_paid / Count
    
    
    
    
End Sub
