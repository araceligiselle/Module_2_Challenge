Attribute VB_Name = "RunItAll"
'Ticker
Sub ticker_allsheets()

Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        Ticker
    Next ws
End Sub

Sub Ticker()
    Dim Ticker_name As String
    Dim Ticker_total As Double
    Ticker_name = ""
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker_name = Cells(i, 1).Value
            Ticker_total = Ticker_total + Cells(i, 2).Value
            
            Range("I" & Summary_Table_Row).Value = Ticker_name
            
            Summary_Table_Row = Summary_Table_Row + 1
            Ticker_total = 0
        Else
            Ticker_total = Ticker_total + Cells(i, 2).Value
        End If
    Next i
  
  
End Sub

'Yearly Change
Sub YearlyChangeOnAllSheets()
 Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Yearly_Change ws
    Next ws
End Sub

Sub Yearly_Change(ws As Worksheet)
    Dim Ticker_name As String
    Dim Opening_price As Double
    Dim Closing_price As Double
    Dim Yearly_Change As Double
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Dim YearStartRow As Long
    YearStartRow = 2
    
    Dim RowCount As Long
    RowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    
    For i = 2 To RowCount
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            Ticker_name = ws.Cells(i, 1).Value
            Opening_price = ws.Cells(YearStartRow, 3).Value
            Closing_price = ws.Cells(i, 6).Value
            Yearly_Change = Closing_price - Opening_price
            
            ws.Range("I" & Summary_Table_Row).Value = Ticker_name
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            
            If Yearly_Change < 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            ElseIf Yearly_Change > 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            End If
            
            Summary_Table_Row = Summary_Table_Row + 1
            YearStartRow = i + 1
        End If
    Next i
End Sub

'Percent Change
Sub Percent_change()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Percent_change_ws ws
    Next ws
End Sub

Sub Percent_change_ws(ws As Worksheet)
    Dim Ticker_name As String
    Dim Opening_price As Double
    Dim Closing_price As Double
    Dim Yearly_Change As Double
    Dim Percent_change As Double
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Dim YearStartRow As Long
    YearStartRow = 2
    
    Dim RowCount As Long
    RowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    
    For i = 2 To RowCount
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            Ticker_name = ws.Cells(i, 1).Value
            Opening_price = ws.Cells(YearStartRow, 3).Value
            Closing_price = ws.Cells(i, 6).Value
            Yearly_Change = Closing_price - Opening_price
            Percent_change = Yearly_Change / Opening_price * 100
            
            ws.Range("I" & Summary_Table_Row).Value = Ticker_name
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            ws.Range("K" & Summary_Table_Row).Value = Format(Percent_change, "0.00%")
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            YearStartRow = i + 1
        End If
    Next i
End Sub

'Total Stock
Sub Total_stock_allsheets()

Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Total_stock ws
    Next ws
End Sub

Sub Total_stock(ws As Worksheet)

 Dim Ticker_name As String
    Dim Opening_price As Double
    Dim Closing_price As Double
    Dim Yearly_Change As Double
    Dim Percent_change As Double
    Dim Total_volume As Double
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Dim YearStartRow As Long
    YearStartRow = 2
    
    Dim RowCount As Long
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To RowCount
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            Ticker_name = Cells(i, 1).Value
            Opening_price = Cells(YearStartRow, 3).Value
            Closing_price = Cells(i, 6).Value
            Yearly_Change = Closing_price - Opening_price
            Percent_change = Yearly_Change / Opening_price * 100
            Total_volume = Application.WorksheetFunction.Sum(Range(Cells(YearStartRow, 7), Cells(i, 7)))
            
            Range("I" & Summary_Table_Row).Value = Ticker_name
            Range("J" & Summary_Table_Row).Value = Yearly_Change
            Range("K" & Summary_Table_Row).Value = Percent_change
            Range("L" & Summary_Table_Row).Value = Total_volume
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            YearStartRow = i + 1
        End If
    Next i

End Sub

'Greatest
Sub greatest_allsheets()

Dim ws As Worksheet
Dim sheetNames As Variant
    sheetNames = Array("2018", "2019", "2020")
     For Each sheetName In sheetNames
        Set ws = ThisWorkbook.Worksheets(sheetName)
        greatest ws
    Next sheetName
End Sub

Sub greatest(ws As Worksheet)

   Dim Ticker_name As String
    Dim Opening_price As Double
    Dim Closing_price As Double
    Dim Yearly_Change As Double
    Dim Percent_change As Double
    Dim Total_volume As Double
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Dim YearStartRow As Long
    YearStartRow = 2
    
    Dim RowCount As Long
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim Greatest_increase As Double
    Dim Greatest_decrease As Double
    Dim Greatest_volume As Double
    Dim Greatest_increase_ticker As String
    Dim Greatest_decrease_ticker As String
    Dim Greatest_volume_ticker As String
    
    Greatest_increase = 0
    Greatest_decrease = 0
    Greatest_volume = 0
    
    For i = 2 To RowCount
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            Ticker_name = Cells(i, 1).Value
            Opening_price = Cells(YearStartRow, 3).Value
            Closing_price = Cells(i, 6).Value
            Yearly_Change = Closing_price - Opening_price
            Percent_change = Yearly_Change / Opening_price * 100
            Total_volume = Application.WorksheetFunction.Sum(Range(Cells(YearStartRow, 7), Cells(i, 7)))
            
            Range("I" & Summary_Table_Row).Value = Ticker_name
            Range("J" & Summary_Table_Row).Value = Yearly_Change
            Range("K" & Summary_Table_Row).Value = Percent_change
            Range("L" & Summary_Table_Row).Value = Total_volume
            
            If Percent_change > Greatest_increase Then
                Greatest_increase = Percent_change
                Greatest_increase_ticker = Ticker_name
            ElseIf Percent_change < Greatest_decrease Then
                Greatest_decrease = Percent_change
                Greatest_decrease_ticker = Ticker_name
            End If
            
            If Total_volume > Greatest_volume Then
                Greatest_volume = Total_volume
                Greatest_volume_ticker = Ticker_name
            End If
            
            Summary_Table_Row = Summary_Table_Row + 1
            YearStartRow = i + 1
        End If
    Next i
    
    Range("O2").Value = Greatest_increase_ticker
    Range("O3").Value = Greatest_decrease_ticker
    Range("O4").Value = Greatest_volume_ticker
    
    Range("P2").Value = Greatest_increase
    Range("P3").Value = Greatest_decrease
    Range("P4").Value = Greatest_volume
End Sub

