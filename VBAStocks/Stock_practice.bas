Attribute VB_Name = "Module2"
Sub Stock()
'1.1 Declare Active Worksheet Name
Dim aws As String
aws = ActiveSheet.Name

'1.2 Setting up Column names

Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"
Range("M1").Value = "Total Stock Volume"

Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

Range("S1").Value = "First Open"
Range("T1").Value = "Last Close"

'---
'2. Summary of each tickers

'2.1 Calculating Unique Stock - not all stock shows up 262 time!

''2.1.1 Ticker Name
''Use Record Macro to select all Column A, Paste to J2, then Remove Duplicates
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("J1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("J:J").Select
    Application.CutCopyMode = False
    ActiveSheet.Range("$J$1:$J$1000000").RemoveDuplicates Columns:=1, Header:= _
        xlYes
        
''2.1.2 Count Number of stocks, and number of entries at raw
Dim stk_cnt As Integer
stk_cnt = Worksheets(aws).Range("J:J").Cells.SpecialCells(xlCellTypeConstants).Count - 1

Dim raw_cnt As Long
raw_cnt = Worksheets(aws).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count - 1

'2.2 Determining Open and Closing price
Dim Year_Chg As Double
Dim Pct_Chg As Double
Dim tot_vol As Long
Dim first_price As Double
Dim last_price As Double
Dim raw_tck_name As String
Dim sum_tck_name As String
Dim raw_tck_name_last As String 'used for last price determination


''Initializing Value
first_price = Range("C2").Value
Cells(3, 19).Value = first_price
Dim i As Long
Dim j As Long
i = 1
j = 1
sum_tck_name = Cells(1, 10).Value
raw_tck_name = Cells(1, 1).Value


Do Until i = stk_cnt + 2
        sum_tck_name = Cells(i, 10).Value
        raw_tck_name = Cells(j, 1).Value
        raw_tck_name_last = Cells(j + 1, 1).Value
        
        'determining open price
        If raw_tck_name <> sum_tck_name Then
        first_price = Cells(j, 3).Value
        Cells(i + 1, 19).Value = first_price
        
        'determining the last close price, yearly change, percent change
            If raw_tck_name_last <> sum_tck_name Then
            Cells(i, 20).Value = Cells(j - 1, 6).Value
            End If
        i = i + 1
        End If
        j = j + 1
Loop

'2.3 Applying Year_Chg, Pct_Chg, tot_vol to Dashboard
Dim k As Integer

k = 2
Do Until k = stk_cnt + 2
    sum_tck_name = Cells(k, 10).Value
    first_price = Cells(k, 19).Value
    last_price = Cells(k, 20).Value
    Year_Chg = last_price - first_price
    
    If first_price = 0 Then
        Pct_Chg = 0
    Else
        Pct_Chg = ((Year_Chg) / first_price)
    End If
    Cells(k, 11).Value = Year_Chg
    Cells(k, 12).Value = Pct_Chg
    
    If Year_Chg < 0 Then
        Cells(k, 11).Interior.ColorIndex = 3
    ElseIf Year_Chg > 0 Then
        Cells(k, 11).Interior.ColorIndex = 4
    End If
    
    
    k = k + 1
Loop

'2.4 total_volume
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(C[-12],RC[-3],C[-6])"
    Range("M2").Select
    Selection.AutoFill Destination:=Range("M2:M" & stk_cnt + 1)
    Range("M2:M" & stk_cnt + 1).Select
    
'2.5 Applying Format
    Columns("L:L").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    
    Range("Q2:Q3").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Range("Q4").Select
    Selection.NumberFormat = "0.00E+00"
'---
'3. Greatest Increase, Decrease, Total Volume

''3.1 Declare "greatest" variables
Dim great_inc_tick As String
Dim great_inc_val As Double
Dim great_dec_tick As String
Dim great_dec_val As Double
Dim great_tot_tick As String
Dim great_tot_val As Double

''3.2 Find the Value of max increase, min increase, max total volume, then express to cell
great_inc_val = WorksheetFunction.Max(Range("L:L"))
Range("Q2").Value = great_inc_val
great_dec_val = WorksheetFunction.Min(Range("L:L"))
Range("Q3").Value = great_dec_val
great_tot_val = WorksheetFunction.Max(Range("M:M"))
Range("Q4").Value = great_tot_val


''3.3 Look up the associated stock ticker(using index match technique)

great_inc_tick = Application.WorksheetFunction.Index(Sheets(aws).Range("J3:J1000000"), Application.WorksheetFunction.Match(great_inc_val, Sheets(aws).Range("L3:L1000000"), 0))
Range("P2").Value = great_inc_tick
great_dec_tick = Application.WorksheetFunction.Index(Sheets(aws).Range("J:J"), Application.WorksheetFunction.Match(great_dec_val, Sheets(aws).Range("L:L"), 0))
Range("P3").Value = great_dec_tick
great_tot_tick = Application.WorksheetFunction.Index(Sheets(aws).Range("J:J"), Application.WorksheetFunction.Match(great_tot_val, Sheets(aws).Range("M:M"), 0))
Range("P4").Value = great_tot_tick

'final touch
Range("J1").Value = "Ticker"

End Sub
