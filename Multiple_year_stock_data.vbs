Attribute VB_Name = "Module1"
Sub testing():
    ' define vars type
    Dim sheet1, sheet2, sheet3 As Worksheet
    Dim dicy As Collection
    Dim n As String
    
    ' set vars value
    Set sheet1 = Worksheets("2016")
    Set sheet2 = Worksheets("2015")
    Set sheet3 = Worksheets("2014")
    Set dicy = New Collection
    
    Total = 0
    
    ' create headers
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    
    
    ' loop through rows get unique tickers
    n = Range("A1", Range("A1").End(xlDown)).Rows.Count
    'MsgBox (n)
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error Resume Next
    For i = 2 To n
       dicy.Add Range("A" & i), Range("A" & i)
    Next i
    'MsgBox (dicy("A"))

    For i = 1 To dicy.Count
        Range("I" & i + 1) = dicy(i)
    Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    ' loop through tickers and add volume
    m = Range("I1", Range("I1").End(xlDown)).Rows.Count
    'MsgBox (m)
    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' loop through tickers and get yearly change
    For j = 2 To m
        If j <> m Then
            st = WorksheetFunction.Match(Range("I" & j), Range("A2:A" & n), 0)
            en = WorksheetFunction.Match(Range("I" & j + 1), Range("A2:A" & n), 0)
        Else
            st = WorksheetFunction.Match(Range("I" & j), Range("A2:A" & n), 0)
            en = n
        End If
        
        diff = Range("F" & en) - Range("C" & st + 1)
        
        If Range("C" & st + 1) <> 0 Then
            perc = diff / Range("C" & st + 1)
        Else
            perc = 0
        End If
        
        Range("J" & j) = diff
        Range("K" & j) = perc
        
        ' if stock grows then green else red
        If diff > 0 Then
            Range("J" & j).Interior.ColorIndex = 4
        Else
            Range("J" & j).Interior.ColorIndex = 3
        End If
        
        Range("K" & j).NumberFormat = "0.00%"
    Next j
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' loop through tickers and get total vol
    For j = 2 To m
        Range("L" & j) = WorksheetFunction.SumIf(Range("A2:A" & n), Range("I" & j), Range("G2:G" & n))
        Range("L" & j).NumberFormat = "0,0"
    Next j
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ' autofit cols
    Columns.AutoFit
End Sub


