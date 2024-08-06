Attribute VB_Name = "Module1"
Sub Statistics():

'Store the Original Worksheet as a Variable
'===================================
    Dim ws_original As Worksheet
    Set ws_original = Worksheets(1)
   
 'Find the last row of a column
 '===================================
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
'Creating Worksheet
'===================================
    Dim ws As Worksheet
    With ThisWorkbook
        Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        ws.Name = "General Statistics"
    End With
    

'Average Function
'===================================
Const startcol As Long = 8
Dim endcol As Long
Dim rng As Range
Dim avg As Double

endcol = ws_original.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        '3 was subtracted since the 'Hyperlink to Atlas Website' is not an integer

' Loop through the specified columns
    For Column = startcol To endcol
        ' Show Title of Value
        ' Shifting the rows to start at the 2nd row and 1st column
        Cells(Column - 6, 1).Value = ws_original.Cells(1, Column).Value
        
        ' Calculate the formula
        Set rng = Range(ws_original.Cells(Column - 5, Column), ws_original.Cells(lastRow, Column))
        
        ' Calculate the average
        On Error Resume Next
        avg = Application.WorksheetFunction.Average(rng)
        If Err.Number <> 0 Then
            avg = "N/A" ' Handle errors or empty ranges
            Err.Clear
        End If
        On Error GoTo 0
        
        Cells(Column - 6, 2).Value = avg
    Next Column

'Min Function
'===================================
Dim min As Double

' Loop through the specified columns
    For Column = startcol To endcol
        ' Show Title of Value
        ' Shifting the rows to start at the 2nd row and 1st column
        Cells(Column - 6, 1).Value = ws_original.Cells(1, Column).Value
        
        ' Calculate the formula
        Set rng = Range(ws_original.Cells(Column - 5, Column), ws_original.Cells(lastRow, Column))
        
        min = Application.WorksheetFunction.min(rng)
        Cells(Column - 6, 3).Value = min
    Next Column
    
'Max Function
'===================================
Dim max As Double

' Loop through the specified columns
    For Column = startcol To endcol
        ' Show Title of Value
        ' Shifting the rows to start at the 2nd row and 1st column
        Cells(Column - 6, 1).Value = ws_original.Cells(1, Column).Value
        
        ' Calculate the formula
        Set rng = Range(ws_original.Cells(Column - 5, Column), ws_original.Cells(lastRow, Column))
        
        max = Application.WorksheetFunction.max(rng)
        Cells(Column - 6, 4).Value = max
    Next Column

'Standard Deviation Function
'===================================
'STDEV = Partial Sampling of Whole Poplutation(-N for Non-Bias)
'STDEVP = Entire Population Sampling (Biased)
Dim sd As Double
Dim sdp As Double

' Loop through STDEV
    For Column = startcol To endcol
        ' Show Title of Value
        ' Shifting the rows to start at the 2nd row and 1st column
        Cells(Column - 6, 1).Value = ws_original.Cells(1, Column).Value
        
        ' Calculate the formula
        Set rng = Range(ws_original.Cells(Column - 5, Column), ws_original.Cells(lastRow, Column))
        
        On Error Resume Next
        sd = Application.WorksheetFunction.StDev(rng)
        If Err.Number <> 0 Then
            sd = "N/A" ' Handle errors or empty ranges
            Err.Clear
        End If
        On Error GoTo 0
        
        Cells(Column - 6, 5).Value = sd
    Next Column

' Loop through STDEVP
    For Column = startcol To endcol
        ' Show Title of Value
        ' Shifting the rows to start at the 2nd row and 1st column
        Cells(Column - 6, 1).Value = ws_original.Cells(1, Column).Value
        
        ' Calculate the formula
        Set rng = Range(ws_original.Cells(Column - 5, Column), ws_original.Cells(lastRow, Column))
        
        On Error Resume Next
        sdp = Application.WorksheetFunction.StDev_P(rng)
        If Err.Number <> 0 Then
            sdp = "N/A" ' Handle errors or empty ranges
            Err.Clear
        End If
        On Error GoTo 0
        
        Cells(Column - 6, 6).Value = sdp
    Next Column
    
'Sample Variance Function
'===================================
'VAR.S = Calculates variance for a sample of a population (Ignores Empty Cells)
'VAR.P = Calculates variance for the entire population
Dim sv As Double
Dim svp As Double

' Loop through for VAR.S
    For Column = startcol To endcol
        ' Show Title of Value
        ' Shifting the rows to start at the 2nd row and 1st column
        Cells(Column - 6, 1).Value = ws_original.Cells(1, Column).Value
        
        ' Calculate the formula
        Set rng = Range(ws_original.Cells(Column - 5, Column), ws_original.Cells(lastRow, Column))
        
        On Error Resume Next
        sv = Application.WorksheetFunction.Var_S(rng)
        If Err.Number <> 0 Then
            svp = "N/A" ' Handle errors or empty ranges
            Err.Clear
        End If
        On Error GoTo 0
    
        Cells(Column - 6, 7).Value = sv
    Next Column

' Loop through for VAR.P
    For Column = startcol To endcol
        ' Show Title of Value
        ' Shifting the rows to start at the 2nd row and 1st column
        Cells(Column - 6, 1).Value = ws_original.Cells(1, Column).Value
        
        ' Calculate the formula
        Set rng = Range(ws_original.Cells(Column - 5, Column), ws_original.Cells(lastRow, Column))
        
        ' Calculate the formula
        On Error Resume Next
        svp = Application.WorksheetFunction.Var_P(rng)
        If Err.Number <> 0 Then
            svp = "N/A" ' Handle errors or empty ranges
            Err.Clear
        End If
        On Error GoTo 0
        
        Cells(Column - 6, 8).Value = svp
        
    Next Column

'Formatting Columns
'===================================
'Column Titles
Cells(1, 2).Value = "Avg"
Cells(1, 3).Value = "Min"
Cells(1, 4).Value = "Max"
Cells(1, 5).Value = "STDev"
Cells(1, 6).Value = "STDevP"
Cells(1, 7).Value = "VAR.S"
Cells(1, 8).Value = "VAR.P"

'AutoFit Text
ws.Columns("A").AutoFit

'Simplify values so it shows only the tenth place
Dim rng_new As Range
Set rng = Range(ws.Cells(2, 2), ws.Cells(63, 8))
rng.NumberFormat = "0.0"

'Delete unnecessary rows
Dim rng_del As Range
Set rng_del = Range("A64:A66")
rng_del.EntireRow.Delete

End Sub
