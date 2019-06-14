VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Ticker()

    'define variable

    Dim ticker_name As String
    Dim total_volume As Variant
    
    total_volume = 0
    
    Dim Summary_row As Variant
    
    Summary_row = 2
    
    Cells(1, 9).Value = "ticker_name"
    Cells(1, 10).Value = "total_volume"
    
    
    'loop through tickers
    For i = 2 To 705714
    
    'check if sam ticker name
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'set ticker name
    ticker_name = Cells(i, 1).Value
    
    'add volume total
    total_volume = total_volume + Cells(i, 7).Value
    
    'print ticker name to new rows
    Range("I" & Summary_row).Value = ticker_name
    
    'print volume total to new rows
    Range("J" & Summary_row).Value = total_volume
    
    'add one to summary row
    Summary_row = Summary_row + 1
    
    'reset toal volume
    total_volume = 0
    
    'if the same cells comes next
    Else
        total_volume = total_volume + Cells(i, 7).Value
        
        End If
    Next i
    
End Sub

