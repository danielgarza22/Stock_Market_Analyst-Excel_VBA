Attribute VB_Name = "Module1"
Sub Stock_Data()

'Declare variables
Dim Ticker As String
Dim LastRow As Long
Dim LastRow2 As Long
Dim Total_Volume As LongLong
Dim i As Long
Dim j As Long

'Create the header for all columns
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percentage Change"
Cells(1, 12).Value = "Total Stock Volume"
Range("K:K").Style = "Percent"
Range("K:K").NumberFormat = "0.00%"

'Assign width to summary columns
Columns("I:L").AutoFit

'Assign Bold to summary header
Range("I1:L1").Font.Bold = True

'Keep track of the Ticker in the summary table
Dim Summary_Table As Integer
Summary_Table = 2

'Get the last row of the table
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Create Loop for all Tickers
For i = 2 To LastRow
    
    'Get the first Open Price value
    If Cells(i, 1) <> Cells(i - 1, 1) Then
        Open_Price = Cells(i, 3)
    End If
    
    'Create condition to validate if its the same Ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'Set the Ticker
        Ticker = Cells(i, 1).Value
    
        'Set close Price
        Close_Price = Cells(i, 6).Value

        'Set Yearly Change
        Yearly_Change = Round(Close_Price - Open_Price, 2)
        
            'Set and Percentage
            If Open_Price <> 0 Then
                Percentage = (Open_Price - Close_Price) / Open_Price
            Else
                Percentage = 0
            End If
     
        'Add the Year Volume
        Total_Volume = Total_Volume + Cells(i, 7).Value
    
        'Print the summary table
        Range("I" & Summary_Table).Value = Ticker
        Range("J" & Summary_Table).Value = Yearly_Change
        Range("K" & Summary_Table).Value = Percentage
        Range("L" & Summary_Table).Value = Total_Volume
    
        'Add one to the summary table row
        Summary_Table = Summary_Table + 1
    
        'Reset the Year Volume
        Total_Volume = 0
    
    Else
        'Add the Year Total
        Total_Volume = Total_Volume + Cells(i, 7).Value
    
    End If

  Next i
        'Get the last row of the summary table
        LastRow2 = Cells(Rows.Count, 10).End(xlUp).Row
        
        For j = 2 To LastRow2
        
            Yearly_Change_Format = Cells(j, 10).Value
        
            'Create a condition to apply format to the Yearly Change data
            If Yearly_Change_Format > 0 Then
    
                'Change Yearly Change background color to blue
                Cells(j, 10).Interior.ColorIndex = 4
    
            'Change Yearly Change background color to red
            ElseIf Yearly_Change_Format < 0 Then
                
                'Change Yearly Change background color to red
                Cells(j, 10).Interior.ColorIndex = 3
    
            End If
        
        Next j

End Sub
