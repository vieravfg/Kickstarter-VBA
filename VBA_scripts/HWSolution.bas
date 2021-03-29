Attribute VB_Name = "HWSolution"
Sub Solution()

   For Each ws In Worksheets
   
   ws.Cells(1, 9).Value = "Ticker"
   ws.Cells(1, 10).Value = "Yearly Change"
   ws.Cells(1, 11).Value = "Percent Change"
   ws.Cells(1, 12).Value = "Total Stock Volume"
   
  'Set an initial variables
  Dim Ticker_Name As String
  Dim Vol_Total As Double
  Dim Summary_Table_Row As Integer
  Dim count As Integer
  Dim yrl_cng As Double
  Dim prc_cng  As Double
  
  Vol_Total = 0
  lastrow = ws.Cells(Rows.count, 1).End(xlUp).Row
  Summary_Table_Row = 2

  ' Loop through all tickers
  For i = 2 To lastrow

    ' Check if we are still within the same Ticker name, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
      
      Ticker_Name = ws.Cells(i, 1).Value
      
      Vol_Total = Vol_Total + ws.Cells(i, 7).Value

      ' Print the Ticker Name in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
      
      cls = ws.Cells(i, 6).Value
      
      count = ws.Application.WorksheetFunction.CountIf(Range("A:A"), Ticker_Name)
      
      opn = ws.Cells(i - count + 1, 3).Value
    
      
      yrl_cng = cls - opn
      
      
    ' Print the Yearly Change
    
      ws.Range("J" & Summary_Table_Row).Value = yrl_cng
    
    
      ' Print the Percent Change
            
             If opn <> 0 Then
              prc_cng = (cls - opn) / opn
              ws.Range("K" & Summary_Table_Row).Value = prc_cng
                ' Change number format for Percent Change %
              ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
              Else
               ws.Range("K" & Summary_Table_Row).Value = NA
               ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00"
              
              End If
              

      ' Print the Total Stock Volume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Vol_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total Stock Volume
      Vol_Total = 0

    ' If the cell immediately following a row is the same Name...
    Else
      ' Add to the Total Stock Volume
      Vol_Total = Vol_Total + ws.Cells(i, 7).Value
                                 
    End If

  Next i

'Highlight conditional formating

lastnewrow = ws.Cells(Rows.count, 9).End(xlUp).Row

'loop through all percent changes
For i = 2 To lastnewrow

 If ws.Cells(i, 10).Value < 0 Then
 
 ws.Cells(i, 10).Interior.ColorIndex = 3
 
 ElseIf ws.Cells(i, 10).Value = 0 Then
 
 ws.Cells(i, 10).Interior.ColorIndex = 0
 Else
 
 ws.Cells(i, 10).Interior.ColorIndex = 4
 
 End If
 
 Next i

'BONUS PART----------------------------------
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

 Dim max As Double
 Dim min As Double
 Dim tic As String
 Dim toc As String
 Dim tick As String
 
'For the Greatest Total Volume
max = ws.Cells(2, 12)

  For i = 1 To lastnewrow
     With ws.Cells(i + 1, 12)
     If .Value > max Then
            max = .Value
            tic = .Offset(0, -3).Value
      End If
      End With
Next i

    ws.Cells(4, 17).Value = max
    ws.Cells(4, 16).Value = tic
    
'Change number format for Total Stock Volume
 ws.Cells(4, 17).NumberFormat = "0.0000E+00"

' For the Greatest % Increase
max = ws.Cells(2, 11)

  For i = 1 To lastnewrow
     With ws.Cells(i + 1, 11)
     If .Value > max Then
            max = .Value
            tick = .Offset(0, -2).Value
      End If
      End With
Next i

    ws.Cells(2, 17).Value = max
    ws.Cells(2, 16).Value = tick
    
' Change number format for Greatest % increase
    
  ws.Cells(2, 17).NumberFormat = "0.00%"

' For the Greatest % Decrease

  min = ws.Cells(2, 11)

  For i = 1 To lastnewrow
     With ws.Cells(i + 1, 11)
     If .Value < min Then
            min = .Value
            toc = .Offset(0, -2).Value
      End If
      End With
Next i

    ws.Cells(3, 17).Value = min
    ws.Cells(3, 16).Value = toc
    
' Change number format for Greatest % increase
    
ws.Cells(3, 17).NumberFormat = "0.00%"

ws.Columns("A:Q").AutoFit
 
 Next ws

MsgBox "Solved"

End Sub

