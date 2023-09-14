Attribute VB_Name = "Module2"
Sub wallstreet()

'loop through each sheet in workbook
For Each ws In Worksheets

'create a for loop for the dataset

'define and calculate the last row of the dataset
last = Cells(Rows.Count, 1).End(xlUp).Row

'define row counter for open price that starts at 0
Dim open_row As Double
open_row = 0

'define year opening price
Dim year_open As Double

'define year closing price
Dim year_close As Double

'define percent change
Dim percent_change As Double

'define yearly change
Dim yearly_change As Double

'define volume sum and set it to 0
Dim volume As Double
volume = 0

'define ticker
Dim ticker As String

' define table row and assign it to start on the second row
Dim table_row As Integer
table_row = 2

'assign table labels
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "Yearly Change"
  ws.Cells(1, 11).Value = "Percent Change"
  ws.Cells(1, 12).Value = "Total Stock Volume"


  'loop through all rows in the data set starting at 2
  For I = 2 To last
  
  
    ' if ticker changes in next row...
    ' syntax from student assignment "Credit Charges"
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
    

     ' assign ticker value
      ticker = ws.Cells(I, 1).Value

         ' assign ticker to table
          ws.Range("i" & table_row).Value = ticker
      
      'add to the total volume
      volume = volume + ws.Cells(I, 7).Value
        
         'assign total volume to table
         ws.Range("L" & table_row).Value = volume
        
      'assign year closing value
       year_close = ws.Cells(I, 6).Value
            
                 'calculate yearly change
                 yearly_change = year_close - year_open
            
            'assign yearly change to table
             ws.Range("j" & table_row).Value = yearly_change
            
                'calculate percentage change
                percentage_change = yearly_change / year_open
                
            'assign percentage change to table
             ws.Range("k" & table_row).Value = percentage_change
                
                'change data format to percentage with 2 decimals
                'syntax from Statology
                'https://www.statology.org/vba-percentage-format/
                ws.Range("k" & table_row).NumberFormat = "0.00%"
                
                'conditional formatting for yearly change
                If ws.Range("j" & table_row).Value > 0 Then
                ws.Range("j" & table_row).Interior.ColorIndex = 4
                ElseIf ws.Range("j" & table_row).Value < 0 Then
                ws.Range("j" & table_row).Interior.ColorIndex = 3
                ElseIf ws.Range("j" & table_row).Value = 0 Then
                ws.Range("j" & table_row).Interior.ColorIndex = 6
                End If
                
                'conditional formatting for percent change
                If ws.Range("k" & table_row).Value > 0 Then
                ws.Range("k" & table_row).Interior.ColorIndex = 4
                ElseIf ws.Range("k" & table_row).Value < 0 Then
                ws.Range("k" & table_row).Interior.ColorIndex = 3
                ElseIf ws.Range("k" & table_row).Value = 0 Then
                ws.Range("k" & table_row).Interior.ColorIndex = 6
                End If
    
            
                
      ' move to next table row
      table_row = table_row + 1
      
      'reset total volume
      volume = 0
      
      'reset opening price row counter
      open_row = 0
    
    'if ticker is the same in next row...
    Else
    
        'add to the total volume
        volume = volume + ws.Cells(I, 7).Value
        
        'add to the year open row counter
        open_row = open_row + 1
        
            'the first row for each ticker represents the first day
            'of the year
            If open_row = 1 Then
            
                'assign year opening value
                 year_open = ws.Cells(I, 3).Value
                 
            End If
            
    End If
    
' next dataset row
  Next I
   
'create a for loop for the table

'define and calculate the last row of the table
table_last = Cells(Rows.Count, 9).End(xlUp).Row

'define greatest percent increase and calculate using max
Dim greatest_increase As Double
greatest_increase = WorksheetFunction.Max(ws.Range("K:K"))

'define greatest percent decrease and calculate using min
Dim greatest_decrease As Double
greatest_decrease = WorksheetFunction.Min(ws.Range("K:K"))

'define greatest total volume and calculate using max
Dim greatest_volume As Double
greatest_volume = WorksheetFunction.Max(ws.Range("L:L"))

'define table ticker
Dim t_ticker As String

'assign table labels
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

    'loop through all rows in the table starting with  2
    For r = 2 To table_last
    
    
        'check table value against greatest percent increase
        If ws.Cells(r, 11).Value = greatest_increase Then
        
            'assign greatest percent increase value to table
            ws.Cells(2, 17).Value = greatest_increase
            
                'change data format to percentage with 2 decimals
                ws.Cells(2, 17).NumberFormat = "0.00%"
            
            'assign table ticker value
            t_ticker = ws.Cells(r, 9).Value
                
                'assign table ticker to table
                ws.Cells(2, 16).Value = t_ticker
                
        End If

        'check table value against greatest percent decrease
        If ws.Cells(r, 11).Value = greatest_decrease Then
        
            'assign greatest percent decrease to table
            ws.Cells(3, 17).Value = greatest_decrease
                
                'change data format to percentage with 2 decimals
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            'assign table ticker value
            t_ticker = ws.Cells(r, 9).Value
            
                'assign table ticker to table
                ws.Cells(3, 16).Value = t_ticker
          
        End If
        
        'check table value against greatest volume
        If ws.Cells(r, 12).Value = greatest_volume Then
        
            'assign greatest volume to table
            ws.Cells(4, 17).Value = greatest_volume
            
            'assign table ticker value
            t_ticker = ws.Cells(r, 9).Value
            
                'assign table ticker to table
                ws.Cells(4, 16).Value = t_ticker
            
         End If



    Next r

  
'autofit column width for the whole sheet
'syntax from Microsoft help article:
'https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit
ws.Range("A:Q").Columns.AutoFit

'next page of the worksheet
Next ws



End Sub




