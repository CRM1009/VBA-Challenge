Sub LoopWorksheet1()
  
        
    'Declare worksheet as an object variable
    Dim ws As Worksheet
    
' Looping process to run through all the worksheets
For Each ws In Worksheets

 
 ' Set an initial variable for holding the yearly change,
 ' percent change, total stock Volume, and summary table
 
    Dim Year_Open As Double
    Dim Year_Close As Double
    Dim Summary_Table_Row As Integer
        
                
    ' Add the word ticker as column header
    ws.Cells(1, 9).Value = "Ticker Symbol"
                
    ' Add the word Yearly Change as column header
    ws.Cells(1, 10).Value = "Yearly Change"
                
    ' Add the word Percent Change as column header
    ws.Cells(1, 11).Value = "Percent Change"
                
    ' Add the word Total Stock Volume as column header
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
        
    ' Keep track of the location for each ticker in the summary table
    Summary_Table_Row = 2

    
    ' Last_Row
    LR = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
          
   For i = 2 To LR
   
   ' Loop to find first value in open column after ticker change
            If ws.Cells(i + 1, 1) = ws.Cells(i, 1) And ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                Year_Open = ws.Cells(i, 3).Value
                
            End If
    
                       
        ' Check if we are still within the same ticker, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        
        ' Set the Ticker
        Ticker_Name = ws.Cells(i, 1).Value
        
         
        ' Calculates Yearly Change
                
        Year_Close = ws.Cells(i, 6).Value
        
        Yearly_Change = Yearly_Change + Year_Close - Year_Open
        
        
        ' Add to the Total Stock Volume
        Volume_Total = Volume_Total + ws.Cells(i, 7).Value
                
                       
        ' Calculate Percent Change and format as percent
        Percent_Change = Percent_Change + (Year_Close - Year_Open) / Year_Open
       
                                           

        ' Print the Ticker Name in the Summary Table
        ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                
        ' Print the Yearly Change
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
        ' Print the Percent Change
        ws.Range("K" & Summary_Table_Row).Value = Percent_Change
        
        ' Format column K as percentage
        ws.Range("K2:K" & LR).NumberFormat = "0.00%"
                
        
        ' Print the Total Stock Volume to the Summary Table
        ws.Range("L" & Summary_Table_Row).Value = Volume_Total
        
        ' Test what value Year_Open is producing
        ' ws.Range("N" & Summary_Table_Row).Value = Year_Open
        
        ' Test what value Year_Close is producing
        ' ws.Range("M" & Summary_Table_Row).Value = Year_Close
         
        
        ' Add one to the Summary Table Row
        Summary_Table_Row = Summary_Table_Row + 1
        
                          
        
        ' Reset the Yearly Change Total
        Yearly_Change = 0
        
        ' Reset the Total Stock Volume
        Volume_Total = 0
         
        ' Reset the Percent Change Total
        Percent_Change = 0
      
                     

            ' If the cell immediately following a row is the same brand...
            Else
            
                      
            ' Add to the Volume Total
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value
                                  
         
            
        End If
       
   Next i
   
   
   ' Loop to format change in color red if negative and green if positive
    J_EndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row


            For j = 2 To J_EndRow

            ' If greater than 0 format color Green, else color RED
                If ws.Cells(j, 10) > 0 Then

                        ws.Cells(j, 10).Interior.ColorIndex = 4

                    Else

                        ws.Cells(j, 10).Interior.ColorIndex = 3
                        
                End If

            Next j

    
    ' Create labels for second summary table
    ' Print word Greatest % Increase in cell N2
    ws.Range("N2").Value = "Greatest % Increase"
    
    ' Print word Greatest % Decrease in cell N3
    ws.Range("N3").Value = "Greatest % Decrease"
    
    ' Print word Greatest Total Volume in cell N4
    ws.Range("N4").Value = "Greatest Total Volume"
    
    ' Print word Ticker in cell O1
    ws.Range("O1").Value = "Ticker"
    
    ' Print word Value in cell P1
    ws.Range("P1").Value = "Value"
    
    
    ' Go to the last row of summary column k
    kEndRow = ws.Cells(Rows.Count, "K").End(xlUp).Row

    ' Define variable to initiate the second summary table value
    Increase = 0
    Decrease = 0
    Greatest = 0
 
        ' Find max/min for percentage change and the max volume Loop
        For k = 3 To kEndRow

            ' Define previous increment to check
            last_k = k - 1

            ' Define current row for percentage
            current_k = ws.Cells(k, 11).Value

            ' Define Previous row for percentage
            prevous_k = ws.Cells(last_k, 11).Value

            ' Greatest total volume row
            volume = ws.Cells(k, 12).Value

            ' Prevous greatest volume row
            prevous_vol = ws.Cells(last_k, 12).Value

   

            ' Find the increase
            If Increase > current_k And Increase > prevous_k Then

                Increase = Increase


            ElseIf current_k > Increase And current_k > prevous_k Then

                Increase = current_k

                ' Define name for increase percentage
                increase_name = ws.Cells(k, 9).Value

            ElseIf prevous_k > Increase And prevous_k > current_k Then

                Increase = prevous_k

                ' Define name for increase percentage
                increase_name = ws.Cells(last_k, 9).Value

            End If

      
            ' Find the decrease
            If Decrease < current_k And Decrease < prevous_k Then

                ' Define decrease as decrease
                Decrease = Decrease

               
            ElseIf current_k < Increase And current_k < prevous_k Then

                Decrease = current_k

                ' Define name for decrease percentage
                decrease_name = ws.Cells(k, 9).Value

            ElseIf prevous_k < Increase And prevous_k < current_k Then

                Decrease = prevous_k
                
                ' Define name for decrease percentage
                decrease_name = ws.Cells(last_k, 9).Value

            End If

      
           ' Find the greatest volume
            If Greatest > volume And Greatest > prevous_vol Then

                Greatest = Greatest

               
            ElseIf volume > Greatest And volume > prevous_vol Then

                Greatest = volume

                ' Define name for greatest volume
                greatest_name = ws.Cells(k, 9).Value

            ElseIf prevous_vol > Greatest And prevous_vol > volume Then

                Greatest = prevous_vol

                ' Define name for greatest volume
                greatest_name = ws.Cells(last_k, 9).Value

            End If

        Next k
        
        
            ' Print Greatest Increase Name
            ws.Range("O2").Value = increase_name
               
            ' Print Geatest Decrease Name
            ws.Range("O3").Value = decrease_name
            
            ' Print Geatest Volume Name
            ws.Range("O4").Value = greatest_name
            
            ' Print Greatest Increase Value
            ws.Range("P2").Value = Increase
            
            ' Format Greatest Increase as percentage
            ws.Range("P2").NumberFormat = "0.00%"
                        
            ' Print Greatest Decrease Value
            ws.Range("P3").Value = Decrease
            
            ' Format Greatest Decrease as percentage
            ws.Range("P3").NumberFormat = "0.00%"
            
            ' Print Greatest Volume Value
            ws.Range("P4").Value = Greatest
    
    
            ws.Columns("I:Q").AutoFit

    
Next

  
End Sub
