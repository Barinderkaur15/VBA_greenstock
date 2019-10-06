# VBA_greenstock
#challenge 

In order to refactor the code following steps were take 
1. Dim tickers() As String </BR>
   Dim startingprices() As Integer</BR>
   Dim endingprices() As Integer</BR>
 </BR>  
 The three arrays were defined without any defined elements , taking into consideration that number of tickers are unknown </br>
  2.To count the number of tickers a for loop is used in order to find the ticker value and then assign it to the array</BR>
   tickerindex = 0</BR>
    for = 2 To RowCount</BR>
 This condition will check the values in two consecutive rows and if the value is same it will assign the value of cell (ticker)   the respective array , this way we dont have to hard code array list in our code .It will dynamically check the value .</BR>
   If Cells(i, 1).value = Cells(i + 1, 1).value Then </BR></BR>
       tickers(tickerindex) = Cells(i, 1).value                                     
        End if</BR>
        If Cells(i, 1).value <> Cells(i + 1, 1).value Then </BR>
        tickerindex = tickerindex + 1 </BR>
          End If</BR>
 3. As we are not providing the limit of the arry to find the maximum and minimu index value lbound and ubound functions are used </BR>
 
Based on the analysis even though the return values over the year decrease for all the company, still RUN and ENPH are able to give positive revenue in both years we are considering.
