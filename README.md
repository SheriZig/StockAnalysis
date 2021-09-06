# Overview of Project:
Steve wants to provide his parents with a tool to evaluate various stocks. They believe that trading volume is an accurate indicator of the value of a stock. 
Utilizing stock data for years 2017 and 2018, a summarized data table that includes the stock ticker, Total Daily Volume and returns based on the starting and ending price of the stock is generated with VBA (Visual Basic Application) within Microsoft Excel. 

# Results: 
After summarizing the data into a table with a tabular view, the stocks were evaluated to compare stock performance for the two years. The code was also refactored from its original form to provide for improved runtime. 

## Analysis of Stock Performance

![image](https://user-images.githubusercontent.com/88912539/132240522-ac7e8678-2038-471f-b1d7-22bbbbb7299c.png)                 
![image](https://user-images.githubusercontent.com/88912539/132240639-5d3842a4-7c9c-4ec8-aca1-1f7d979c467e.png)







Overall, stocks performed better in 2017 than in 2018 with 92% of the stocks showing a positive return in 2017. In 2018, only 17% of the stocks closed the year with a positive return. Analyzing the Total Daily Volume for the 12 stocks revealed that 7 of the stocks had a higher return and higher volume. Five of the stocks had a higher return but lower volume when comparing the two years. When analyzing stock ticker “DQ” notice that the Total Daily Volume was significantly lower in 2017 (35,796,200 vs 107,873,900) but the returns were significantly higher (199.4% vs -62.6%).  There are many factors that determine the price of stocks. Trading volume is not a consistent measure of performance and should not be the only factor  considered when making stock purchase decisions. 


## Analysis of Refactored Code Performance 
Original Code 2017

![image](https://user-images.githubusercontent.com/88912539/132240984-cfd38723-dafd-4db4-bdda-d7687774a1ad.png) 


Refactored Code 2017
![image](https://user-images.githubusercontent.com/88912539/132241158-e2370ee3-b9ba-4916-8c04-e30c2ebd3f72.png)

Original Code 2018
![image](https://user-images.githubusercontent.com/88912539/132241117-5b320a97-4599-44d5-a1d0-0681cd2727ec.png) 

Refactored Code 2018
![image](https://user-images.githubusercontent.com/88912539/132241207-6e2ba630-5156-4daf-94e1-cee2c97d19b5.png)

FOR Loop Refactored Code 
````````````````
For i = 0 To 11
    
       tickerVolumes(i) = 0
    
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
   
    '3a) Increase volume for current ticker
        
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        
     '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1) <> tickers(tickerIndex) And Cells(i, 1) = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
        End If
        
     '3c) check if the current row is the last row with the selected ticker
      
        If Cells(i + 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
     
     '3d Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
         tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
    Worksheets("All Stocks Analysis").Activate
        
    Cells(i + 4, 1) = tickers(i)
    Cells(i + 4, 2) = tickerVolumes(i)
    Cells(i + 4, 3) = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i

```````````````

The images above represent the performance improvements that were realized with the refactored code. The original code contained an outer FOR loop with a Nested loop. The refactored code eliminated the Nested loop and consolidated to one significant loop. Using the index instead of the ticker also enhanced performance. Finally, the output was provided as a separate loop instead of being part of the outer FOR loop as was the case with the original code.  
While the increased performance was only 1 second, the dataset only contained 3013 rows of data. The impact of the performance improvement will be fully realized when working with datasets of 1 million + rows of data. 

###### Advantages/disadvantages of refactoring code: 
1.	One advantage of refactoring code is to increase performance. As demonstrated with the refactoring the 2017 and 2018 stock data a small change in the code will yield demonstrable differences in performance. 
2.	An additional advantage will be found in readability. Refactoring may make the code easier for other developers to read and easily distill the intent of the code.  
3.	A disadvantage to refactoring code in a production environment, is the cost in terms of time and effort. The benefit of the refactoring must outweigh the cost. A careful evaluation should be conducted to determine if refactoring is advantageous.  
4.	A second disadvantage could be unintended consequences. Refactoring code may cause an unintended break in another part of the code introducing a bug. The may be unrealized at the time of refactoring and previously tested code may now contain a bug that will be sent to customers in the field. 
There are advantages and disadvantages to refactoring existing code. Careful consideration should be given to the scope of the change and possible consequences to making meaningful changes. 
In the case of the original VBA code that was refactored, it was time consuming to change the structure of the code to remove the nested FOR loop based on the provided dataset the improved performance was not significant. However, if the intent is to run the code on a dataset containing thousands or millions of rows of data any improvement to the performance will be exponential. 


