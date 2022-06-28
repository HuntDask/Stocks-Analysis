# Stocks-Analysis
## Overview of Project
For our analysis project, we examined various stocks for Steve to help his parents in their stock market success. In our project, we created a VBA code for both stocks in 2017 and 2018 to search for trends in 12 distinct stocks we looked at over the course of a year period. This will lead to a better understanding of the trends occuring so that Steve can give his parents the best advice and adivsing for their stock market purchases. Our purpose of the project was to create and reorganize a table of stocks over the 2017-2018 period looking at the stocks total daily volume along with their returns to help Steve reach his desired end goals for his parents. Looking at these different characteristics of these stocks from a year to year basis is the most efficient way to understand and analyze the stocks for trends accordingly.
## Results
As we look at the two tables showing the trends of the stocks in 2017 to 2018, we can see that most of these stocks took a dramtic downfall in return during the one year period. During 2017, 11 out of 12 stocks had a positive rate of return. As we look at 2018, only 2 out of 12 had a positive rate of return both years. The only stocks that were able to have a positive rate of return during the years examined were ENPH and RUN, which would be the best stocks for Steve and his parents to invest in. From our VBA coding, we were able to refactor the code to make our execution of data analysis more efficient. The refactored script was much more efficient than the nested loop we used earlier in the analysis previously before the challenge. The refactored code using Volumes/Starting and Ending prices made the coding more efficient than the pervious loop in our original version.

![All Stocks (2017)](https://user-images.githubusercontent.com/107444840/175656162-cc30e7c5-be18-4ffa-b5b6-8567dc8ce360.png)
![All Stocks (2018)](https://user-images.githubusercontent.com/107444840/175656220-7de64d87-4390-4e7e-8a09-d56753a85327.png)
## Summary
### Advantages/Disadvantages of refactoring code
For refactoring our code in Excel VBA, the biggest advantage is that is makes the orignial code much more efficient than before. It is important when you are rapidly trying to analyze stocks and data quickly in order to discover future trends in the stock market so you can capitialize. As for a disadvantages, the biggest disadvtange for refacotring is that you could entirely mess up your code if you refacotr incorrectly. If you do not save your previous code and refactor incorrectly, it could potentially ruin your whole coding further messing up your stock analysis and tables in excel. 
### Pros/Cons of refactoring original VBA script
The biggest advantage for being able to refactor the original VBA script is  you can use the original data you find and build off of it in a different module while keeping your same pervious data from before. It is important to look over your original coding so that any chagnes you make you can match and make sure they line up correctly to be as efficient as possible. A disadvantage of this refactoring however is that you are not able to fully understand the syntax of the new refacoring code comapred to your pervious data you already collected. It is hard to have a real understanding of the syntax in the refactored module because the origional depicts all the neccessarry functions and methods that are already correct, not the refactored module.

## VBA Analysis Challenge
[VBA Analysis .zip](https://github.com/HuntDask/Stocks-Analysis/files/8981604/VBA.Analysis.zip)

[VBA Challenge.xlsm.zip](https://github.com/HuntDask/Stocks-Analysis/files/8981611/VBA.Challenge.xlsm.zip)

[VBA Resources.zip](https://github.com/HuntDask/Stocks-Analysis/files/8981605/VBA.Resources.zip)


## Refactored Code For Assignments Upload Issues
'3) Initialize array of all tickers
Dim tickers(12) As String

tickers(0) = "AY"
tickers(1) = "CSIQ"
tickers(2) = "DQ"
tickers(3) = "ENPH"
tickers(4) = "FSLR"
tickers(5) = "HASI"
tickers(6) = "JKS"
tickers(7) = "RUN"
tickers(8) = "SEDG"
tickers(9) = "SPWR"
tickers(10) = "TERP"
tickers(11) = "VSLR"



'4a) Activate data worksheet
Worksheets(yearValue).Activate

'4b) Get the number of rows to loop over
RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'5a) Create a ticker Index

Dim tickerIndex As Single
tickerIndex = 0

'5b) Create three output arrays

ReDim tickerVolumes(12) As Long
ReDim tickerStartingPrices(12) As Single
ReDim tickerEndingPrices(12) As Single

'6a) Initialize ticker volumes to zero
    
For i = 0 To 11
tickerVolumes(i) = 0

Next i
'6b) loop over all the rows

For i = 2 To RowCount

7a) Increase volume for current ticker
   
tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 9).Value
    
'7b) Check if the current row is the first row with the selected tickerIndex.
If Cells(i - 1, 2).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 7).Value
        
        
End If
    
'7c) check if the current row is the last row with the selected ticker
If Cells(i + 1, 2).Value <> tickers(tickerIndex) Then
tickerEndingPrices(tickerIndex) = Cells(i, 7).Value
        

'7d Increase the tickerIndex.
tickerIndex = tickerIndex + 1
        
End If

Next i

'8) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
For i = 0 To 11
    
Worksheets("All Stocks Analysis").Activate
tickerIndex = i
Cells(i + 4, 1).Value = tickers(tickerIndex)
Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
    
Next i


