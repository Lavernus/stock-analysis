# Analysis of Green Energy Stock

## Overview of Project

The client has created a workbook in Excel that contains information on the stocks of different green energy companies. He wants to use this information in order to help his parents choose which stocks to buy. I have previously completed an analysis on this dataset, but the client wishes to expand the dataset to include the entire stock market in order to diversify his parents' portfolio. 
### Purpose

Using VBA in order to automate the process, this analysis will refactor the code used in the previous analysis to be more efficient in order to handle larger datasets in a timely manner. It will display the total daily volume of each company as well as their returns over 2017 and 2018, which will give the client and his parents an idea of how each company is performing so they can decide which ones they would like to invest in.
## Results

In order to quickly analyze the data within the workbook, I utilized VBA to automate the process since the dataset was too large to work on manually.
### Setup
I kept the setup of the previous code the same, as it was simply establishing a table in a worksheet that would contain the information the client needed. The array used to keep track of the tickers in the previous codes was also kept.

Things began to change when I made the outputs of the code arrays instead of variables. I created three arrays to keep track of the outputs for each company, with "tickerVolumes" keeping track of total daily volume, "tickerStartingPrices" keeping track of the starting price, and "tickerEndingPrices" keeping track of the ending price. I also created a new variable called "tickerIndex" that would be the marker keeping track of which ticker we were analyzing at the moment, which I set to zero to begin with. After that,  I used a loop to define the tickerVolume to start at zero for every ticker we were analyzing. The code at this point looked like this:
```
Dim tickerIndex As Integer    
        tickerIndex = 0

Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single

For i = 0 To 11  
        tickerVolumes(i) = 0
Next i
```
### Analyzing the Stocks

Once I had all of the variables changed to arrays, I could use these to loop through the data and keep track of the values I needed. I created a loop that would run through the entire dataset, with the tickerIndex defining which ticker we would be analyzing during that loop. During each loop, the daily volume of the stocks matching the tickerIndex would be totaled and assigned to the totalVolumes of that tickerIndex. The starting price of the stock would then be assigned to the tickerStartingPrices of that loop's tickerIndex if the stock was the first that matched the tickerIndex of that loop. This was mirrored with the ending price, except the ending price would only be assigned to the tickerEndingPrices of that loop only if it was the last stock to match that loop's tickerIndex. Once the last stock to match that loop's tickerIndex was reached, the tickerIndex would increase by one and the loop would start again, now analyzing the ticker that matched the next tickerIndex. The code at this point looked like this:
```
For i = 2 To RowCount
    
    If Cells(i, 1).Value = tickers(tickerIndex) Then
            
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
    End If
        
    If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1) = tickers(tickerIndex) Then
            
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
    End If

    If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1) = tickers(tickerIndex) Then
            
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

        tickerIndex = tickerIndex + 1

    End If
    
Next i
```
### Displaying the Results

Once I had all of the arrays containing the data I needed, I could use them to output the results into the table I created during the setup. The code needed was similar to how I did it in the previous analysis, except it was outside the loop that was doing the actual analysis, which meant that it took less time to do the analysis since the output wasn't being repeated each loop. All I needed to do was replace the variables with the arrays I created and the tickers, total daily volumes, and returns were placed in the table. The code looked like this:

```
For i = 0 To 11
        
    Worksheets("All Stocks Analysis").Activate
    
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
Next i
```
The formatting of the table didn't need to be changed at all, and the scripts that measured the time it took to run the code I kept since I wanted to compare how much time it took to run the new code to the old code. 
### Analysis of the Results
#### Green Stocks Over 2017 and 2018

After running the code for both years, it returns two tables.

With the color coded "Return" column we can easily see that in 2017 the only stock to experience negative return was TERP, while other stocks returned a huge margin, with SEDG and DQ nearly tripling their returns and ENPH and FSLR over double theirs. We can see in 2018, however, that all stocks except ENPH and RUN experienced loss, with VSLR and TERP escaping the worst of it, having only lost 5% and 3.5% respectively. On the positive side, ENPH seems to be experiencing steady growth, and though it has dropped a bit in 2018, it is still a sizable amount at 81.9%. A safe bet may be to buy ENPH stock to bank on continued growth, or hop on FSLR stock while its in a dip and hope its past success could occur again.

Another interesting point to examine would be the volumes of stock each company is dealing in. Large volumes of stock may look impressive, but in practice means that any individual stock will react less to whatever gains or losses happen to the company. This could be good if you wanted to minimize risk, or bad if you wanted big returns on whatever you bought. FSLR, SEDG, and ENPH are dealing in some of the largest volumes across both years, while AY seems to be keeping a low volume of stocks. 

#### Effect of Refactoring

As mentioned in the background, I first started with a copy of the code that had the same utility as the one described in this analysis. The difference between the two is that the first edition, while having less scripts, took longer to run through the data since it would find the total volume, starting price, and ending price for a ticker and then output it into the table before looping back and moving on to the next ticker. This meant that, in practice, it was completing more steps than the new code since every time it found the values of the variables for the ticker it would output it into the table. The new code, however, keeps track of those values in arrays and outputs it into the table in its own loop that takes much less time at the end. This results in a much lower runtime, with the time elapsed going from these:



to these:


## Summary
Refactoring code offers a number of benefits. Improving the design of the overall code can increase the useability of it in the future by making it more adaptive to changes. It also gives you an opportunity to increase the readibility of the code which will make it easier to understand and easier to edit in the future. Mistakes and bugs that were glossed over the first time can also be found in a second or third read through.

Unfortunately, everything has its disadvantages. Ideally, you could keep refactoring code forever in order to make the most optimal version. Real life requires you to balance time and money, however, and refactoring eats into both when there are other projects to be completed. You can also introduce bugs that weren't present in the original code, creating a mess instead of an improvement.

This project benefited from the refactoring greatly; its redesign improved its logic so that it will be prepared for larger datasets in the future, and run faster. The readibility also benefited, since comments were added and whitespace was standardized across the code. The only drawback was that bugs that originally were not present were introduced, which meant that more time had to be taken in order to fix code that orignally worked.     
