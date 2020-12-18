# **Stock Analysis**
## Overview of Project
  For individuals seeking to purchase stocks in a particular company, it is important to first look at the data surrounding these stocks and their historical performance.  We want to make sure any advice provided to individuals seeking to buy stock is sound and based on strong data points.  We set out to review the past performance of a dozen different stocks in order to see which stocks have been providing positive returns.  After the initial pool of stocks was identified, the scoop of the project broadened and we realized the necessity of developing a code that could deliver the same level of analysis that was performed on a dozen stocks, to the entire stock index without tying up all of our computers resources.  Since we already had the VBA script to run the analysis for the dozen chosen stocks, refractoring was necessary in order to run this same script over a larger array of stocks. 
  ## Procedures Performed
    After initially downloading the data into Excel, we set out to determine the key variables to be analyzed in order to demonstrate a stocks performance.  One of the first variables we honed in on was the daily volume of stock trades.  The reason for starting with daily volume, is to help build confidence in the true value of a stock.  The more a stock was traded, the more accurate the price of that stock would be.  In order to add and store this volume data for each index, we created a conditional statement to sum the daily volume for all instances of a particular stock index.  The code below was used, after first initializing both the tickerIndex and the tickerVolumes to zero: 
            'If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            End If'
    
    This code to calculate volume was built as part of a 'for' loop, that would run through all rows of data, while increasing the *tickerIndex* by one once it reached the last row of data for a particular ticker symbol.  We accomplished this by using another conditional code 
    'If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then 
   tickerIndex = tickerIndex + 1
   End If'
   The subsequent stored volume data was then projected onto a summary table on the tab "All Stocks Analysis" contained within our analysis [VBA_Challenge](https://github.com/kroman3105/stock-analysis/blob/master/VBA_Challenge.xlsm).  We performed a similar analysis to identify the return on each stock, and added some conditional color fill formatting to clearly show those stocks that had positive returns.
  'If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If'
        
 ## Refactoring the Code
  The original code developed to analyze this data performed a nested loop function that ran over the 12 tickers identified.  While this approach was effective and returned results in roughly 0.75 seconds per year analyzed, there were concerns about the amount of time it would take to run this analysis for every possible stock index.  This is why we took on the exercise of refactoring the code to bring the processing time of running this analysis down, which would in turn make it a useful code even when increasing the dataset.  In order to achieve this, instead of creating a nested loop to run capture the volumes, starting prices and ending prices, an array was implemented.  Using an array as opposed to a nested loop, brought the processing time down a full half second per year as evidenced by the runtimes captured for [2017](https://github.com/kroman3105/stock-analysis/blob/master/Resources/VBA_Challenge_2017.PNG) and [2018](https://github.com/kroman3105/stock-analysis/blob/master/Resources/VBA_Challenge_2018.PNG).  
  
## Insights Gained from the Data
  While 2017 was clearly a good year for all stocks analyzed (other than the negative return delivered by TERP), it is important to not look at a single year of returns in a vacuum.  That is why our code was built with the ability to run the analysis over several years by prompting the user with a message box. 
  'yearValue = InputBox("What year would you like to run the analysis on?")'
  When looking at the combined 2017 and 2018 returns, we can gain a little more insight on how consistently these stocks gained positive or negative returns.  In reviewing the year over year data, we see that only ENPH and RUN returned a positive return for each year, so based on the data available we would be inclined to advise potential investors to consider these stocks given their consistent positive returns.  However, it should be noted that only having two years of data is a small sample size to choose for, so further analysis would be required that would include more years of data.  It also should be pointed out that in the stock market, past success is no guarantee for future returns.
  
  ##Summary
    As the amount of data made available to us continues to grow, it is important to review our code to see if refactoring is necessary in order to improve its efficiency as it ingests more data.  The advantages of refactoring code are clear from the run time improvements we saw on our analysis.  Refactoring code can make it run more efficiently, allowing that same code to be applied to a larger dataset.  One of the disadvantages discovered in refactoring code is the greater potential for error that exists when not factoring in all of the implications to a refactored line to the rest of the code.  Changing from a nested loop to an array has implications on almost every line of code, and it's easy to miss some of the minor notations changes that need to be implemented when code is refactored.  While refactoring of the original VBA script clearly brought a more efficient run time, one of the downsides of the refactoring is the additional code needed on the front end to effectively establish and zero the opening arrays.  I could see this being a step mishandled by new coders, while the script for the VBA prior to the refactoring was a little more intuitive.  Overall, it's hard to argue with the gains in efficiency, so I think this particular VBA script needed to be refactored in order to achieve this.       
   
