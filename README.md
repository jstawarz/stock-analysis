# Stock Analysis

## Overview of Project

In this analysis, we are working to provide our client, Steve, with a refactored VBA macro that will help analyze stock data from the past few years.
We previously built a VBA macro that would analyze the stock data for 12 "green stocks" in an effort to identify which stocks performed well over the
2017 and 2018 years. However, we will now be expanding on that code in order for the macro to work on a larger dataset. In essence, we will be editing,
or refactoring the previously designed code to ensure that it will run efficiently on a larger set of stock data.

### Purpose

The purpose of this project will be to create a reliable way to analyze stock data using VBA. Using the previously engineered code, we will work to
refactor it, making it more efficient and ready for use with larger datasets. In this analysis, we will test to see if the refactoring did indeed
accomplish its mission of making the code more efficient. 

## Results

### Analysis of Stocks in 2017 and 2018

Based on the analysis, 2017 proved to be a profitable year for many green stocks. Of the 12 stocks analyzed, 11 provided net returns to their
sharholders, with DQ, SEDG, and ENPH leading the pack with 199.4%, 184.5% and 129.5% returns respectively. 

![This is an image](/jstawarz/stock-analysis/blob/main/Resources/AllStocks_2017.PNG)

In comparison, 2018 was a much more challenging year for green stocks with only 2 of the stocks yielding positive returns. They were RUN,
with a 84.0% return, and ENPH with a 81.9% return.

![This is an image](/jstawarz/stock-analysis/blob/main/Resources/AllStocks_2018.PNG)

### Refactored Code

With both years, we can demonstrate an overall success in refactoring the Stocks analysis macro. Originally, the code for the 2017 analysis ran
in 0.898 seconds, and in 0.883 seconds for the 2018 analysis. 

![This is an image](/jstawarz/stock-analysis/blob/main/Resources/Original_2017.PNG)
![This is an image](/jstawarz/stock-analysis/blob/main/Resources/Original_2018.PNG)

Post-refactoring, the macro for 2017 ran in 0.219 seconds and for 2018 ran in 0.172 seconds.

![This is an image](/jstawarz/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)
![This is an image](/jstawarz/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)


The main difference in the coding done for each analysis is that the refactored code utilizes a new variable, tickerIndex, which allows for
the code to match the three output arrays, tickerVolumes, tickerStartingPrices and tickerEndingPrices by the ticker symbol. Ultimately,
this saves us from having to perform a nested loop, and instead loop through the data one time. Overall, the refactored code was a success
with shorter run times for each.

![This is an image](/jstawarz/stock-analysis/blob/main/Resources/Refactored_Code.PNG)

## Summary

In general, the analysis shows one of the advantages of refactored code. It demonstrates that existing code can be recycled to create a more
efficient method for completing a similar task. That said, refactoring can be disadvantageous as well. Oftentimes, refactoring is being 
done on someone else's code. Depending on the complexity, code without notes, or written unclearly can be incredible difficult to refactor 
and may lead to enormous ammounts of time spent debugging. Even at the end of refactoring, you could potentially make previously functioning code,
unuseable. 

In this specific case, refactoring VBA code worked to run the code more efficiently. The specific subs can even be held on the same Excel sheet
so that comparison can be made between the two. It could be said that both codes worked well to summarize and provde the data needed, however
the original code may have been too slow to analyze the entire stock market like Steve ultimately wanted. 
