# VBA Stock Market Analysis

## Overview of Project
In this analysis we were helping Steve to analyze how different stocks performed. We refactoring the code to loop through all of the data one time in order to collect all the data. We were then comparing whether refactoring decreased the run speed.

## Results
### The Refactored Code
#### Step 1
Step one required refactoring the code to include a tickerIndex which would be used to access the correct index across the four arrays. Then the output arrays needed to be created.
#### Step 2
Step 2 was to create a for loop to initialze the tickerVolumes to zero and then loop through all the rows in the spreadsheet.
#### Step 3
Step 3 created a function to add the ticker volume to for the current stock ticker. It then had functions to determine if the row was the first and last row of the ticker so that it could create the starting and ending prices.
#### Step 4
Finally step four created the output arrays for the ticker, volume and the percent return.

### Comparison of Stock Performance from 2017 to 2018
2017 was a great year for the stocks. Only TERP had a negative return. Four stocks (DQ, ENPH, FSLR, SEDG) all had over 100% return on investment for the year.
2018, however, was not a great stock year, with only ENPH and RUN having postive returns. 

With this in mind the best stock to invest in would be RUN as it had a positive return in both years and increased from a 5.5% return to an 84% return between 2017 and 2018.

### Execution Times
The run time for the 2017 analysis was improved by .10 seconds

The run time for the 2018 analysis was improved by .03 seconds

## Summary

### Advantages
The most important advantage of refractoring this code is that it helped the program run faster. When running multiple datasets the faster the code runs the more efficient one can be in their analysis. Refactoring code can also help in debugging as it can elimate problems related to multiple coping and pasting of for loops that could cause errors. Refractoring can also make code cleaner and easier to understand for someone who did not develop code or if someone is returning to the code months after developinging it.

### Disadvantages
One disadvantage is that refactoring the code may take a significant amount of time when the original code was already working properly. Refactoring the code runs the risk of introducing new bugs. Both of this could make refactoring expensive and risky in the eyes of a management team.

### Pros and Cons Applied to the Original VBA Script
The original VBA script worked consistently and in my case refactoring caused me to create a new bug that had the return being divided by zero. This was because the new naming structure introduced caused me to write tickerStartingPrice instead of tickerStartingPrices. In this instance it took some time to determine the bug since the phrases were so similar. 
The pro of refactoring the orginial code was that it did run faster which would be helpful if more years were introduced.
