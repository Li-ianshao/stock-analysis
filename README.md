# stock-analysis
## Overview of Project
This project is to help Steve to analyze the green energy stocks for his parents investment.
## Results
The difference between two of the analysis is that the refactored one spend less time to calculate daily volume and yearly return. 

Because the refactored script doesn't use the nested for loop, so it load all 3012 rows of trading data for calculating every tickers' daily volume only once for all, on the other hand, the original script use nested for loop whitch make VBA load data every time when ticker change.

That is the main reason, the refactored script spend much more less time then the other one.

## Summary
### What are the advantages or disadvantages of refactoring code?
The advantage of refactoring code is that it spend less time, and the code is more clear then nested for loop, and the disadvantage is the for loop that initialize the tickerVolumes to zero, I think this for loop is not necessary.

### How do these pros and cons apply to refactoring the original VBA script?
