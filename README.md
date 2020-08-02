# Stock Analysis

## Overview of Project

Previously, we assisted a friend with creating a macro that pulls in data related to the total volume and return from specific stock data sets. 
The macro we previously built worked great for small data sets, but the scope of the project expanded to include significantly larger data sets and we wanted to refactor our code to ensure performance.

The analysis is the file below:

[VBA_Challenge.xlsm] (https://github.com/pritchardjeff/Stock_Analysis/blob/master/Resources/VBA_Challenge.xlsm)

### Results

At a high level, we reduced the number of sheet changes in our code from 2 per stock ticker to 2 total, resulting in a change of performance from approximately .80 seconds to .17 seconds for each year.

See run times in the next section to show actual changes.

#### Run Times

##### Original Code Run times

![VBA_Challenge_2017_Original.png] (https://github.com/pritchardjeff/Stock_Analysis/blob/master/Resources/VBA_Challenge_2017_Original.png)

![VBA_Challenge_2018_Original.png] (https://github.com/pritchardjeff/Stock_Analysis/blob/master/Resources/VBA_Challenge_2018_Original.png)


##### Refactored Code Run times

![VBA_Challenge_2017.png] (https://github.com/pritchardjeff/Stock_Analysis/blob/master/Resources/VBA_Challenge_2017.png)

![VBA_Challenge_2018.png] (https://github.com/pritchardjeff/Stock_Analysis/blob/master/Resources/VBA_Challenge_2018.png)

## Detail of efficiencies 

1. Created a ticker index to reduce the amount of data stored for the purpose of comparison. 
	- It is more efficient to compare a 1 to 1 than "SPWR" to "SPWR."

2. Created arrays for output dimensions for totalVolume, startingPrice, and endingPrice
	- Using an array for an output dimension allowed us to calculate and store all our data points in the same loop.
	- The original code forced us to switch between the Data and Output sheets for each individual ticker for a total of 24 sheet changes within the macro. 
	- The refactored code calculates and stores all our outputs in one loop and then populates the data in the second loop for a total of 2 sheet changes within the macro.

### Code Detail with Annotations

![VBA_Challenge_Original_Code.png] (https://github.com/pritchardjeff/Stock_Analysis/blob/master/Resources/VBA_Challenge_Original_Code.png)

![VBA_Challenge_Refactored_Code.png] (https://github.com/pritchardjeff/Stock_Analysis/blob/master/Resources/VBA_Challenge_Refactored_Code.png)


## Summary

### Pros and Cons of Refactoring Code

Often, when code is originally written the developer is trying to figure out how to get the code to work as intended and writes with the pretext of refactoring the code after it is originally written.

#### Pros

Refactoring leads to code that is more efficient, easier to understand, and more maintainable.

#### Cons

Refactoring may take more time than anticipated and may not result in performance that is more noticeable to users. 

The biggest concern with refactoring should be around the opportunity cost related to the developer’s time. 

At the end of the day, one must think through the value of the potentially more efficient refactored code vs the addition of new features.

### Pros and Cons of the original vs refactored VBA script

#### Pros

Faster run time by eliminating the need for the code to switch between the output and input sheets for each individual ticker.

#### Cons

The code has exactly the same output as it had previously and if a larger data set with more tickers was added, we would still need to adjust the code because the tickers in our array are all hardcoded and not dynamic.

### Closing statement

We greatly decreased the amount of time needed to run the code on our current data set, but our code will still need to be manually adjusted whenever the underlying tickers are updated. Therefore we are one step closer to what our end user needs, but will still need to make more adjustments before we have a finished product.




