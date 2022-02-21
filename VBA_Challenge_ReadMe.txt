# Analysis of Stock Data with VBA

## Overview of Project
An analysis of stock performance for 12 different Green Energy stocks using VBA.

### Purpose
The purpose of this project was to analyze the performance of 12 different Green energy companies to determine how one company (ticker name DAQO) compared to its peers. If DAQO performed poorly, the next task was to identify other companies from the target list for investment. To gauge company performance stock data from 2017 and 2018 was used. The final goal was to refactor the code to increase performance to potentially be used with larger datasets.

## Results
Initial analysis of DAQO performance revealed a loss of 62.6% in 2018, signaling it was not a good investment. Following this, I wrote an initial set of code to analyze the total volume of stocks traded and the percent return on those trades for 12 different Green Energy stocks in 2017 and 2018. Of the stocks analyzed, 'ENPH' and 'RUN' netted the highest consecutive returns, highlighting them as the best overall investment. This initial analysis took ~0.723s for 2017 and ~0.754s for 2018, "s" standing for seconds. Following this I refactored the code to make it more efficient, with the end result being ~6x faster. The process of refactoring is detailed in the below 'Analysis' section'

### Analysis
I began the refactoring process by creating a roadmap of the steps I would need to take for the code to function as intended. I then copied over the basic code that likely wouldn't need to be changed, being the headers, input box, ticker array and worksheet activation before placing them in their appropriate positions. This done, I reviewed the code again and determined the most optimal way to increase effeciency would be to reduce the number of nested for loops in the code. To do this, I created a new variable to hold the ticker array, and then created a series of arrays to hold ticker volume, starting price and ending price before setting the value of all these fields to 0. I then used these fields to write another for loop that would be able to read the data all at once before filling in the required fields, rather than having to loop through the data for each new stock price. 

### Analysis of Outcomes Based on Goals

### Challenges and Difficulties Encountered

## Results
The net result was a nearly 6x increase in speed, with analysis time decreasing from 0.723s to 0.109s and 0.754s to 0.125s for 2017 and 2018 respectfully (as shown by the pictures below).

###Pros and Cons of Refactoring
The biggest advantage for refactoring code is in its efficiency. The first and most obvious efficiency increase is through code run-rate, with succesfully refactored code running faster and (ideally) taking up less resources to complete. The other upside is efficiency in understanding. A cleaner code with less lines and (ideally) more explanation of steps makes it much more palatable for others to read, understand and use. This means that succesfully refactored code can more easily be taken and used as a model for projects outside of its initial function, whereas this may not be possible for long and messy code that has never been refactored.

The downside of refactoring is the cost of human time. While making the code simple, fluid and easy to digest may be ideal, it may not be feasible to do when facing strict deadlines. If the emphasis is placed on getting the code to work as fast as possible, irregardless of how it looks, refactoring could be a luxury a coder isn't able to enjoy. 

###Pros and Cons of Refactoring for this Project
The ~6x decrease in runtime for the code is a clear benefit. It is also easier to follow what the code is doing compared to the previous iteration.

The downside may be that the code has become too specialized for this particular task, and may be harder to utilize for other projects.
