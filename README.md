# Stock Snalysis: VBA of Wall Street


## Overview and Purpose of Project

I have been helping my friend Steve analyze stock market data for his parents as they’ve jumped into the world of investing. Steve feels his folks need some guidance based on solid data, so we’ve prepared data sets in Excel using VBA to summarize a dozen stocks’ performances over the past year and the year before that. I’ve been able to set up the worksheet to be completely user friendly for Steve, so that with one click of a button he can begin an analysis and choose the year he wishes to view. With another click of a button, he can clear out the current analysis and then start another if he wishes. 

Now, to truly guide his parents in the right direction, Steve would like to build upon our success and set up the worksheet to handle the *entire* stock market’s data over the last few years. This expansion of our data set requires me to refactor my initial code so that it runs quicker and more efficiently. I will attempt to do so, and then test my code on our initial data set of a dozen stocks to look for speed improvements. 


### Project Results 

The entire process of refactoring the stock analysis code was both fascinating and frustrating. As I’m writing this, I’m looking at both code sets side-by-side, and it’s so interesting at how each set is both so similar and so different. 

In our initial code set, we created a For Loop to initialize over the “tickers” array, and then created a Nested For Loop within that to initialize (and where applicable, aggregate) volume, starting price, and ending price for each stock. The Nested loop repeated once down the rows in the spreadsheet for the current stock in the “tickers” array, then our original Loop outputted that stock’s results/outputs in our output worksheet, and finally, the For Loop began again for the next stock in the ”tickers” array until all the stocks in the array were analyzed. 

In our refactored code set, we still created our “tickers” array, but then created an Index variable to follow along with it, as well as to follow along with three other arrays we created for volumes, starting prices, and ending prices. ![](/Resources/TI_Array.png) 

Then we created three independent For Loops. 

**Number One** set all of the ticker volume array values as zero.
![](/Resources/Vol_Loop.png)

**Number Two** ran our analysis for every stock, like our Nested For Loop did in the initial code set. This chunk of code, however, used our Index variable as its index instead of – I’m not sure how one should term it – a “regular” index, like “i” or “j” as its index. 

![](/Resources/FL_Vol_SP.png)

As shown, this Loop set values for every stock’s volume, starting price, and ending price. Furthermore, it also included code to increase the Index variable once the stock being analyzed no longer appeared in the data worksheet. 

![](/Resources/EP_TickerI_EndL.png)

This code is included because unlike our initial code, all of our volume and price data values are now in arrays and are now indexed just the same as our ticker data and need to be progressed as well. 

**Number Three** served to output our data. This part of the code was where my frustration and inexperience set in, because for a long time, I expected to use our Index variable as our index to output our data from our tickers, volume, and prices arrays, not comprehending that the value of that Index variable had already been aggregated by my code (in the "Number Two" For Loop) to a value beyond where it held any data – and NOT back to zero, where I assumed it was sitting. I spent *hours* moving code, recoding code, rewriting code, begging for results from the code, asking why I was indexing code that already had indexes! It wasn’t until I spent a half hour in class office hours with Alex, one of our TAs, that we went over the code line-by-line, determined everything worked until the results, and the light bulb lit up to use a “regular” index to output the results. Thank you, Alex!!

![](/Resources/Output.png)


In the end, I was quite impressed with how much the code refactor sped up the analysis. I really didn’t think that one For Loop and its Nested For Loop would take twice as fast to run as three For Loops using three times the arrays with an index tracking them all.

![](/Resources/VBA_Challenge_2017.png)

![](/Resources/VBA_Challenge_2018.png)


### Project Summary


In general, I’ve found so far that refactoring code is incredibly beneficial for the efficiency of a program. I imagine at immense scale the productivity and efficiency of a code base makes or breaks a piece of software in the business world. Why would a person purchase a competitor’s software if it does exactly the same process but takes considerably longer? I am a little hard-pressed to think of many disadvantages. Perhaps there would be more levels of abstraction in the code as more Indexes and variables are created (as well as other objects I don't even know about yet!). For future developers, the code would only be as good as its documentation. This statement is likely true for any code, but I would think it would only become more pronounced as a code is trimmed down, refactored, restructured, etc. 

To that point, with my own beginner's level of experience, I actually prefer the original VBA script we put together. I can follow what’s going on there much easier compared to the refactored script. The advantage of the original VBA script, and I suppose with first draft scripts in general, is that it’s simplified and easy to understand. Its disadvantage is that it’s not efficient. The refactored VBA script is more abstract, more complicated, and requires a greater proficiency to translate – that’s the disadvantage to the human developer. Its far greater advantage, however, is that its efficiency can’t be beat. 


