 ### **Stock Data Analysis with Microsoft VBA**

To perform this analysis, we have:

'Stock_Data.xlsm' which includes:

* Datasets of Stocks for years 2017,2018
* Output sheet 'All Stocks Analysis'
* VBA Script 
* Executed through action buttons 'Run Analysis for all Stocks' and 'Clear Worksheet'. 

~ It also includes 'Resources' folder for analysis requirements.

# **Overview of Project**

## Background
Client would like to invest in green energy stocks. They are specifically interested in DQ stocks. But an optimal analysis should be able to provide the seeker, which stocks look good to invest in, based upon analysis of given stocks over year 2017 and 2018. Also, we can deduce performance of DQ stock.

To do this analysis, we have the following as briefed below.

There is a dataset for year 2017 and 2018, presented in the form of rows and columns in excel Worksheets named '2017' and '2018' in Workbook named 'Stock_Data'. In these worksheets, we have date-wise trading information in the stock market for various stocks. These stocks here are represented under 'Ticker' column A of worksheets.Then we have different columns which provide Open Market Value of the Ticker under Column 'Open', Highest value of stock for the day under Column 'High', Lowest value of the stock for the day under 'Low' column, Closing value of the stock for the day under 'Close' and so on, along with Total Volume traded on the date under 'Volume' column.

### Current Stage of Analysis:
We already have a VBA script that helps us to analyze a dozen of stocks to find out 'Total Yearly Volume' of the stock and 'Yearly return'. We are able to choose, for which year, out of '2017' and '2018', we want to see these 2 values for, for all the tickers (stocks) in the available dataset, through action button 'Run analysis for all stocks" and 'Clear Worksheet' to clear our Analysis worksheet. 

## Purpose of this Analysis:
We need to refactor already provided starter VBA script which should enable us to perform better than the original script. Here we measure performance in terms of execution time. 

### Expected Outcome:
We should be able to present through comparison, between original and refactored script, which script performs better through 'Execution Times'. 

Also, how the stocks performed over these two years. Thirdly, we should be able to present how refactored script is better than original script and what aspects are at disadvantageous side of this refactoring. 

## Results:

After refactoring of the code, we have the following results, as explained from the following two aspects:

1. Comparison of stocks' performances between 2017 and 2018
2. Comparioson of Execution times of Original Script and Refactored Script

To explain above , we have the below outcomes, for Current(Refactored) Script and for Previous (Original Script)

#### Stock Performance in Year **2017** & **Current** Execution Time (Refactored Script)

![Year_2017_Current](https://github.com/kirtibhandari/stock-analysis/blob/main/Resources/VBA_2017.png)


#### Stock Performance in Year **2017** & **Previous** Execution Time (Original Script)

![Year_2017_Previous](https://github.com/kirtibhandari/stock-analysis/blob/main/Resources/2017_Previous_Execution.png)

#### Stock Performance in Year **2018** & **Current Execution** Time (Refactored Script)

![Year_2018_Current](https://github.com/kirtibhandari/stock-analysis/blob/main/Resources/VBA_2018.png)


#### Stock Performance in Year **2018** & **Previous** Execution Time (Original Script)

![Year_2018_Previous](https://github.com/kirtibhandari/stock-analysis/blob/main/Resources/2018_Previous_Execution.png)

From above we cane see:
### A: Stock Performance in year 2017 and 2018

In year 2017, stocks such as AY, CSIQ, FSLR, JKS, SPWR performed quite well in terms of 'Volume' traded in these two years and also gave positive 'Yearly Return' as compared to year 2018.

Except stock TERP, which gave negative return in both years, 2017 & 2018. Surely this stock is not advisable to invest in.

Year 2018, was overall didn't do quite well for almost all stocks except ENPH & RUN. These two stocks were tremendously profitable in 2018. ENPH returns were comparitively less profitable in year 2018, but stock RUN was the most profitable, from return of just 5.5% to 84.0% in year 2018.

Other stocks such as DQ, HASI, RUN, SEDG, TERP, VSLR were traded more in terms of 'Volume' in year 2018 than in 2017, but, overall, they all were in loss i.e. negative returns. So , surely, 'Volume' of stock traded in an year does not seem to be a right measure to choose a stock to invest into. 

Overall stock trading did well in 2017. 

DQ stock which client thought, because it was traded lot in year 2018, over year 2017 could be a good investment, is surely not a good idea as per this analysis. Though it was traded with very high difference in 'Volume' over year 2017 to year 2018 but it's yearly return was a huge decline from 199% in 2017 to -62.6% in year 2018.

### B: Execution times of Previous ('Original') script vs Current('Refactored') script

As we can clearly see from the above pictures, for execution times, that , our  Current('Refactored') script performed better than Previous('Original') script. The new execution times for year 2017 & 2018 are 0.1455078 seconds and 0.1416016 seconds respectively, definitely lesser than previous execution times of 0.8085938 seconds and 0.7617188 seconds for years 2017 and 2018 respectively.

Below are the screenshots of code snippets, that were different in both the scripts, which played an important role in reduction of execution times, making Refactored script perform better than Original script.

#### **Original Script Main Logic**

![Original Script](https://github.com/kirtibhandari/stock-analysis/blob/main/Resources/Original_Script_Main_Logic.png)

#### **Refactored Script Main Logic**

![Refactored Script](https://github.com/kirtibhandari/stock-analysis/blob/main/Resources/Refactored_Script.png)

Both the scripts provide same results, but with different code, where refactored code used lesser resources and executed with lesser computation and processing times. Original script has a code that executes main logic 'For' loop 36156 times whereas the 'For' loop in Refactored script got executed only for 3011 times, which is significantly a saver.

## Summary

### Detailed Statement on advantages and disadvantages of refactoring code in general

ADVANTAGES:
* A refactored code helps in reducing execution times with faster computations.
* Lesser computations and hence lesser resource utilization when code refactoring is done
* Efficient logic through code refactoring is easy to understand and trace/connect in terms of nested conditions or control structures, when less number of code lines.
* Also, a refactored understandable code is easy for bug fixes.
* A refactored code is an opportunity to design a software well so that it can easily be integrated with future technological advancements.

DISADVANTAGES:
* No new functionality is added.
* Refactoring might need expertise and hence senior working employees and hence, expensive to deploy them for these tasks.
* Refactoring sometimes can be time consuming and outcomes can be better but not with a significant quantum.
* Code refactoring is not easy and might need tighter couplings of conditionals and control structures, which might not be comprehensible by to all, in present and future.
* It might be hard to convince clients to invest in refactoring of code as all clients might not appreciate the final value it will add to a software depending upon their requirements og just the functionality they want to see in their software.

### Detailed Statement on the advantages and disadvantages of the original and refactored VBA script


In terms of initialization of 'tickers' array of stocks, both, original and refactored scripts have a scope of improvement in the context that it automatically picks up unique ticker names from the available data set and automatically gets stored in an array.

ADVANTAGES:
* Refactored script is better performing as it got main For loop executed only 3011 times in total as compared to refactored script where the main logic For loop had to execute/compute 36156  times.
* Lesser computations in refactored script lead to lesser execution times and hence less resources used such as memory, power etcetera.
* Refactored script is more understandable where logic is separated in specific lines of code and initialization of required variables, arrays as well as printing of output is done in separate code segments leading to better code design. Whereas in original script, initialization of an array and a variable , main logic For loop, a nested For loop, conditional structures and printing of output, all under one code segment. This surely is not the best way to write a code.
* It is easy to locate bugs in refactored script than in original due to its better code design in addition to being adaptive and maintainable.

DISADVANTAGES:
* We need to decide on what level of performance do we intend to achieve through code refactoring. Here, after spending hours of time to get all code running, to achieve the desired functionality seems lesser useful on current small data set. If the same logic works for entire stock market, we would better be able to justify code refactoring benefit if processing times reduced significantly and not only just as a part of a second.
* We have used more control structures in refactored script than original script. We used all initializations at almost one place, logic at one and printing at one place. This might not work when we would need temporary variables or outputs that need to be presented in real time, not at the time when all results have been computed.
* If it was not our client in this case, and there is a client and service provider relationship where client is a stock investment software developer, he would have charged more for providing a software which works on refactored script because of its time consumption. Thus, end user or client of software might end up paying more where the functionality provided is same.


Source used for better understanding:
https://stackoverflow.com/questions/43983284/what-are-the-advantages-and-disadvantages-of-refactoring-code-smell-in-software







