# stock-analysis
Using VBA to conduct stock market analysis for company performance in 2017 and 2018

# Stock Analysis

## Overview
<What was the purpose of this analysis?>
The first purpose of this analysis was to enable the viewer to quickly observe two key indicators  of performance for 12 companies that are publicly traded in the US stock market. The first indicator was "total volume," which reflected the total number of trade transactions that a given company's stock underwent in one of the two years for which there was data (2017 and 2018). The second indicator was "return," which describes a stock's return in a given year.

The second purpose of this analysis was to refactor code that, though functional, included a nested loop that made the run-time needlessly lengthy.

## Results
<Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution time of the original script and the refactored script>

### Comparing Yearly Yields
Stock trade volumes and yields varied greatly between 2017 and 2018, as the comparison chart below demonstrates.

![Volume and Return Comparison](https://github.com/temersonzetina/stock-analysis/blob/main/Yearly_Volume_Return_Comparison.png)

In summmary:

* 2 of the 12 yielded positive gains in both 2017 and 2018 (ENPH and RUN)
* 1 (TERP) had negative yields in both years
* The remaining 9 yielded positive gains in 2017 and negative gains in 2018

### Comparing Run Times
The original script, which involved a nested For loop:

* Ran the 2017 code in X seconds

![2017 Run-time (Original code)](https://github.com/temersonzetina/stock-analysis/blob/main/VBA_Challenge_2017_Original.png)

* Ran the 2018 code in Y seconds

![2018 Run-time (Original code)](https://github.com/temersonzetina/stock-analysis/blob/main/VBA_Challenge_2018_Original.png)

The refactored script bypassed the nested loop and was able to produce the same information by looping through the ticker data just once:

* Ran the 2017 code in X seconds
![2017 Run-time (Refactored)](https://github.com/temersonzetina/stock-analysis/blob/main/VBA_Challenge_2017_Refactored.png)

* Ran the 2018 code in Y seconds
![2018 Run-time (Refactored)](https://github.com/temersonzetina/stock-analysis/blob/main/VBA_Challenge_2018_Refactored.png)

## Summary
The process of refactoring brings with it a set of advantages and disadvantages that developers should consider before revisiting an existing project.

### Advantages
-Reduced run-time; faster access to information allows developers and users alike to execute functions with smaller delays; In the context of stock market investing, where an increasing number of transactions are automated, timely information provides an undeniable advantage to brokers and investors alike
-Cleaner code; Refactoring in many cases results in more elegant, shorter code to execute the same functions; In such cases, debugging is much simpler and may save developers' time when reviewing or altering the code in the future

### Disadvantages
-Limited time; Technically, if the code is already functional, it doesn't "need" to be improved. If developer capacity is stretched, refactoring adds another task to an already-long list of work actions. In that sense, refactoring may distract from other projects.
-Conceptually complex; While not a "disadvantage" per se, refactoring may also require a deeper conceptual understanding that requires the developer to deepen their skill and knowledge of essential concepts. For instance, in the case of the VBA Challenge, creating a "ticker index" that allowed us to bypass a nested For loop and circulate through the data just once meant conceptualizing another array that enables greater efficiency. This adds a layer of complexity that - though not a hindrance - does require the developer to advance their understanding of data science.

<How do these pros and cons apply to refactoring the original VBS script?




