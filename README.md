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

* Ran the 2017 code in 0.86 seconds

![2017 Run-time (Original code)](https://github.com/temersonzetina/stock-analysis/blob/main/VBA_Challenge_2017_Original.png)

* Ran the 2018 code in 0.91 seconds

![2018 Run-time (Original code)](https://github.com/temersonzetina/stock-analysis/blob/main/VBA_Challenge_2018_Original.png)

The refactored script bypassed the nested loop and was able to produce the same information by looping through the ticker data just once:

* Ran the 2017 code in 0.16 seconds

![2017 Run-time (Refactored)](https://github.com/temersonzetina/stock-analysis/blob/main/VBA_Challenge_2017_Refactored.png)

* Ran the 2018 code in 0.16 seconds

![2018 Run-time (Refactored)](https://github.com/temersonzetina/stock-analysis/blob/main/VBA_Challenge_2018_Refactored.png)

## Summary
The process of refactoring produces a set of advantages and disadvantages that developers should consider well before revisiting an existing project.

### Advantages
**Reduced run-time**

One of the key advantages to refactoring is its ability to provide faster access to information. This allows developers and users alike to execute functions with smaller delays. In the context of financial investing, where an increasing number of transactions are automated, timely information provides an undeniable advantage to brokers and investors alike.

**Cleaner code**

Refactoring in many cases results in more elegant, more concise code that executes the same functions as the original. Refactored code makes debugging much simpler and may save developers' time when reviewing or altering the code in the future.

### Disadvantages

**Limited time**
Technically, if the code is already functional, it doesn't "need" to be improved. If developer capacity is stretched, refactoring adds another task to an already-lengthy list of deliverables. In that sense, refactoring may distract from other projects.

**Conceptually complexity**
While not a "disadvantage" per se, refactoring may require developers to develop a deeper conceptual understanding of certain concepts. For instance, in the case of the VBA Challenge, creating a "ticker index" that allowed us to bypass a nested For loop and circulate through the data just once meant conceptualizing and integrating another array. On the one hand, this array enabled greater efficiency. On the other, it added a layer of complexity that - though not a hindrance - requires that the developer have existing expertise or that they advance their understanding of software processes.

<How do these pros and cons apply to refactoring the original VBS script?>


