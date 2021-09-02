# stock-analysis
Using VBA to conduct stock market analysis for company performance in 2017 and 2018
# Stock Analysis

## Overview
The first purpose of this analysis was to enable the viewer to quickly observe two key indicators of performance for 12 companies that are publicly traded in the US stock market. The first indicator was "total volume," which reflected the total number of trade transactions that a given company's stock underwent in one of the two years for which there was data (2017 and 2018). The second indicator was "return," which describes (by percentage) a stock's increase or decrease in value over a given year.

The second purpose of this analysis was to refactor code that, though functional, included a nested loop that made the run-time needlessly lengthy.

It should be noted that though the source data for this project includes performance information for 12 companies, the script should work on a much larger dataset of corporations.

## Results

### Comparing Yearly Yields
Stock trade volumes and yields varied greatly between 2017 and 2018, as the comparison chart below demonstrates.

![Volume and Return Comparison](https://github.com/temersonzetina/stock-analysis/blob/main/Yearly_Volume_Return_Comparison.png)

To summarize comparative performance:

* 2 of the 12 companies (ENPH and RUN) yielded positive gains in both 2017 and 2018
* 1 company (TERP) had negative yields in both years
* The remaining 9 companies yielded positive gains in 2017 and negative gains in 2018

### Comparing Run Times
One of the key purposes of refactoring is to improve run-time. The original script, which involved a nested For loop:

* Ran the 2017 performance data in 0.86 seconds

![2017 Run-time (Original code)](https://github.com/temersonzetina/stock-analysis/blob/main/VBA_Challenge_2017_Original.png)

* Ran the 2018 performance data in 0.91 seconds

![2018 Run-time (Original code)](https://github.com/temersonzetina/stock-analysis/blob/main/VBA_Challenge_2018_Original.png)

The refactored script bypassed a nested loop from the original script and was able to produce the same information by looping through the ticker data just once. The refactored script:

* Ran the 2017 performance data in 0.16 seconds

![2017 Run-time (Refactored)](https://github.com/temersonzetina/stock-analysis/blob/main/VBA_Challenge_2017_Refactored.png)

* Ran the 2018 performance data in 0.16 seconds

![2018 Run-time (Refactored)](https://github.com/temersonzetina/stock-analysis/blob/main/VBA_Challenge_2018_Refactored.png)

## Summary
The process of refactoring produces a set of advantages and disadvantages that developers should consider well before revisiting an existing project.

### Advantages
**Reduced run-time**

One of the key advantages to refactoring is its ability to provide faster access to information. This allows both developers and users to execute functions with smaller delays. In the context of financial investing, an industry that increasingly relies upon automation and data modeling, timely information provides an undeniable advantage to brokers and investors alike.

**Cleaner code**

Refactoring in many cases results in more elegant, more concise code that executes the same functions as the original. Refactored code makes debugging much simpler and may save developers' time when reviewing or altering the code in the future.

### Disadvantages

**Limited time**

Technically, if the code is already functional, it doesn't "need" to be improved. If developer capacity is stretched, refactoring adds another task to an already-lengthy list of deliverables. In that sense, refactoring may distract from other projects.

**Conceptually complexity**

While not a "disadvantage" per se, refactoring may require developers to develop a deeper conceptual understanding of certain tasks and the skills that go along with them. For instance, in the case of the VBA Challenge, creating a "ticker index" that allowed us to bypass a nested For loop and circulate through the data just once meant conceptualizing and integrating another set of arrays. On the one hand, these arrays enabled greater efficiency, because we were able to store all of the necessary information about each of the company's performance data during a single run through the data. Conversely, refactoring added a layer of complexity that - though not a hindrance - involved a time-consuming process to determine how to accurately code the new loop. Taken generally, this process of refactoring may challenge developers in similar ways, asking them to problem-solve for a project that already "works." Of course, in the long run, this deepened expertise is of great benefit to an analyst and their company. But in the short-term, when capacity may run thin and engineering time is precious, refactoring may prove burdensome.

### Pros and Con's of Refactoring the VBA Script
The VBA script ran significantly faster after refactoring, each revised analysis taking around 1/5 the time it took during the original analysis. This seems largely due to the fact that the refactored code looped through the data worksheet just once, as opposed to twelve times as in the case of the original code.

In terms of having "cleaner code," eliminating the nested for loop simplified the process of looping through the worksheet data. However, as described above, it also meant understanding how to use the ticker index variable. Figuring out how to integrate this element into phrases alongside the output arrays may have been the single most difficult challenge associated with this project. This speaks to the difficulty related to deepening one's knowledge of software processes while writing more efficient code.


