# VBA of Wall Street

## Overview of Project

The purpose of this project was to refactor VBA code in the macro-enabled workbook [VBA Challenge.xlsm](/VBA Challenge.xlsm) to reduce it's run time and in conjunction reduce the required processing dedication.

### Purpose

The primary purpose of the workbook is to output an analytical summary of the worksheet desired by the user. The macro will produce a list of tickers, their total daily volume, and their return.

Within the excel workbook, there are two modules attached with it. Module 1 contains VBA code for additional other macros created during the instructional section. Module 2 contains refactored code for the primary purpose of the new workbook.

## Results

Below are screenshots of the code & resultant processing times from Module 1- the original code.

#### Original code & time to process

![VBA_Challenge_2017Original.png](/resources/VBA_Challenge_2017Original.png)
![VBA_Challenge_2018Original.png](/resources/VBA_Challenge_2018Original.png)

In the original code, the code uses two loops but they are a nested loop structure.

![NestedForLoop.png](/resources/NestedForLoop.png)

The refactored VBA script still uses two loops but are not in a nested loop structure.

![TwoForLoop.png](/resources/TwoForLoop.png)

There are advantages to this touched on later in the summary. After refactoring the VBA code, the code runs faster by a minimum factor of three. The code may be found in from Module 2 of the workbook.

#### Refactored code & time to process

![VBA_Challenge_2017.png](/resources/VBA_Challenge_2017.png)
![VBA_Challenge_2017.png](/resources/VBA_Challenge_2018.png)

## Summary

#### Refactoring Advantages & Disadvantages

Refactoring code in general improves the design of the code and thus software overall. For instance, by taking advantage of repetitive code and using for loops, it is possible to reduce the number of lines in a code which subsequently reduces the code's file size. Refactoring frequently makes code and software easier to understand via the clean up, commenting and changes to conciseness. An example would be duplicate code, excessively long lines of code that make the code difficult to read, understand, or debug. To follow that, refactoring helps finding bugs because it is easier to find errors when approaching systematically piece by piece after an initial working code is completed.

Refactoring can be disadvantageous in that it costs time and consequently money. The term "spaghetti code" is a classic slang term to refer to non refactored code. Furthermore, it is not necessarily always possible to refactor a code, especially if it's spaghetti code due to overwhelming dependancies. 

#### Advantages & Disadvantages with respects to this case study

Refactoring the VBA script in this case study increased the speed of processing by a minimum factor of three. While the scaling was within 1 seconds worth of time, the potential of savings for a larger series of data would be significant. With respects to data management, by using a series of arrays instead of lone variables, there is an additional order applied to the script. If at any point a specific volume, starting price or ending price is required, there is an array to pull the information out and be used elsewhere where as in the unfactored code the information would be lost immediately on performing analysis of the next ticker. In the case of the original code, the script was structured to perform a nested loop to account for each ticker individually before outputting the results directly. In the refactored code, the script is structured in a single loop to analyze the data, then a second loop to present the data.  Even though the number of loops is the same, because the loops in the refactored code are not nested, it results in a shorter runtime. 

However, there are disadvantages to the code in general. The script will miss information if the data is not sorted by Ticker (either ascending or descending) and Date-Ascending. A method to remedy this would be to create an additional two arrays- one for start date, one for final date. The starting data for the two arrays would be different but would be replaced if another date that is earlier or later respectively and by virtue also allow the starting price and end price to be updated by the same row index.
