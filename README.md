# Refactor VBA Code and Measure Performance

## Overview of Project

Steve wants the stock dataset to expand to include the entire stock market over the last two years. 

### Purpose

Since the dataset will eventually expand to thousands of rows, the VBA code needs to be refactored to loop through all the data at one time. 
To determine if the refactoring is successful, the run times of the refactored VBA script need to be compared to the original VBA script.


### Results

The refactored code does produce small gains in efficiency.That small gain in efficiency will help when the dataset grows to thousands of rows.

2017 stock results time:

![alt text](https://github.com/sarifrey/stock_analysis/blob/main/Resources/VBA_Challenge_2017.png)


2018 stock return results time:
![alt text](https://github.com/sarifrey/stock_analysis/blob/main/Resources/VBA_Challenge_2018.png)


### Summary

The advantage of refactoring code is I did not see any. It was minor tweaks in code that required a lot of effort for small gains.
The disadvantage was in refactoring code figuring out how to best refactor the code.

The advantage of the refactored script is the nominal efficiency in getting results. Even a small increase in speed of results can help Steve do his work more efficiently. 
The advantage of the original code is that is easier to follow if one is not very familiar with VBA.

The disadvantage of the original VBA script was learning how to write it with little experience and background in VBA. It was not an intuitive process. 
The disadvantage of the refactored VBA script was learning how to make the changes.
