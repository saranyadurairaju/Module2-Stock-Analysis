# Stock Analysis with VBA

Sometime Users may require to automate some aspects of Excel, such as repetitive tasks, frequent tasks, generating reports, etc. VBA is used to automate tasks and perform several other functions beyond creating and organizing spreadsheets. 

## Overview of Project

Steve wants to find the total daily volume and yearly return for each stock. **Daily volume** is the total number of shares traded throughout the day; it measures how actively a stock is traded. **The yearly return** is the percentage difference in price from the beginning of the year to the end of the year. 


### Purpose

Steve may want to perform the Analysis each year and for a different set of stocks in the future. With this in mind, we should create a flexible macro for running multiple stocks, for all available years.

Also, Steve may want to perform his analysis on larger datasets, and he wants to know how fast his VBA code will compile the results. So, we need to add a script that will calculate how long the code takes to execute and output the elapsed time in a message box.


## Analysis 

As there are large number of data for different set of stocks during many years, we may need to do some analyze and write a generic Macro which will work for any data in future. 

* Finding the number of Rows

	RowCount = Cells(Rows.Count, "A").End(xlUp).Row

* Using Arrays for multiple Stocks

	Dim tickers(12) As String
	
	Dim tickerVolumes(12) As Long
    	
	Dim tickerStartingPrices(12) As Single
    	
	Dim tickerEndingPrices(12) As Single

* Getting the user Input "Year" to automate 

	yearValue = InputBox("What year would you like to run the analysis on?")

* Calculating the run time of the program

	MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
	
## Challenges

* Finding the correct way to automate the whole process and getting the result with a minimum run time is little challenging. But this helps us to try doing the things in many ways, which results in better solution. 

* We have to make sure the data is sorted in Ascending order based on the Stocks (Ticker) and given in the same order inside VB Script as well.

## Results

Below is the Macro enabled excel worksheet for Stock Analysis. 

![All_Stocks_Analysis](https://github.com/saranyadurairaju/Module1-Final-Assignment-Analysis/blob/main/Kickstarter_Challenge.xlsx)


#### The performance between 2017 and 2018

* We have compared the same set of Stocks for the Year 2017 and 2018 with Volumes, Open & Close Values etc.

* To find the Return value of any particular stock, we are using the start and end of year Closing values.

* In 2017 only " TERP" stock returns is dropped but all the other 11 Stock values are hiked.

* In 2018 "ENPH" and "RUN" stocks return are hiked, but all the other 10 Stocks are dropped. 

* By seeing all the above results, Stocks are performed well in 2017 compare to 2018.

![VBA_Challenge__2017 Vs 2018](https://github.com/saranyadurairaju/Module1-Final-Assignment-Analysis/blob/main/Outcomes_vs_Goals.png)


#### The original and the refactored script

* In the original code we are reading the entire data again and again to calculate the Volumes and Return values for all the Stocks, which takes lot of Elapse time.

* So, we are using Arrays to save all the values while we are reading the data for the first time and storing the values in corresponding arrays.

* We are also using "For" loops and initialization of Values to optimize the script.

![VBA_Challenge_Coding](https://github.com/saranyadurairaju/Module1-Final-Assignment-Analysis/blob/main/Outcomes_vs_Goals.png)
 

## Summary

### Refactoring Code

Code Refactoring is a way of restructuring and optimizing existing code without changing its behavior. It is a way to improve the code quality and a sharp weapon for developers in their maintenance activities. 

**Advantages**

- Make the code clean and organized
- Helps Finding any mistakes
- Improves Program run time
- Interesting and challenging thing to do

**Disadvantages**

- Time consuming work 
- Difficult to do when the Program is too big
- Developer should have a complete understanding
- Current program output shouldn't be impacted 


### Pros and Cons of Refactoring Stock Analysis script

The Original and refactoring Script gave the same result in our Stock Analysis script. So, its clear that we didn't impact the any output. But Refactoring has its own Pros and Cons.

**Pros**

- It was very interesting to do
- Learned many new things
- Elapse time improved a lot
- Program became very structured

**Cons**

- Programmer should understand the concepts clearly
- Have to handle all the arrays properly
- It will be a problem if the data is not in order


Everything is automated and Steve can help his parent's for their investment with the output data in no time. Wishing Steve all the very best!
