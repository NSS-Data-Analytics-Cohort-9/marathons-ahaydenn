### Nashville Rock-and-Roll Marathon and Half-Marathon Race Results (2016 - 2019)

1. Create a new blank workbook and load the 8 marathon and half-marathon tables from marathons.accdb, a MS Access database called (Data --> Get Data --> From Database --> From Microsoft Access Database).  
    a. Load each table to a separate worksheet. Give the worksheets *_meaningful_* names.  
    b. Be sure to format the Time column after load (HH:MM:SS). If you get any load errors, try to understand whether you can reload the data and correct them.  
    c. Create another worksheet and name it `analysis`. This is where you should do your analysis work.  

2. Use built-in functions to answer the following questions. Be sure to display your analysis work in a neat, easy-to-follow format.    
    a. Find the fastest time, slowest time, median, and mean times for each of the 8 races.  
    b. Find the mean and median marathon finish times for all 4 years combined.  
    c. Find the mean and median half-marathon finish times for all 4 years combined.  

3. First time marathoners sometimes set a goal of beating “Oprah’s time,” Oprah Winfrey’s time (04:29:20) in the 1994 Marine Corps Marathon when she was 40 years old.  
        a. How many runners each year beat Oprah’s time?  
        b. What percentage of runners in each of the 4 marathons beat Oprah’s time?  

4. Quartiles help group ordered data into buckets – the first quartile separates the lowest 25 percent of the data from the top 75 percent, the second quartile (same as the median) divides the dataset in half, and the third quartile separates the lowest 75 percent from the top 25 percent. Find the values that define the first, second, and third quartiles for each half-marathon.

5. Is there a year in which runners seem slower or faster? Formulate a hypothesis for any apparent differences.

6. Scott Wietecha has won the Rock and Roll Marathon for 7 years in a row. Compute and display the difference between Wietecha’s time and the next fastest runner for each year.

#### BONUS: Find the top three marathon runners for 2016, 2017, 2018, and 2019. Remove any duplicates across years (i.e., Scott Wietecha should only be in the list once). How many unique runners finished at one of the top 3 spots in the past 4 races? In how many of the 4 years did each top 3 runner finish? How has each runner’s time changed from year to year?

##### Create a brief (5 minute) presentation to communicate your findings from this exercise.

## Rock and Roll Marathons Bonus Questions
In these bonus questions, you'll get to explore some additional features of Excel.
1. First, notice that when you imported the data, it came in as an [Excel table](https://support.microsoft.com/en-us/office/create-and-format-tables-e81aa349-b006-4f8a-9806-5af9df0ac664). Excel tables make it easier to perform analysis and write functions on datasets. If you have not already done so, rewrite your formulas for the fastest, slowest, median, and mean time using [Structured References](https://support.microsoft.com/en-us/office/using-structured-references-with-excel-tables-f5ed2452-2337-4f71-bed3-c8ae6d2b276e). This means that you should be referring to the table names and column names rather than using explicit cell references (rows and columns).
2. The [INDIRECT](https://support.microsoft.com/en-au/office/indirect-function-474b3a3a-8a26-4f44-b491-92b6306fa261) function allows you to make references using strings. If you have not already done so, redo your calculations for Question 2 so that you only need to write one function per year which can be copied down to calculate values for all years.
3. When performing calculations or making comparisons to a constant value, you can often improve the readability of your spreadsheet by using named cells. You can read more about working with named cells [here](https://support.microsoft.com/en-au/office/define-and-use-names-in-formulas-4d0f13ac-53b7-422e-afd2-abd7ff379c64). If you have not already done so, create a cell which contains Oprah's time and name this cell "OprahTime". Reference this cell by name in your calculations on Question 3.
4. Working with Excel tables make certain operations easy to accomplish. For example, let's say we wanted to see if there were any duplicate names in the 2016 marathon. Follow these steps to determine this.  
a. First, create a column that counts how many times a name appears in the Name column of the 2016 marathons table. (Hint: Use the COUNTIF function). Notice how when you write your function, it automatically populates the entire column of the teable.
b. Next, you can use this new column to filter the table down to just the rows with duplicates.
c. From what you can see, do any of these look like data errors or are they all legitimate duplicate names?  
d. Now, reset the filter so that you are viewing the entire table again.  
e. Let's use another approach to looking for potential duplicates - [conditional formatting](https://support.microsoft.com/en-us/office/use-conditional-formatting-to-highlight-information-fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f). Highlight the Name column of your table. Under the Home tab, select Conditional Formatting -> Highlight Cell Rules -> Duplicate Values. This will highlight any rows which are duplicate names.  
f. At this point you could scroll through the table to find duplicates, but with a few thousand rows, this is cumbersome. Instead, we can use the filter feature of the table. Click the filter option on the Name column and select "Filter by Color" and utilize the "Filter by Cell Color" to retrieve the duplicate rows.
5. Let's make a visualization to see the overall distribution of finish times during the 2016 marathon.  
a. Start by creating a new column which contains the finish time of each runner in minutes. Hint: This can be done using the HOUR, MINUTE, AND SECOND functions.  
b. Now, using this column, create a [histogram](https://www.mathsisfun.com/data/histograms.html). You can create such a chart by going to the Insert tab, going to the charts section and choosing histogram under the "Insert Statistic Chart" button. (See [this page](https://support.microsoft.com/en-us/topic/create-a-histogram-in-excel-a15d4de8-a432-72cd-9434-1a7f3e88698e) for more information).  
c. What do you notice from this histogram?  
 6. You've seen from the previous calculations that marathon times were slower for 2017 compared to 2016, but now let's see if the times of individual runners were slower.   
 a. Make a new sheet, "Year_to_Year_Comparison".   
 b. In this sheet, create a table that contains the name of each runner who participated in the 2016 full marathon.   
 c. In the second column, put the finish times for those runners for the 2016 marathon. Hint: you may want to use a lookup function such as VLOOKUP or [XLOOKUP](https://support.microsoft.com/en-us/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929) to pull these numbers.
 d. Add a third column that shows the 2017 time for runners that also participated in the 2017 marathon.  
 e. Filter the data down to just the rows for runners who participated both in 2016 and 2017.    
 f. What percentage of runners who participated in both the 2016 and 2017 marathons had a slower finish time in 2017 compared to 2016?  
 g. Who had the largest improvement in time from 2016 to 2017? Who had the biggest drop off?  
 h. Build off of what you did to pull the 2017 times to determine how many runners competed in every year from 2016 through 2019. Out of these runners, which had the largest improvement in finish time from 2016 to 2019?




