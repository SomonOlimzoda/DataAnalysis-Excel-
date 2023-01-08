
## Lab 1
Download the lab file [Lab1Start v5.xlsx](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Lab1Start%20v5.xlsx) to answer the questions below.

For the **Yearly Category Revenue** chart, display the total revenue for each year as a separate line.

```
1. Create a Total Revenue line in the range C7 to H7 by summing the annual revenue in rows 3 through 6. 
2. Label the row as Total Revenue in cell B7
3. Select the Yearly Category Revenue chart.
4. The data source for the chart is highlighted.
5. Drag the highlighted selection to include the total row (B7 to H7)
```

### Question 1

What does the Yearly total sales look like?

:heavy_check_mark: Similar to the yearly bikes revenue


### Question 2

What does the Yearly total sales look like?

:heavy_check_mark: 72%

For the Revenue by Country chart, sort the countries to display the one with the highest revenue showing at the top and the one with the lowest revenue at the bottom.

```
Select the data source for the Revenue by Country chart.
Sort the data by Revenue from Largest to Smallest ensuring you also sort the countries with the revenue.
Select the Revenue by Country chart.
Select the Y axis, the axis that shows the country name.
Select Categories in reverse order in the Axis options.
```

### Question 3

Rank the Countries from the highest to lowest revenue. (Highest in Zone 1, Lowest in Zone 6)

:heavy_check_mark: 1. United States
2. Austalia
3. United Kingdom
4. Germany
5. France
6. Canada

## Quiz time 

Download the quiz file [SalesForCourse_quizz_table.xlsx](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/SalesForCourse_quizz_table.xlsx) to answer the questions below.

### Question 1

What is the difference between a table and a range in Excel? Select four that apply.

:heavy_check_mark: In tables, you can add many types of totals for each column without writing any formulas.
:heavy_check_mark:Tables are formatted with alternate colors by default.
:heavy_check_mark:Formulas in tables refer to other columns by name and not by regular Excel reference.
:heavy_check_mark:Formulas in table columns are automatically applied to new rows.

### Question 2

If you filter the table to show only sales in United Kingdom, what will be calculated in the Total row?

:heavy_check_mark: The totals will recalculate according to the filtered rows.

### Question 3

You are only interested in the data for male customers in France. After inserting a Total Row to the Excel table and setting the appropriate filters, answer the following questions.

How many rows are displayed (what is the record count)? Do not include thousands separator (for example commas or semi-colons).

:heavy_check_mark: 2573
 
What is the total revenue for the displayed rows? Only provide whole number. Do not provide symbols, thousands separators, or any decimals places.

:heavy_check_mark: 1770366
 
What is the average revenue for the displayed rows? Only provide numbers, no currency symbols, thousands separators, and report to two decimal places (including the decimal in the format 000.00).

:heavy_check_mark: 688.06

## Lab 2

Download the lab file [Lab2Start v5.xlsx](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Lab2Start%20v5.xlsx) to answer the questions below.

The first thing you need to do is to convert the data into an Excel table.

Once you do that, you can add total row, filter the data, and the total will reflect the total only for the filtered data. Let's try this. Add a total row for the table, and use the Sum aggregation to show the total of the Revenue column and then filter the data only for United States.

### Question 1

What is the total revenue for all the sales in the United States?

:heavy_check_mark: 27975547

Now, you need to add several columns, derived from existing columns in the data. Before adding columns, it is good practice to clear any filters you previously applied.

First, let's add a "Month" column. Insert a new column to the left of the Customer ID column, and use formula to derive the month of sales from the Date column.
Here we use the Text() function and look for a format code in the examples that would work on a date field.

### Question 2

What is the total revenue for all the sales in the month of December?

:heavy_check_mark: 9086931

Next, let's add an "Age Group" column. Remember to clear any filters you previously applied. Insert a new column to the left of the Customer Gender, and use formula to derive the age group from the Customer Age column. Let's group the customers based on the following criteria:

* Youth (<25)
* Young Adults (25-34)
* Adults (35-64)
* Seniors (>64)
Here we use the nested IF() functions. Alternatively, you can use the IFS() function if it is available in your version (2016 + updates from O365)

### Question 3

What is the total revenue for all the sales for Young Adults Age Group?

:heavy_check_mark: 30655614

Now, let's add a "Frame Size" column. Insert a new column to the left of the Order Quantity, and use a combination of the IF() and RIGHT() functions to derive the frame size of a bicycle from the last two characters of the Product column, when the Product Category is Bikes. Otherwise, leave it blank.
Here we use the IF() function to test for Product Category=”Bikes” and if it does, use the RIGHT() function to extract the last two characters of the Product column.

### Question 4

What is the total revenue for all the bikes with frame size 62 for the customer age group Seniors?

:heavy_check_mark: 12452

Last but not least, let's add a "Profit" column. Insert a new column to the right of the Revenue, and use formula to derive the Profit from both the Revenue and Cost columns. Show the total for the Profit column. Use the Sum aggregation in the total row of the table, for the Profit column(Profit is Revenue minus Cost).

### Question 5

What is the total profit for United States sales in the month of October 2015, for customer age group Adults?

:heavy_check_mark: 138419

...




















