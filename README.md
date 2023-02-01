# Data Analysis with Excel

### Scenario
I'm a new marketing manager of an established Bicycle company. The company sells bicycles and accessories, such as clothing and other accessories to bikers in six countries.

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/bicycle.png)


The company has just hired Lucy as its new Sales manager. I'm tasked to introduce Lucy to the company, its product portfolio and its sales performance since 2011. To do this, I should ask Jack, the IT manager to prepare some data for me.

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Lucy%20and%20Jack.png)
##

### Lab 1
Download the lab file [Lab1Start v5.xlsx](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Lab1Start%20v5.xlsx) to answer the questions below.

#### Question 1

What does the Yearly total sales look like?

#### Question 2

What does the Yearly total sales look like?

#### Question 3

Rank the Countries from the highest to lowest revenue. 
##

### Quiz time 

Download the quiz file [SalesForCourse_quizz_table.xlsx](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/SalesForCourse_quizz_table.xlsx) to answer the questions below.

#### Question 1

What is the difference between a table and a range in Excel?

#### Question 2

If you filter the table to show only sales in United Kingdom, what will be calculated in the Total row?

#### Question 3

How many rows are displayed (what is the record count)?
##

### Scenario

While my first attempt to show the company's performance to Lucy was not bad, clearly she has a lot more requirements than what you provided. She wants to know more about the year over year sales, sliced into different categories, sub-categories, and countries. She also wants to see additional information such as customer demographics.

Jack has provided me with a different data source. This time the data has more than one hundred thousand rows. Before I can create additional reports to Lucy, first I need to prepare the data.
##

### Lab 2

Download the lab file [Lab2Start v5.xlsx](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Lab2Start%20v5.xlsx) to answer the questions below.

#### Question 1

What is the total revenue for all the sales in the United States?

#### Question 2

What is the total revenue for all the sales in the month of December?

#### Question 3

What is the total revenue for all the sales for Young Adults Age Group?

#### Question 4

What is the total revenue for all the bikes with frame size 62 for the customer age group Seniors?

#### Question 5

What is the total profit for United States sales in the month of October 2015, for customer age group Adults?
##

### Quiz time 

Download the quiz file [SalesForCourse_quizz_table.xlsx](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/SalesForCourse_quizz_table.xlsx) to answer questions 4 through 7 below.

#### Q1. Drag and Drop

You have the following Pivot Table fields. You need to calculate the number of sales transactions by age. Drag the appropriate field names and aggregated field names to the appropriate Pivot Table Areas (Filters, Rows, Columns, or Values) to perform the task

#### Q2. Drag and Drop

You have the following Pivot Table fields. Your manager asks you to build a pivot table that shows revenue by state and country. How would you arrange the field names and aggregated field names to the appropriate Pivot Table Areas (Filters, Rows, Columns, or Values)?

#### Question 3

You have a pivot that is based on an Excel table. You need to add new rows to the Excel table and want to be able to see the values for these new rows in the PivotTable. What should you do?

#### Question 4

Using the data provided, create a pivot with “Country “as rows, “Month” as columns, and “Revenue” as values. Drag “Year” to the filter area. What does the value 286,779 in the first summarized field of the pivot table (uppermost left value) represent?
##

#### Q5. Numerical Input

Create a new PivotTable with “Customer Gender” as rows, “Product Category” as columns, and “Revenue” as values.

a. How much revenue came from the purchase of bikes by women? Provide a whole number without currency symbols, thousands separators (for example a comma or semi-colon), or decimal places.

b. How much revenue came from clothing purchased by men? Provide a whole number without currency symbols., thousands separators, or decimal places.
##

#### Question 6

Create a new PivotTable with “Country” as rows, “Product Category” as columns, “Year” as a filter, and “Revenue” as values. Filter by year 2015.

Drill down into the data for the cell that represents the revenue for bike sales in Germany, 2015. A new worksheet is created with a table in it. Which option below best describes the content of this table?
##

#### Q7. Numerical Input

a. Total revenue from bike sales from the country with the highest revenue from bike sales in 2016:

b. Number of sales transactions for bikes in the US in 2016:

c. Average number of items per transaction for clothing purchased in the US during 2016 rounded to the nearest whole number:

d. The largest revenue value registered in a single transaction in the Germany during 2015:

e. The largest revenue value registered in a single transaction during 2015:
##

### Scenario

Now I have prepared the data in an Excel table, I can start to create pivot tables to aggregate the data and create some reports. 

From conversation with Lucy, I know that she is interested in looking into the yearly sales data broken down by countries, product categories, and age groups.
##

### Lab 3A

Download the lab file [Lab3Start v5.xlsx](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Lab3AStart%20v5.xlsx) to answer the questions below.

First, let's start by naming the Excel table. Name the Excel table prepared in the previous lab to SalesTable. From now on, every time you add a pivot table, it should be based on this SalesTable, unless mentioned otherwise.

Now, proceed to add your first pivot table. Insert a new pivot table based on the SalesTable to a new sheet. Arrange the layout so that the pivot table displays the Product Category and Sub Category in the Rows, Year in the Columns, and Revenue (Sum of) as the Values.
##

#### Question 1

Which year did the company start selling Touring Bikes?
##

Insert another pivot table to the same sheet, next to the existing pivot table. Arrange the layout so that the pivot table displays the Country and State in the Rows, Year in the Columns, and Revenue (Sum of) as the Values. Sort the pivot table by Sum of Revenue so that the Country and State with the highest revenue is displayed first.
##

#### Question 2

Rank the States for Germany, from the highest to lowest revenue.

#### Question 3

What about for the year 2013? Rank from the highest to lowest revenue. 
##

Let's add another pivot table. This time arrange the layout so that the pivot table displays the Frame Size in the Rows and Revenue (Sum of) as the Values. Hide the rows that do not have a Frame size (blank Frame size), then sort the pivot table by Sum of Revenue so that the Frame size with the highest revenue is displayed first.
##

#### Question 4

Which Frame Size has the highest revenue?
##

Last but not least, add another pivot table with Age Group as the Rows and Revenue (Sum of) as the Values. You will learn how to custom sort the Age Group in the next module. But for now, sort the pivot table by Sum of Revenue so that the Age Group with the highest revenue is displayed first.
##

#### Question 5

Which Age Group has the lowest revenue?
##

Save your Excel file - you will need this work in the next Exercise (Lab 3B).

### Lab 3B

Now you can start adding some charts to the sheet.

First, add a pivot chart for the pivot table that shows yearly sales (revenue) by Country (the pivot table you created for question 2 in Lab 3A). Select a Line chart to display the yearly trend. Make sure that the Years are located in the X axis, the Revenue in the Y axis, and the Countries as categories.
##

#### Question 1

In this chart, you can clearly see the sales trend for each Country. Which country’s trend displays the most fluctuation in sales when compared to the other countries?
##

Add another pivot chart for the pivot table that shows yearly sales (revenue) by Product Category (the pivot table you created for question 1 in Lab 3A). Select a Column chart to display the yearly sales by category so that the years are together.
##

#### Question 2

In which year does the Bikes category have the lowest revenue?
##

Add another pivot chart, this time for the pivot table that shows Revenue by Age Group (the pivot table you created for question 5 in Lab 3A). Select a Pie chart to display the proportion of each Age Group (remember the chart styles) with data labels, formatted to two decimal points.
##

#### Question 3

In this chart, you can clearly see the proportion of sales (revenue) for each Age Group. What is the proportion of sales (revenue) for Young Adults?
##

Add another pivot chart, this time for the pivot table that shows Revenue by Frame size (the pivot table you created for question 4 in Lab 3A). Select a Bar chart to display the order of revenue by Frame size. Sort the Y axis to show the Frame size that has the highest revenue on the top.
##

#### Question 4

In this chart, you can clearly see the order of sales (revenue) by Frame size. Which frame size had the lowest revenue?
##

### Quiz time 

Download the quiz file [SalesForCourse_quizz_table.xlsx](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/SalesForCourse_quizz_table.xlsx) to answer Question 1 below.
##

#### Question 1

Using the data provided, create a pivot with “Country” as rows, “Product Category” as columns, and “Revenue” as values. Slice months by “April.” What does the grand total for revenue (2,200,490) represent?

#### Question 2

Which objects can be sliced by slicers? Choose 3 that apply.

#### Question 3

When would you use both a slicer and a filter based on the same field?
## 

#### Question 4

You have the following slicers.

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Q2.png)

Why are Bike Racks, Fenders and Gloves at the bottom of the Sub Category slicer in a lighter color?
##

#### Question 5

You have the following slicers.

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Q3.png)

What will happen if you click "Bike Racks"?
##

### Scenario

I have created several pivot tables and pivot charts for Lucy.

You can download the latest report [here](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Lab4AStart%20v5.xlsx).

So far everything has been well received by her. However, she would like to have easier ways to slice and dice the reports and charts herself.

I sat down with Lucy, and come up with several different ways that Lucy could slice the data

- Year
- Country
- Customer Gender
- Age Group
- Product Category
- Sub Category
- Frame size
##

### Lab 4A



