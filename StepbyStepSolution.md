## Data Analysis with Excel

### Scenario
I'm a new marketing manager of an established Bicycle company. The company sells bicycles and accessories, such as clothing and other accessories to bikers in six countries.

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/bicycle.png)

The company has just hired Lucy as its new Sales manager. I'm tasked to introduce Lucy to the company, its product portfolio and its sales performance since 2011. To do this, I should ask Jack, the IT manager to prepare some data for me.

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Lucy%20and%20Jack.png)

Now, it's my job to present this data in a compelling manner.
##

### Lab 1

Download the lab file [Lab1Start v5.xlsx](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Lab1Start%20v5.xlsx) to answer the questions below.

The first thing I'd like to do is to present the data graphically. I will do this by creating charts for the three groups of data received from Jack, our IT manager. I give each chart a descriptive title as shown in the three charts below.

A solution should look like this (numbers may not be the same).
##

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/1.png)

Once you have created the three charts above, let's customize them a little bit more.
##

For the **Yearly Category Revenue** chart, display the total revenue for each year as a separate line.

1. Create a Total Revenue line in the range C7 to H7 by summing the annual revenue in rows 3 through 6. Label the row as Total Revenue in cell B7
2. Select the **Yearly Category Revenue** chart.
3. The data source for the chart is highlighted.
4. Drag the highlighted selection to include the total row (B7 to H7)
##

#### Question 1

What does the Yearly total sales look like?

:heavy_check_mark: Similar to the yearly bikes revenue
##

For the **Revenue by Category** chart, display the percentage for each category.

1. Select the **Revenue by Category** chart.
2. From the Chart Tools **Design** tab, change the chart style to display the data as a percentage including the data label as percentages. 
You can do this by either selecting a style from the **Chart Styles** group or from the **Chart Layouts** group by selecting one of the **Quick Layouts**.
##

#### Question 2

What percentage of the total revenue comes from the **Bikes** category?

:heavy_check_mark: 72%
##

For the **Revenue by Country** chart, sort the countries to display the one with the highest revenue showing at the top and the one with the lowest revenue at the bottom.

1. Select the data source for the **Revenue by Country** chart.
2. Sort the data by **Revenue** from **Largest to Smallest** ensuring you also sort the countries with the revenue.
3. Select the **Revenue by Country** chart.
4. Select the Y axis, the axis that shows the country name.
5. Select **Categories in reverse order** in the **Axis options**.
##

#### Question 3

Rank the Countries from the highest to lowest revenue. 

:heavy_check_mark: 1. United Stated

:heavy_check_mark: 2. Australia

:heavy_check_mark: 3. United Kingdom

:heavy_check_mark: 4. Germany

:heavy_check_mark: 5. France

:heavy_check_mark: 6. Canada
##

### Quiz time 

Download the quiz file [SalesForCourse_quizz_table.xlsx](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/SalesForCourse_quizz_table.xlsx) to answer the questions below.
##

#### Question 1

What is the difference between a table and a range in Excel?

:heavy_check_mark: In tables, you can add many types of totals for each column without writing any formulas and tables are formatted with alternate colors by default.

:heavy_check_mark: Formulas in tables refer to other columns by name and not by regular Excel reference.

:heavy_check_mark: Formulas in table columns are automatically applied to new rows.
##

#### Question 2

If you filter the table to show only sales in United Kingdom, what will be calculated in the Total row?

:heavy_check_mark: The totals will recalculate according to the filtered rows.
##

#### Question 3

You are only interested in the data for male customers in France. After inserting a Total Row to the Excel table and setting the appropriate filters, answer the following questions.

a. How many rows are displayed (what is the record count)?

:heavy_check_mark: 2573

b. What is the total revenue for the displayed rows?

:heavy_check_mark: 1770366

c. What is the average revenue for the displayed rows?

:heavy_check_mark: 688.06
##

### Scenario

While my first attempt to show the company's performance to Lucy was not bad, clearly she has a lot more requirements than what you provided. She wants to know more about the year over year sales, sliced into different categories, sub-categories, and countries. She also wants to see additional information such as customer demographics.

Jack has provided me with a different data source. This time the data has more than one hundred thousand rows. 

Before I can create additional reports to Lucy, first I need to prepare the data.

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/preparing_data.png)
##

### Lab 2

Download the lab file [Lab2Start v5.xlsx](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Lab2Start%20v5.xlsx) to answer the questions below.

The first thing I need to do is to convert the data into an Excel table.

Once I do that, I can add total row, filter the data, and the total will reflect the total only for the filtered data.

Add a total row for the table, and use the Sum aggregation to show the total of the **Revenue** column and then filter the data only for **United States**.
##

#### Question 1

What is the total revenue for all the sales in the **United States**?

:heavy_check_mark: 27975547
##

Now, what I need is add several columns, derived from existing columns in the data. Before adding columns, it is good practice to clear any filters you previously applied.

First, let's add a "Month" column. Insert a new column to the left of the **Customer ID** column, and use formula to derive the month of sales from the **Date** column.
##

#### Question 2

What is the total revenue for all the sales in the month of **December**?

:heavy_check_mark: 9086931
##

Next, let's add an "Age Group" column. Remember to clear any filters you previously applied. Insert a new column to the left of the **Customer Gender**, and use formula to derive the age group from the **Customer Age** column. Let's group the customers based on the following criteria:

- Youth (<25)
- Young Adults (25-34)
- Adults (35-64)
- Seniors (>64)
##

#### Question 3

What is the total revenue for all the sales for **Young Adults** Age Group?

:heavy_check_mark: 30655614
##

Now, let's add a "Frame Size" column. Insert a new column to the left of the** Order Quantity**, and use a combination of the IF() and RIGHT() functions to derive the frame size of a bicycle from the last two characters of the **Product** column, when the **Product Category** is **Bikes**. Otherwise, leave it blank.
##

#### Question 4

What is the total revenue for all the bikes with frame size **62** for the customer age group **Seniors**?

:heavy_check_mark: 12452
##

Last but not least, let's add a "Profit" column. Insert a new column to the right of the **Revenue**, and use formula to derive the Profit from both the **Revenue** and **Cost** columns. Show the total for the **Profit** column. Use the Sum aggregation in the total row of the table, for the **Profit** column.
##

#### Question 5

What is the total profit for **United States** sales in the month of **October 2015**, for customer age group **Adults**?

:heavy_check_mark: 138419
##

### Quiz time 

Download the quiz file [SalesForCourse_quizz_table.xlsx](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/SalesForCourse_quizz_table.xlsx) to answer questions 4 through 7 below.
##

#### Q1. Drag and Drop

You have the following Pivot Table fields. You need to calculate the number of sales transactions by age. Drag the appropriate field names and aggregated field names to the appropriate Pivot Table Areas (Filters, Rows, Columns, or Values) to perform the task

**Customer Age, Revenue, Sum of Customer Age, Count of Customer Age, Sum of Revenue**

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Q10.png)

:heavy_check_mark: Customer Age as Rows, Count of Customer Age as Values
##

#### Q2. Drag and Drop

You have the following Pivot Table fields. Your manager asks you to build a pivot table that shows revenue by state and country. How would you arrange the field names and aggregated field names to the appropriate Pivot Table Areas (Filters, Rows, Columns, or Values)?

**State, Country, Count of Revenue, Sum of Revenue**

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Q11.png)

:heavy_check_mark: Country, State as Rows, Sum of Revenue as Values
##

#### Question 3

You have a pivot that is based on an Excel table. You need to add new rows to the Excel table and want to be able to see the values for these new rows in the PivotTable. What should you do?

:heavy_check_mark: Add rows anywhere in the table, and refresh the pivot table.
##

#### Question 4

Using the data provided, create a pivot with “Country “as rows, “Month” as columns, and “Revenue” as values. Drag “Year” to the filter area. What does the value 286,779 in the first summarized field of the pivot table (uppermost left value) represent?

:heavy_check_mark: The aggregated value for January in France in 2015 and 2016.
##

#### Q5. Numerical Input

Create a new PivotTable with “Customer Gender” as rows, “Product Category” as columns, and “Revenue” as values.

a. How much revenue came from the purchase of bikes by women?

:heavy_check_mark: 5728781

b. How much revenue came from clothing purchased by men? 

:heavy_check_mark: 1817262
##

#### Question 6

Create a new PivotTable with “Country” as rows, “Product Category” as columns, “Year” as a filter, and “Revenue” as values. Filter by year 2015.

Drill down into the data for the cell that represents the revenue for bike sales in Germany, 2015. A new worksheet is created with a table in it. Which option below best describes the content of this table?

:heavy_check_mark: The table contains all rows in SalesTable corresponding to 2015, Bikes and Germany
##

#### Q7. Numerical Input

a. Total revenue from bike sales from the country with the highest revenue from bike sales in 2016:

:heavy_check_mark: 2257187

b. Number of sales transactions for bikes in the US in 2016:

:heavy_check_mark: 1415

c. Average number of items per transaction for clothing purchased in the US during 2016 rounded to the nearest whole number:

:heavy_check_mark: 2

d. The largest revenue value registered in a single transaction in the Germany during 2015:

:heavy_check_mark: 3089

e. The largest revenue value registered in a single transaction during 2015:

:heavy_check_mark: 3224
##

### Scenario

Now I have prepared the data in an Excel table, I can start to create pivot tables to aggregate the data and create some reports. 

From conversation with Lucy, I know that she is interested in looking into the yearly sales data broken down by countries, product categories, and age groups.

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/creating_report.png)
##

### Lab 3A

Download the lab file [Lab3Start v5.xlsx](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Lab3AStart%20v5.xlsx) to answer the questions below.

First, let's start by naming the Excel table. Name the Excel table prepared in the previous lab to **SalesTable**. From now                 ou add a pivot table, it should be based on this SalesTable, unless mentioned otherwise.

Now, proceed to add your first pivot table. Insert a new pivot table based on the SalesTable to a new sheet. Arrange the layout so that the pivot table displays the **Product Category** and **Sub Category** in the **Rows**, **Year** in the **Columns**, and **Revenue** (Sum of) as the **Values**.
##

#### Question 1

Which year did the company start selling **Touring Bikes**?

:heavy_check_mark: 2013
##

Insert another pivot table to the same sheet, next to the existing pivot table. Arrange the layout so that the pivot table displays the **Country** and **State** in the **Rows**, **Year** in the **Columns**, and **Revenue** (Sum of) as the **Values**. Sort the pivot table by **Sum of Revenue** so that the **Country** and **State** with the highest revenue is displayed first.
##

#### Question 2

Rank the **States** for **Germany**, from the highest to lowest revenue.

**Saarland, Nordrhein-Westfalen, Hessen, Brandenburg, Hamburg, Bayern**

:heavy_check_mark: 1. Hessen

:heavy_check_mark: 2. Saarland

:heavy_check_mark: 3. Nordrhein-Westfalen

:heavy_check_mark: 4. Hamburg

:heavy_check_mark: 5. Bayern

:heavy_check_mark: 6. Brandenburg

##

#### Question 3

What about for the year 2013? Rank from the highest to lowest revenue. 

**Saarland, Nordrhein-Westfalen, Hessen, Brandenburg, Hamburg, Bayern**

:heavy_check_mark: 1. Saarland

:heavy_check_mark: 2. Hessen

:heavy_check_mark: 3. Nordrhein-Westfalen

:heavy_check_mark: 4. Hamburg

:heavy_check_mark: 5. Bayern

:heavy_check_mark: 6. Brandenburg
##

Let's add another pivot table. This time arrange the layout so that the pivot table displays the **Frame Size** in the **Rows** and **Revenue** (Sum of) as the **Values**. Hide the rows that do not have a Frame size (blank Frame size), then sort the pivot table by **Sum of Revenue** so that the **Frame size** with the highest revenue is displayed first.
##

#### Question 4

Which Frame Size has the highest revenue?

:heavy_check_mark: 42
##

Last but not least, add another pivot table with **Age Group** as the **Rows** and **Revenue** (Sum of) as the **Values**. You will learn how to custom sort the Age Group in the next module. But for now, sort the pivot table by **Sum of Revenue** so that the **Age Group** with the highest revenue is displayed first.
##

#### Question 5

Which **Age Group** has the lowest revenue?

:heavy_check_mark: Seniors
##

### Lab 3B

Now I can start adding some charts to the sheet.

First, I add pivot chart for the pivot table that shows yearly sales (revenue) by Country (the pivot table you created for question 2 in Lab 3A). Select a **Line** chart to display the yearly trend. Make sure that the **Years** are located in the X axis, the **Revenue** in the Y axis, and the **Countries** as categories.
##

#### Question 1

In this chart, you can clearly see the sales trend for each Country. Which country’s trend displays the most fluctuation in sales when compared to the other countries?

:heavy_check_mark:  Australia
##

Add another pivot chart for the pivot table that shows yearly sales (revenue) by Product Category (the pivot table you created for question 1 in Lab 3A). Select a Column chart to display the yearly sales by category so that the years are together.
##

#### Question 2

In which year does the Bikes category have the lowest revenue?

:heavy_check_mark: 2014
##

Add another pivot chart, this time for the pivot table that shows Revenue by Age Group (the pivot table you created for question 5 in Lab 3A). Select a **Pie** chart to display the proportion of each **Age Group** (remember the chart styles) with data labels, formatted to two decimal points.
##

#### Question 3

In this chart, you can clearly see the proportion of sales (revenue) for each **Age Group**. What is the proportion of sales (revenue) for **Young Adults**?

:heavy_check_mark: 35.95%
##

Add another pivot chart, this time for the pivot table that shows Revenue by Frame size (the pivot table you created for question 4 in Lab 3A). Select a **Bar** chart to display the order of revenue by **Frame size**. Sort the Y axis to show the **Frame size** that has the highest revenue on the top.
##

#### Question 4

In this chart, you can clearly see the order of sales (revenue) by Frame size. Which frame size had the lowest revenue?

:heavy_check_mark: 54
##

### Quiz time 

Download the quiz file [SalesForCourse_quizz_table.xlsx](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/SalesForCourse_quizz_table.xlsx) to answer Question 1 below.
##

#### Question 1

Using the data provided, create a pivot with “Country” as rows, “Product Category” as columns, and “Revenue” as values. Slice months by “April.” 

What does the grand total for revenue (2,200,490) represent?

:heavy_check_mark: Total combined revenue for April in 2015 and 2016
##

#### Question 2

Which objects can be sliced by slicers?

:heavy_check_mark: Pivot Charts

:heavy_check_mark: Pivot Tables

:heavy_check_mark: Tables
##

#### Question 3

When would you use both a slicer and a filter based on the same field?

:heavy_check_mark: When you want to search items and apply the results on multiple PivotTables
## 

#### Question 4

You have the following slicers.

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Q2.png)

Why are Bike Racks, Fenders and Gloves at the bottom of the Sub Category slicer in a lighter color?

:heavy_check_mark: During 2015, there were no sales for these sub categories.
##

#### Question 5

You have the following slicers.

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Q3.png)

What will happen if you click "Bike Racks"?

:heavy_check_mark: All Pivot Tables connected to both year and sub category slicers will be blank.
##

### Scenario

I have created several pivot tables and pivot charts for Lucy.

So far everything has been well received by her. However, she would like to have easier ways to slice and dice the reports and charts herself.

I sat down with Lucy, and come up with several different ways that Lucy could slice the data

- Year
- Country
- Customer Gender
- Age Group
- Product Category
- Sub Category
- Frame size

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/creating_dashboard.png)
##

### Lab 4A

Download the lab file [Lab4AStart v5.xlsx](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Lab4AStart%20v5.xlsx) to answer the questions below.

Start by adding a new sheet named **Dashboard**. Then move (Cut and Paste) the four charts that you have to that sheet. Arrange the charts as appropriate.
##

For easy reference, let's give the charts titles if they don't have any, or rename them as appropriate. Name the four charts as follows:

- Yearly Sales by Country
- Yearly Sales by Category
- Sales by Frame Size
- Sales by Age Group
##

You can now add slicers to the sheet. Select the Yearly Sales by Country chart, and add seven slicers corresponding to Year, Country, Customer Gender, Age Group, Product Category, Sub Category and Frame Size. Arrange the slicers as appropriate.
##

The next thing you need to do is to connect these slicers to the charts. Let's do this chart by chart.

1. Connect the **Yearly Sales by Country** chart to all slicers, except the **Year** slicer. In another words, disconnect the **Year** slicer from the **Yearly Sales by Country**
2. Connect the **Yearly Sales by Category** chart to the **Age Group**, **Country**, **Customer Gender**, and **Frame Size** slicers.
3. Connect the Sales by **Frame Size** chart to all slicers, except the **Frame Size** slicer.
4. Connect the Sales by **Age Group** chart to all slicers, except the **Age Group**.
##

For niceties, you can clear the **Gridlines** on the Dashboard,

Once I've done the above, I'am ready to present the dashboard to Lucy.
##

#### Question 1

A quick glance on the **Yearly Sales by Country** shows that **Australia** has an unusual trend compared to the other countries. Which year does **Australia** have the least sales (revenue)??

:heavy_check_mark: 2011
##

Create an additional pivot chart to show **Sales by Country** using **Pie** chart. Show percentages for each slice of the pie and connect the chart to all the slicers, except the **Country** slicer. Overall, Australia commands 25% of the company's total sales. But in some of the years, this proportion changes.
##

#### Question 2

What is the percentage of **Australia** sales (of total sales) in the year that it has the least sales (previous question)?

:heavy_check_mark: 28%
##

#### Question 3

Let's filter the charts by **Australia** using the **Country** slicer. What might be the cause of this trend?

:heavy_check_mark: The sharp fluctuations of bicycle sales.
##

Based on the previous answer, create an additional pivot chart to show **Sales by Category** using a **Pie** chart. Show percentages for each slice of the pie and connect the chart to all the slicers, except the **Category** slicer.

For the next two questions, filter the charts by **Australia** using the country slicer and play around with the **Year** filter. Notice for different years, the changes in composition of **Australia's** sales by **Category**.
##

#### Question 4

Based only on years where all three categories have sales (there are three slices in the pie) in Australia, which year does the **Bikes** category have the lowest proportion of sales?

:heavy_check_mark: 2014
##

#### Question 5

What is the percentage of **Bikes** sales (of total sales) in that year (based on the previous question)?

:heavy_check_mark: 59%
##

### Lab 4B

Create an additional pivot chart to show **Sales by Customer Gender** using **Pie** chart. Show percentages for each slice of the pie and connect the chart to all the slicers, except the **Customer Gender** slicer.
##

#### Question 1

Which statement best describes the proportion of sales by customer gender?

:heavy_check_mark: Sales are pretty much evenly split between male and female customers.
##

#### Question 2

Which Bike's frame size is more popular for each **Gender**? Place the **Frame size** into the appropriate box.

**38, 42, 44, 46, 48, 50**

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Q12.png)

Zone 1 :heavy_check_mark: 42

Zone 2 :heavy_check_mark: 38
##

What about Customer Gender vs age group? Right now the **Sales by Age Group** chart does not differentiate by **Gender**. Modify this chart to be a **Column** chart. Show the **Customer Gender** side-by-side for each age group. Ensure that the chart is connected to all slicers. Last but not least, sort the **Age Group** appropriately.
## 

#### Question 3

In **Australia**, which **Age Group** has more sales revenue to females than to males?

:heavy_check_mark: Young Adults (25-34)

:heavy_check_mark: Adults (35-64)
##

### Quiz time 

#### Question 1

You have a sales table that contains “Quantity,” “Unit Price,” and “Unit Cost.” How could you add the following calculations to be used in a PivotTable?

Total Cost?
 
:heavy_check_mark: Add a calculated column in the table (Total Cost = Unit Cost * Quantity).

Total Revenue?

:heavy_check_mark: Add a calculated column in the table (Total Revenue = Unit Price * Quantity).

Total Profit?

:heavy_check_mark: Add a Pivot Table calculated field (Total Profit = Revenue - Cost).

:heavy_check_mark: Add a calculated column in the table (Total Profit = Revenue - Cost).

Profit Margin?

:heavy_check_mark: Add a PivotTable calculated field: Profit Margin = (Revenue - Cost) / Revenue.
##

#### Question 2

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Q4.png)

Conditional formatting (Green – Yellow – Red Color Scale) was applied on the above PivotTable. What can you change to the formatting to better visualize possible outliers in the data?

:heavy_check_mark: Modify the percentiles for maximum and minimum values.
##

### Scenario

Lucy was impressed with the dashboard you created.

With the dashboard, she is able to narrow down her interest. Specifically, she is interested with the sales in Australia. She would like to perform simple profitability analysis on the product category and sub-category, for specific year. 

Furthermore, she wants to have the option of going deeper to product level.

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/profit_and_outliers.png)
##

### Lab 5

Download the lab file [Lab5Start v5.xlsx](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Lab5Start%20v5.xlsx) to answer the questions below.

I started by adding a new sheet named **Details**. Add a pivot table that shows **Order Quantity**, **Revenue**, and **Profit** by **Product Category**. Add the **Year** and **Country** field to the filter list.

Next, add the **Sub Category** field to the rows of the pivot table. Make sure the rows are fully expanded. Now you have quite a few numerical data shown in the pivot table, which might not be as easy to analyze as before. Apply** Data bars** conditional formatting to the **Order Quantity**, **Revenue** and **Profit** column on **Sub Category** level. Use blue, green, and yellow respectively.

Filter the pivot table for **Australia** and **2016**, and answer the following questions.
## 

#### Question 1

Which **Sub Category** sold the most quantity?

:heavy_check_mark: Tires and Tubes

#### Question 2

Which Sub Category has the most revenue?

:heavy_check_mark: Road Bikes
##

Now, add a calculated field to your pivot table, name the field **Margin** with the value derived from the **Profit** and **Revenue** column. Format the field as percentage with two decimal places.
##

#### Question 3

What is the total margin for Australia in the year 2016?

:heavy_check_mark: 33.74%
##

#### Question 4

Using the same filters, which category has the lowest margin?

:heavy_check_mark: Clothing
##

Now, apply Color scales conditional formatting to **Margin** for the **Sub Category** level.

:heavy_check_mark: Caps
##

#### Question 5

Which Sub Category has the lowest margin?

:heavy_check_mark:
##

Let's dig a little deeper. Add the **Product** field to the **Rows**, under **Sub Category**. Ensure that the **Year** and **Country** filter still select **2016** and **Australia** respectively. Apply **Color scales** conditional formatting only to the cells showing the **margin** for the **Product**.
##

#### Question 6

Which product has the least margin?

:heavy_check_mark: Short-Sleeve Classic Jersey, M
##

### Quiz time 

#### Question 1

What is the main advantage of referencing cells in a Pivot Table with the GetPivotData() function rather than by regular cell reference?

:heavy_check_mark: It returns the same values even if the layout of the data changes.
##

#### Question 2

What are the disadvantages of referencing cells in a Pivot Table with the GetPivotData() function and not by regular cell reference?

:heavy_check_mark: The function is not adjusted when it is dragged across cells.

:heavy_check_mark: It is very complicated to edit the function and to write it manually.
##

#### Question 3

What are the advantages of creating aggregates using SUMIFS() rather than using a Pivot Table? 

:heavy_check_mark: SUMIFS() formulas can be arranged in flexible formats in the grid and formulas' comparison operators are not restricted to equality.

:heavy_check_mark: SUMIFS() formulas do not require refreshes to reflect changes in the data.
##

#### Question 4

What are the disadvantages of creating aggregates using SUMIFS() rather than using a Pivot Table? 

:heavy_check_mark: A range with SUMIFS formulas is much more difficult to change if you need to include more or fewer fields, whereas Pivot Tables shrink and expand dynamically based on the data.

:heavy_check_mark: On a report based on SUMIFS() formulas, you are unable to implement simple drill downs and drill ups.
## 

### Scenario

Lucy seems intrigued your presentation on Australia's sales.

Now, she wants to review the sales difference (growth) for Product Category and Sub Category between year 2015 and 2016. She wants to be able to filter the report by all the options available, including Customer Gender and Age Group.

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/yearly_trend.png)
##

### Lab 6

Download the lab file Lab6 [Startv5.xlsx](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Lab6Start%20v5.xlsx) to answer the questions below.

Now, I will create two reports. One with the sales for each year and the other with the % Change in Revenue from Year to Year. Each report is a cross-tabular format with **Product Categories** and **Sub categories** on the rows and **Year** on the columns, with **Sum of Revenue** as the aggregate data. When you are done, neither report will be based on a Pivot Table. However, you will enable a Pivot Table filter to allow the selection of **Country**, **Customer Gender**, and **Age Group** for these reports. Visually, you would like to have something like this:
##

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Q5.png)
##

Let's start by adding a new sheet named **Growth**. Starting in cell B1, add a new pivot table using SalesTable as the source with **Country**, **Customer Gender**, and **Age Group** as Filters, **Product Category** and **Sub Category** as Rows, **Year** as Columns, and **Sum of Revenue** as Values.

Now, copy the structure of the rows and columns to the cells next to the pivot table, that is, copy the **Years** labels in C6:H6 (excluding the Grand Total) to range L6:Q6, and the Product categories and Sub categories labels in B7:B26 (excluding the Grand Total) to range K7:K26. Highlight and copy range B6:H26, place your cursor in cell K6, right click and from the Paste Options, select Formatting (R).

Alternatively, you can select B6:H26, place your cursor in cell K6, right click and from the Paste Options, select Values (V) and then, once again right click in cell K6, and from the Paste Options, select Formatting (R). After which, delete the data area.
##

Populate the data area of the newly created cross tabular structure using SUMIFS() function. Assuming that you have:

- K7: Product Category eg "Accessories"
- Q6: Year eg. "2016"
- C1: Country Filter
- C2: Customer Gender Filter
- C3: Age Group Filter
##

Then, the cell that corresponds to **Year 2016** for the **Product Category Accessories** would have the following formula

```bash
=SUMIFS(SalesTable[Revenue], 
SalesTable[Product Category],$K7, 
SalesTable[Year],Q$6, 
SalesTable[Country],IF($C$1="(All)","*",$C$1), 
SalesTable[Customer Gender],IF($C$2="(All)","*",$C$2), 
SalesTable[Age Group],IF($C$3="(All)","*",$C$3))
```

##
Confirm that your filter boxes (cells C1, C2 and C3) contain the same value that is used in the formula (All). If not, then replace the “(All)” with the exact value found in your filter boxes (it could be All or your language’s word for All. (Note that you will have to confirm these values for the next set of formulas as well)

Now, copy this formula (using Ctrl C keyboard combination) and paste it (using Ctrl V keyboard combination) into the other Years and Product Category cells for all three categories (Accessories, Bikes, Clothing). When you are done, your calculated values in cells L7:Q7, L16:Q16 and L20:Q20 should match their respective pivot table values.
##

In addition, assuming that you have:

- K8: Sub Category eg "Bike Racks"
- Q6: Year eg. "2016"
- C1: Country Filter
- C2: Customer Gender Filter
- C3: Age Group Filter
##

Then, the cell that corresponds to **Year 2016** for the **Sub Category Bike Racks** would have the following formula

```bash
=SUMIFS(SalesTable[Revenue], 
SalesTable[Sub Category],$K8, 
SalesTable[Year],Q$6, 
SalesTable[Country],IF($C$1="(All)","*",$C$1), 
SalesTable[Customer Gender],IF($C$2="(All)","*",$C$2), 
SalesTable[Age Group],IF($C$3="(All)","*",$C$3))
```
##
Now, copy this formula and paste it into the other Years and Sub Category cells. When I finish it, my calculated values in these cells should match their respective pivot table values.

Once my table is populated, I remove the fields from the Rows, Columns, and Values of the Pivot table, so that only the filters remains. Align the filters with my cross tabular report and add sparklines (to show the series high point) next to the column containing the last year of data.
##

Upon completion the first cross tabular report, I will create another one with the same structure next to it (excluding the 2011 column). This time, for % Change in Revenue, which is essentially the difference between the two years, divided by the previous year. For example, one formula could be:

=IF(C7<>0,(D7-C7)/C7,0)

Present the data as percentage without decimal places. Apply an **Icon Sets** conditional formatting to differentiate positive and negative growth. Edit your conditional formatting icon set rules to display the icons similar to the rule below:
##

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Q6.png)

When you're done, you should have something like this (the below is just an example, your values may not necessarily match):

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Q7.png)
##

#### Question 1

Without applying any filter, which year does the **Accessories** category have negative growth?

:heavy_check_mark: 2015
##

#### Question 2

Filter the report for **Youth Age Group**. Which two years do the **Accessories** category have negative growth?

:heavy_check_mark: 2014, 2016
##

#### Question 3

Without applying any filter, which two years do the **Bikes** category have negative growth?

:heavy_check_mark: 2014, 2016
##

#### Question 4

Filter the report for **Australia**. Which year do the **Bikes** category have the highest growth?

:heavy_check_mark: 2015
##

#### Question 5

Keep the Australia filter. In the year that **Bikes** sales have the highest growth (previous question), which **Sub Category** of **Bikes** has the highest growth?

:heavy_check_mark: Mountain Bikes
##

### Quiz time

#### Question 1

How would you modify the following formula to include age ranges in the following report?

```bash
=SUMIFS(SalesTable[Revenue],SalesTable[Customer Gender],B$4,SalesTable[Country],$A5)
```
![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Q8.png)

:heavy_check_mark:  

```bash
=SUMIFS(SalesTable[Revenue],SalesTable[Customer Gender],B$4,SalesTable[Country],$A5,SalesTable[Customer Age],">="&$B$1,SalesTable[Customer Age],"<="&$B$2)
```
##

#### Question 2

What is the main purpose of using the Treemap and Sunburst charts to display data? 

:heavy_check_mark: To show the relative size at different levels of a hierarchy.
##

#### Question 3

You want to create a SUMIFS() formula that references a specific cell in your spreadsheet as you copy it to a range of data containing multiple columns. What should you do?

:heavy_check_mark: Use a mix of absolute and relative reference in your formula.
##

### Scenario

Upon reviewing the growth report I created, (download the [file](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Lab7Start%20v5.xlsx) here), Lucy asked for a report that shows composition of Product Categories and Sub Categories based on certain filters, including Year, Country, Customer Gender, and Age Group.

Specifically, Lucy wants to see the report visualized using a hierarchical chart.

![App Screenshot](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/hierarchy.png)
##

### Lab 7

Download the lab file [Lab7Start v5.xlsx](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Lab7Start%20v5.xlsx) to answer the questions below.

Let's add another sheet to the workbook and name it **Composition**.

Start by adding a new pivot table with **Year**, **Country**, **Customer Gender**, and **Age Group** as Filters,** Product Category** and **Sub Category** as Rows, and **Sum of Revenue** as Values.

Format the pivot table to show in **Tabular** form, **do not show subtotals** and **grand totals**, and **repeat all item labels**.

Now, copy the structure of the rows and columns to the cells next to the pivot table, that is the **Product categories** and **Sub categories** as Rows. Delete the data area for now.
##

Populate the data area of the newly created cross tabular structure using SUMIFS() function. Assuming that you have:

- F7: Product Category eg. "Accessories"
- G7: Sub Category eg. "Bike Racks"
- C1: Year filter
- C2: Country Filter
- C3: Customer Gender Filter
- C4: Age Group Filter
##

Then, the cell that corresponds the Product Category Accessories and Sub Category Bike Racks would have the following formula:

```bash
=SUMIFS(SalesTable[Revenue], 
SalesTable[Product Category],$F7, 
SalesTable[Sub Category],$G7, 
SalesTable[Year],IF($C$1="(All)",">0",$C$1), 
SalesTable[Country],IF($C$2="(All)","*",$C$2), 
SalesTable[Customer Gender],IF($C$3="(All)","*",$C$3), 
SalesTable[Age Group],IF($C$4="(All)","*",$C$4)) 
```
##
Once your table is populated, remove the fields from the Rows, Columns, and Values of the Pivot table, so that only the filters remains. Align the filters with your cross tabular report. Create a treemap or sunburst chart based on the data in the cross tabular report. Name the chart "Sales Composition".
##

#### Question 1

Explore the sales composition of **Bikes** category for each **Age Group**. Which **Age Group** does the composition (rank of sales) differ than the rest?

:heavy_check_mark: Seniors
##

#### Question 2

Now explore the sales composition of **Bikes** category for each **Age Group**, for the **Male** customers. Which **Age Group** does the composition (rank of sales) differ than the rest?

:heavy_check_mark: The composition are the same across Age Group for Male customers.
##

#### Question 3

Clear all filters. Now, filter for the year **2016** and **Germany**. Rank the sales from the highest to lowest for the **Clothing** category. 

**Vests, Caps, Jerseys, Shorts, Gloves, Socks**

:heavy_check_mark: 1. Jerseys

:heavy_check_mark: 2. Vests

:heavy_check_mark: 3. Caps

:heavy_check_mark: 4. Gloves

:heavy_check_mark: 5. Shorts

:heavy_check_mark: 6. Socks
##

#### Question 4

Keep the filter settings and add filter by **Male** customers. Rank the sales from the highest to lowest for the **Clothing** category.

**Vests, Caps, Jerseys, Shorts, Gloves, Socks**

:heavy_check_mark: 1. Jerseys

:heavy_check_mark: 2. Vests

:heavy_check_mark: 3. Gloves

:heavy_check_mark: 4. Caps

:heavy_check_mark: 5. Shorts

:heavy_check_mark: 6. Socks
##

#### Question 5

Keep the filter settings and add filter by **Youth Age Group**. Rank the sales from the highest to lowest for the **Clothing** category.

**Vests, Caps, Jerseys, Shorts, Gloves, Socks**

:heavy_check_mark: 1. Jerseys

:heavy_check_mark: 2. Gloves

:heavy_check_mark: 3. Caps

:heavy_check_mark: 4. Vests

:heavy_check_mark: 5. Socks

:heavy_check_mark: 6. Shorts
##

### Quiz time

#### Question 1

You want combine data in two ranges/tables in a single pivot. What should you do?

:heavy_check_mark: Click ‘More Tables…’ from the pivot field list, and create relationships between the different ranges.

:heavy_check_mark: Combine the different ranges into one range using VLOOKUP formulas.
##

### Scenario

Lastly, Lucy wants to have more information about the customer demographics in addition to the already available age and gender. Since my data (download [file here](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Lab8Start%20v5.xlsx)) has Customer ID for each row, you can "connect" these rows with your customer demographics database. My customer demographics is stored in an Access database, which you can download [here](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Demographics.mdb).

Workaround: Alternatively, the customer demographics is also provided as a txt file [here](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Customer_demographics.txt), in case you could not work with Access database.
##

### Lab 8

Download a file [Lab8Start v5.xlsx](https://github.com/SomonOlimzoda/DataAnalysisExcel/blob/main/Lab8Start%20v5.xlsx). 

This Excel file contains two charts that. The Composition worksheet contains a **Treemap** and **Sunburst** chart. 

1. From a new worksheet, go to the **Data** tab and on the **Get External Data** group, select the **From Access** option. This will bring up the **Select Data Source** dialog box.

2. Navigate to where you saved the **Demographics.mdb** file, select it and click on the **Open** button.

3. From the **Import Data** dialog box, ensure that **Table** is selected for how you want to view the data, within the existing worksheet (or select New Worksheet if you didn’t start from a new one), then click on the **OK** button.

4. There is only one table in this database so the import is rather simple. Rename the worksheet **Demographic Data**. Your import should result in 18,484 rows (plus the header) and seven columns of data.

5. Create a new pivot table **from the SalesTable** then click **More tables**. Remember to select **Yes** to Excel’s prompt to create a new pivot table.

6. Now, add the **MaritalStatus** and **TotalChildren** fields from the **Customer_demographic** table as **Rows** and **Order Quantity** from the **SalesTable** as **Values**. In addition, add the **Product Category** as filters.  
##

#### Question 1

For those customers who bought bikes, what are the top **three** (bought the most quantity) customer profiles (marital status and number of children)?

:heavy_check_mark: Single without child

:heavy_check_mark: Married with one child

:heavy_check_mark: Married with two children
##

Now, remove the MaritalStatus and TotalChildren fields from the Rows and replace them with the YearlyIncome field from the Customer_demographic table.
##

#### Question 2

What are the three income brackets for customers who bought the most bikes (quantity), in order from the income bracket with the highest number of sales (in Zone 1) to income bracket with the lowest number of sales (in Zone 3)?

**30000, 40000, 50000, 60000, 70000**

:heavy_check_mark: Zone 1: 40000

:heavy_check_mark: Zone 2: 60000

:heavy_check_mark: Zone 3: 70000
##

Now, remove the YearlyIncome field from the Rows and replace it with the EnglishEducation field from the Customer_demographic table.
##

#### Question 3

For those customers who bought bikes, what are the top **two** (bought the most quantity) education levels?

:heavy_check_mark: Partial College

:heavy_check_mark: Bachelors
##

Lastly, remove the **EnglishEducation** field from the Rows and replace it with the **HouseOwnerFlag** field from the **Customer_demographics** table. Format the **Sum of Order Quantity** to show as Percentage of Grand Total with two decimal places.
##

#### Question 4

What is the percentage of the customers who bought **Bikes** and are house owners?

:heavy_check_mark: 68.72%



















