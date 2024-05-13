Putting data in Tables is among the simplest ways to be efficient in Excel. It allows you to place column headers, which allow for a function to act cohesively on a group of cells. 

*Eg*

= UPPER(Table1[#Headers]] )

>This will create a upper case column headers of the table headers in the cell you wrote the function in.

**The same can be extended to table footers, even the table itself.**

![image](https://github.com/Glen-Ochieng/Useful-Excel-Functions-for-Data-Analysis./assets/155974295/fcdd732b-b23a-466d-8ebf-2b8f908cc1e6)

You can also name the tables to make the referencing more meaningful.

![image](https://github.com/Glen-Ochieng/Excel-for-Data-Analysis./assets/155974295/a72407f4-74cc-4812-b3cd-be1b6ca00701)

Naming the table enables you to perform functions on the table even if more rows are added without the altercations causing an issue in totals. Plus it makes the functions look cleaner. 

*To reference the table and a specific column you cite it inside square brackets*

= SUM(sales[fresh])

>This will sum all entries in the fresh column in the sales table.


**Referring to data by object name instead of cell location minimizes potential formula issues arising from changing the table's size and placement. Tables become crucial in preventing problems like missing data in a PivotTable when new rows are added.** 


# Creating a new column from existing rows in a table

If you need to create a new column from the ratio of two columns then using @ is quite handy.

Assume you have two columns in a table named bill_length and bill_depth, then you want to create a column called bill_ratio which divides the two. 

*Syntax*

> =[@bill_legtnh]/[@bill_depth]

*The square brackets are becasue you are referencing a table entry and the @ tells excel to perform the division row by row*
