Putting data in Tables is among the simplest ways to be efficient in Excel. It allows you to place column headers, which allow for a function to act cohesively on a group of cells. 

*Eg*

= UPPER(Table1[#Headers]] )

>This will create a upper case column headers of the table headers in the cell you wrote the function in.

<b>The same can be extended to table footers, even the table itself.

![image](https://github.com/Glen-Ochieng/Useful-Excel-Functions-for-Data-Analysis./assets/155974295/fcdd732b-b23a-466d-8ebf-2b8f908cc1e6)

You can also name the tables to make the referencing more meaningful.

![image](https://github.com/Glen-Ochieng/Excel-for-Data-Analysis./assets/155974295/a72407f4-74cc-4812-b3cd-be1b6ca00701)

Naming the table enables you to perform functions on the table even if more rows are added without the altercations causing an issue in totals. Plus it makes the functions look cleaner. 

*To reference the table and a specific column you cite it inside square brackets*

= SUM(sales[fresh])
