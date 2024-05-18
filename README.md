# Useful-Excel-Functions

## If function 

=if(E2>10,1,0)

*If two conditions both must be met*

=if(and(E2>10,G3<56),1,0)

*If either of the conditions must be met*

=if(or(E2>10,G3<56),1,0)

*If a condition has not been met*

=if(not(F2="PG"),1,0) 

*NB*
Strictly use double quotation marks and not single quotation marks!!

## Sumifs
Used to sum and simulateneously check a condition of a particular range

=sumifs(rangetosum, rangetocheck,criteriatocheck)

*Criteria to check can increase but the range to sum remains one*
=sumifs(rangetosum, rangetocheck1,criteriatocheck1, rangetocheck2,criteriatocheck2,etc)

*=SUMIFS(G2:G8808,E2:E8808,">1990",E2:E8808,"<2000")*

*=SUMIFS(G2:G8808,E2:E8808,">1990",E2:E8808,"<2000",B2:B8808,"Movie")*

## Sortby

*Syntax*
Used to sort a range by a certain condition or conditions.

=sortby(rangetosort, rangetocheck1,ascending/descending-order1, rangetocheck2,ascending/descending order 2,etc)

![image](https://github.com/Glen-Ochieng/Excel-for-Data-Analysis./assets/155974295/337e1a3a-dc34-4d70-9642-34e6a02e53f7)


## Countifs
Used to count and simulateneously check a condition of a particular range

*Unlike sumsifs() , here don't include the range to count/sum . Just start with the criteria*

=countifs(rangetocheck,criteriatocheck)

*Criteria to check can increase but the range to count remains one*
=countifs (rangetocheck1,criteriatocheck1, rangetocheck2,criteriatocheck2,etc)

## COUNTA
Unlike the function COUNT(), COUNTA() counts **not only numbers but also texts**. Thus COUNTA() counts all non-empty cells in a range.   

## UNIQUE
Shows unique entries in that column.

*Syntax*

UNIQUE(A:A)

## LEFT
It will extract a specified number of characters starting from the left hand side.

*Syntax*

=LEFT(cell, specified no. of characters)
## RIGHT
It will extract a specified number of characters starting from the right hand side.

*Syntax*

=RIGHT(cell, specified no. of characters)


## TRIM
Removes unwanted spaces from a cell without removing the spaces between the words in that cell.

*Syntax*

=TRIM(A1)

## FILTER

Filters an array based on a certain condition(s). The condition could be a logical expression, value to match or a formula. 

*Syntax*

=FILTER(cell_range, condition, [if empty_then do this])

*Example*

=FILTER(A:G,A="Sticky Notes", )

![image](https://github.com/Glen-Ochieng/Excel-for-Data-Analysis./assets/155974295/1383c0bc-0a0d-4dbe-bbe9-4e895f9e0f38)

**However, the two conditions must be enclosed in brackets**

![image](https://github.com/Glen-Ochieng/Excel-for-Data-Analysis./assets/155974295/04235123-e0bf-499a-b0bd-691addd550e3)


## VLOOKUP
It looks up the value or a string in a table and returns what you specify.It is used to when you want to find the column associated with a cell say you have an id cell and you want to find information associated with that id that's located in a certain column. 

*Syntax*

=VLOOKUP(lookupvalue- the cell, table array- this will be multiple columns,column index number- the column number in the table array of the information you want, range lookup- normally true or false statement but usually use false to write the exact match, true returns an approximate match)

*=vlookup*(E2,E:G,3,False)*


## XLOOKUP

A replacement to the VLOOKUP 

![image](https://github.com/Glen-Ochieng/Excel-for-Data-Analysis./assets/155974295/4256fa83-4f9b-49db-bf5a-96879072f78d)


*Example*

![image](https://github.com/Glen-Ochieng/Excel-for-Data-Analysis./assets/155974295/26d7e5f1-077c-4bfd-9cf1-46fab9db5ee4)


 =XLOOKUP(transactions[product_id],products[product_id],products[product_name])

This looks up the corresponding values matched to the product_id in from the transactions table to the products table product_id and returns the entries in the product name column

## VSTACK

This allows vertical appending of arrays. This is particulartly useful for merging data from multiple sheets. 

#### NB

This function is only avaliable in Excel 365, but you can work around it by using Google Sheets.

*Syntax*

=VSTACK(range1, range2)  eg =vstack(a2:e10,a12:e35)

![image](https://github.com/Glen-Ochieng/Excel-for-Data-Analysis./assets/155974295/fbd41f9d-1976-4f6f-a8bd-f1e39e8fcf4a)

## CONVERT
Transforms a time variable from one format to another format, i.e days to hours, hours to days, hours to mins.

*Assume cell B2 = 6, whereby 6 represents 6 hours* 

To change from hour to day 

=CONVERT(B2,"hr","day")

To change from hour to min

=CONVERT(B2, "hr", "min")

To change from hour to year
=CONVERT(B2, "hr", "yr")

### NB: For conversion to a year consider a year has 365.25 days


## DATE
Converts three separate values into a date.

*Syntax*

DATE(year,month,day)

Assume A2=3, A3=11, A4=2012.

The formula to get a date would be:
=DATE(A4,A3,A2)

Result = 3/14/2012
