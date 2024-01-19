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

## Countifs
Used to count and simulateneously check a condition of a particular range

*Unlike sumsifs() , here don't include the range to count/sum . Just start with the criteria*

=countifs(rangetocheck,criteriatocheck)

*Criteria to check can increase but the range to count remains one*
=countifs (rangetocheck1,criteriatocheck1, rangetocheck2,criteriatocheck2,etc)

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

## VLOOKUP
It looks up the value or a string in a table and returns what you specify.It is used to when you want to find the column associated with a cell say you have an id cell and you want to find information associated with that id that's located in a certain column. 

*Syntax*

=VLOOKUP(lookupvalue- the cell, table array- this will be multiple columns,column index number- the column number in the table array of the information you want, range lookup- normally true or false statement but usually use false to write the exact match, true returns an approximate match)

*=vlookup*(E2,E:G,3,False)*
 
