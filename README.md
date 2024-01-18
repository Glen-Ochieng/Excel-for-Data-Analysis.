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


## Countifs
Used to count and simulateneously check a condition of a particular range

*Unlike sumsifs() , here don't include the range to count/sum . Just start with the criteria*

=countifs(rangetocheck,criteriatocheck)

*Criteria to check can increase but the range to count remains one*
=sumifs (rangetocheck1,criteriatocheck1, rangetocheck2,criteriatocheck2,etc)
