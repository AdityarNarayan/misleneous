1. Prepare Your Summary Table Layout
Let’s say you want your summary table to start at Results!A2 with these headers in row 1:



A1
B1
C1
D1
E1




Beneficiary Name
Total Credit
Running Total
Cumulative %
Flag




2. List Unique Beneficiary Names (Column A)
In Results!A2, enter:
Excel formulae=UNIQUE(FILTER(Data!AO2:AO10000, (Data!L2:L10000="ACH In")*(Data!AN2:AN10000=CustomerFilter)))

This lists all unique beneficiary names for the selected customer and category.

3. Calculate Total Credit per Beneficiary (Column B)
In Results!B2, enter:
Excel formulae=SUMIFS(Data!K2:K10000, Data!AO2:AO10000, A2, Data!L2:L10000, "ACH In", Data!AN2:AN10000, CustomerFilter)

Drag this formula down alongside your list in column A.

4. Sort Beneficiaries by Total Credit (Optional)
If you want to sort by total credit, you can use SORT:
Excel formulae=SORT(A2:B100, 2, -1)

But for simplicity, you can skip this step for now.

5. Compute Running Total (Column C)
In Results!C2, enter:
Excel formulae=SUM($B$2:B2)

Drag this formula down.

6. Compute Cumulative Percentage (Column D)
First, calculate the grand total in a separate cell, e.g., Results!F1:
Excel formulae=SUM(B2:B100)

Then in Results!D2, enter:
Excel formulae=C2/$F$1

Drag down.

7. Flag Top 50% (Column E)
In Results!E2, enter:
Excel formulae=IF(D2<=0.5, "Include", "")
