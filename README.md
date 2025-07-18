Step-by-Step: Results Table for "ACH IN" Only
1. Unique Beneficiaries for Selected Customer and "ACH IN"
In Results!A5:
Excel formulae=UNIQUE(FILTER(Data!C2:C1000, (Data!A2:A1000="ACH IN")*(Data!B2:B1000=CustomerName)))


2. Total Credit per Beneficiary
In Results!B5:
Excel formulae=SUMIFS(Data!D$2:D$1000, Data!C$2:C$1000, A5, Data!A$2:A$1000, "ACH IN", Data!B$2:B$1000, CustomerName)

Copy down.

3. Sort Beneficiaries by Total Credit (Optional)
If you want to sort the table by total credit descending, use:
Excel formulae=SORTBY(
  UNIQUE(FILTER(Data!C2:C1000, (Data!A2:A1000="ACH IN")*(Data!B2:B1000=CustomerName))),
  --(LAMBDA(x, SUMIFS(Data!D$2:D$1000, Data!C$2:C$1000, x, Data!A$2:A$1000, "ACH IN", Data!B$2:B$1000, CustomerName))(
    UNIQUE(FILTER(Data!C2:C1000, (Data!A2:A1000="ACH IN")*(Data!B2:B1000=CustomerName)))
  )),
  -1
)

Or, use the helper columns and sort manually.

4. Running Total and Cumulative %
C5 (Running Total):
Excel formulae=SUM($B$5:B5)

Copy down.
D5 (Cumulative %):
Excel formulae=C5/SUM(B$5:B$100)

Copy down.

5. Flag Beneficiaries Comprising ≤ 50%
E5:
Excel formulae=IF(D5<=0.5,"Include","")

Copy down.

6. (Optional) Show Only "Include" Beneficiaries
To display only those beneficiaries, use:
Excel formulae=FILTER(A5:E100, E5:E100="Include")
