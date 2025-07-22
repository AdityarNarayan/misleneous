=VSTACK( {"Customer_ID","Sum of Credit"}, IF( COUNTIF(Sheet1!E3:E106, "Include") > 5, FILTER(Sheet1!A3:B101, Sheet1!E3:E106="Include"), TAKE(SORT(Sheet1!A3:B101, 2, -1), 5) ), A1:B1 )
