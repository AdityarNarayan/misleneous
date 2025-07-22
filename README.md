AND(
  NOT(ISNUMBER(A1)),
  ISERROR(MATCH(LOWER(A1), {"total","summary","n/a"}, 0)),
  COUNTIF($A$1:$A$100, A1) > 1
)
