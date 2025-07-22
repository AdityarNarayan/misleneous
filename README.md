=AND(
  ISERROR(--A1), 
  LOWER(A1)<>"total",
  LOWER(A1)<>"summary",
  LOWER(A1)<>"n/a",
  COUNTIF($A$1:$A$100, A1) > 1
)
