=LET(
    names, FILTER(A11:A100, E11:E100="Include"),
    credits, FILTER(B11:B100, E11:E100="Include"),
    n, ROWS(names),
    summary, HSTACK(names, credits),
    grandRow, {"Grand Total", SUM(credits)},
    VSTACK(summary, grandRow)
)
