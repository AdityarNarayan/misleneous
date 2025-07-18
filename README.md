=LET(
    cat, "ACH In",
    cust, CustomerFilter,
    ben, FILTER(Data!AO2:AO10000, (Data!L2:L10000=cat)*(Data!AN2:AN10000=cust)),
    benU, UNIQUE(ben),
    benAmt, BYROW(benU, LAMBDA(b, SUMIFS(Data!K2:K10000, Data!AO2:AO10000, b, Data!L2:L10000, cat, Data!AN2:AN10000, cust))),
    sortIdx, SORTBY(SEQUENCE(ROWS(benU)), -benAmt),
    benU2, INDEX(benU, sortIdx),
    benAmt2, INDEX(benAmt, sortIdx),
    runTot, SCAN(0, benAmt2, LAMBDA(a, b, a+b)),
    tot, SUM(benAmt2),
    cumPct, runTot/tot,
    flag, IF(cumPct<=0.5, "Include", ""),
    result, HSTACK(benU2, benAmt2, runTot, cumPct, flag),
    IFERROR(result, "")
)
