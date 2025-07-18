=LET(
    ben, UNIQUE(FILTER(Data!AO2:AO10000, (Data!L2:L10000="ACH In")*(Data!AN2:AN10000=CustomerFilter))),
    benAmt, BYROW(ben, LAMBDA(b, SUMIFS(Data!K2:K10000, Data!AO2:AO10000, b, Data!L2:L10000, "ACH In", Data!AN2:AN10000, CustomerFilter))),
    sorted, SORT(HSTACK(ben, benAmt), 2, -1),
    runTot, SCAN(0, INDEX(sorted,,2), LAMBDA(a, b, a+b)),
    tot, SUM(INDEX(sorted,,2)),
    cumPct, runTot/tot,
    flag, IF(cumPct<=0.5, "Include", ""),
    HSTACK(INDEX(sorted,,1), INDEX(sorted,,2), runTot, cumPct, flag)
)
