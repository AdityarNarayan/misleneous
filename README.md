=IF(OR(L2="ACH IN", L2="Checks Deposited", L2="Wires IN"),
    IF(U2<>"", U2, IF(AM2<>"", AM2, LEFT(M2, 15))),
    IF(OR(L2="ACH OUT", L2="Checks Issued", L2="Wires Out"),
        IF(Y2<>"", Y2, IF(AM2<>"", AM2, LEFT(M2, 15))),
        IF(OR(L2="Internal Transfer IN", L2="Internal Transfer Out", L2="Cash In", L2="Cash Out", L2="Other In", L2="Other Out"),
            IF(AM2<>"", AM2, LEFT(M2, 15)),
            "OOS category"
        )
    )
)
