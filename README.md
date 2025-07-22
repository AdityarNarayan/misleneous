=SUM(
    IF(
        COUNTIF(Sheet1!E3:E106, "Include") > 5,
        FILTER(Sheet1!B3:B101, Sheet1!E3:E106="Include"),
        TAKE(SORT(Sheet1!B3:B101, 1, -1), 5)
    )
)
