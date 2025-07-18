let
    Source = Excel.CurrentWorkbook(){[Name="tbl_Data"]}[Content],
    FilteredRows = Table.SelectRows(Source, each [Customer Name] = CustomerName)
in
    FilteredRows
