// finance_raw: File Finance gốc (chỉ tạo connection, không nạp vào Data Model).
// Nạp: Only Create Connection
let
    Source = Excel.Workbook(
        File.Contents(SourceFolder & "finance_revenue.xlsx"), null, true),
    Navigation = Source{[Item = "tbl_finance_revenue", Kind = "Table"]}[Data],
    ChangedType = Table.TransformColumnTypes(Navigation, {
            {"posting_date", type date},
            {"fin_store_code", type text},
            {"fin_channel", type text},
            {"entry_type", type text},
            {"revenue_recognized", Currency.Type},
            {"refunds", Currency.Type},
            {"net_revenue", Currency.Type},
            {"vat_amount", Currency.Type},
            {"total_incl_vat", Currency.Type}
        })
in
    ChangedType
