// dim_date: Bảng dim: nạp và đặt kiểu.
// Nạp: Only Create Connection + Add this data to the Data Model
let
    Source = Excel.Workbook(
        File.Contents(SourceFolder & "Nexa_Sales_P1_Dataset.xlsx"), null, true),
    Navigation = Source{[Item = "tbl_dim_date", Kind = "Table"]}[Data],
    ChangedType = Table.TransformColumnTypes(Navigation, {
            {"date", type date},
            {"date_key", Int64.Type},
            {"year", Int64.Type},
            {"quarter", type text},
            {"month", Int64.Type},
            {"month_id", Int64.Type},
            {"month_label", type text},
            {"day", Int64.Type},
            {"weekday_no", Int64.Type},
            {"weekday_name", type text},
            {"is_weekend", type logical},
            {"has_data", type logical},
            {"month_end_date", type date}
        })
in
    ChangedType
