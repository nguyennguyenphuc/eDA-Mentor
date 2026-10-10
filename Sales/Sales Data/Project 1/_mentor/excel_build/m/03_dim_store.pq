// dim_store: Bảng dim: nạp và đặt kiểu.
// Nạp: Only Create Connection + Add this data to the Data Model
let
    Source = Excel.Workbook(
        File.Contents(SourceFolder & "Nexa_Sales_P1_Dataset.xlsx"), null, true),
    Navigation = Source{[Item = "tbl_dim_store", Kind = "Table"]}[Data],
    ChangedType = Table.TransformColumnTypes(Navigation, {
            {"store_key", Int64.Type},
            {"store_code", type text},
            {"store_name", type text},
            {"channel_code", type text},
            {"country", type text},
            {"region", type text},
            {"state", type text},
            {"square_meters", Int64.Type},
            {"open_date", type date}
        })
in
    ChangedType
