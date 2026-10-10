// dim_order_status: Bảng dim: nạp và đặt kiểu.
// Nạp: Only Create Connection + Add this data to the Data Model
let
    Source = Excel.Workbook(
        File.Contents(SourceFolder & "Nexa_Sales_P1_Dataset.xlsx"), null, true),
    Navigation = Source{[Item = "tbl_dim_order_status", Kind = "Table"]}[Data],
    ChangedType = Table.TransformColumnTypes(Navigation, {
            {"order_status", type text},
            {"status_name", type text},
            {"count_in_revenue", type text},
            {"rule", type text}
        })
in
    ChangedType
