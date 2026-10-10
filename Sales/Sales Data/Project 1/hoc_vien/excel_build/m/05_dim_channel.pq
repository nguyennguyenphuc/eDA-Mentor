// dim_channel: Bảng dim: nạp và đặt kiểu.
// Nạp: Only Create Connection + Add this data to the Data Model
let
    Source = Excel.Workbook(
        File.Contents(SourceFolder & "Nexa_Sales_P1_Dataset.xlsx"), null, true),
    Navigation = Source{[Item = "tbl_dim_channel", Kind = "Table"]}[Data],
    ChangedType = Table.TransformColumnTypes(Navigation, {
            {"channel_code", type text},
            {"channel_name", type text},
            {"fin_channel", type text},
            {"revenue_recognition_rule", type text}
        })
in
    ChangedType
