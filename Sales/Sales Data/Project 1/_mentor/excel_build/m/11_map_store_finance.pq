// map_store_finance: Bảng quy đổi mã Finance → store_key (chỉ tạo connection).
// Nạp: Only Create Connection
let
    Source = Excel.Workbook(
        File.Contents(SourceFolder & "Nexa_Sales_P1_Dataset.xlsx"), null, true),
    Navigation = Source{[Item = "tbl_map_store_finance", Kind = "Table"]}[Data],
    ChangedType = Table.TransformColumnTypes(Navigation, {
            {"fin_store_code", type text},
            {"store_key", Int64.Type},
            {"fin_channel", type text},
            {"valid_from", type date},
            {"valid_to", type date},
            {"note", type text}
        })
in
    ChangedType
