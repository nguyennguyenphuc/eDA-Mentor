// dim_promotion: Bảng dim: nạp và đặt kiểu.
// Nạp: Only Create Connection + Add this data to the Data Model
let
    Source = Excel.Workbook(
        File.Contents(SourceFolder & "Nexa_Sales_P1_Dataset.xlsx"), null, true),
    Navigation = Source{[Item = "tbl_dim_promotion", Kind = "Table"]}[Data],
    ChangedType = Table.TransformColumnTypes(Navigation, {
            {"promo_code", type text},
            {"promo_name", type text},
            {"discount_pct", Percentage.Type},
            {"window_from_mmdd", type text},
            {"window_to_mmdd", type text},
            {"channel", type text}
        })
in
    ChangedType
