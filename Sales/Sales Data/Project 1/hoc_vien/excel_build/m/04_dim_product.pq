// dim_product: Bảng dim: nạp và đặt kiểu.
// Nạp: Only Create Connection + Add this data to the Data Model
let
    Source = Excel.Workbook(
        File.Contents(SourceFolder & "Nexa_Sales_P1_Dataset.xlsx"), null, true),
    Navigation = Source{[Item = "tbl_dim_product", Kind = "Table"]}[Data],
    ChangedType = Table.TransformColumnTypes(Navigation, {
            {"product_key", Int64.Type},
            {"product_name", type text},
            {"brand", type text},
            {"color", type text},
            {"subcategory_key", type text},
            {"subcategory", type text},
            {"category_key", type text},
            {"category", type text},
            {"list_price_usd", Currency.Type},
            {"unit_cost_usd", Currency.Type}
        })
in
    ChangedType
