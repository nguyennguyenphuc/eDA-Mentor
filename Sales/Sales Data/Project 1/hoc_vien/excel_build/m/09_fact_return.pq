// fact_return: Hoàn trả: đặt kiểu, ghép store_key/channel/product từ dòng bán gốc.
// Nạp: Only Create Connection + Add this data to the Data Model
let
    Source = Excel.Workbook(
        File.Contents(SourceFolder & "Nexa_Sales_P1_Dataset.xlsx"), null, true),
    Navigation = Source{[Item = "tbl_fact_return", Kind = "Table"]}[Data],
    ChangedType = Table.TransformColumnTypes(Navigation, {
            {"return_id", type text},
            {"order_number", Int64.Type},
            {"line_item", Int64.Type},
            {"return_date", type date},
            {"return_qty", Int64.Type},
            {"return_amount", Currency.Type},
            {"return_reason", type text}
        }),
    // lấy cửa hàng, kênh, sản phẩm từ dòng bán gốc
    MergedSales = Table.NestedJoin(ChangedType, {"order_number", "line_item"},
        fact_sales_line, {"order_number", "line_item"}, "sales",
        JoinKind.LeftOuter),
    ExpandedSales = Table.ExpandTableColumn(MergedSales, "sales",
        {"store_key", "channel_code", "product_key"}),
    ChangedType2 = Table.TransformColumnTypes(ExpandedSales,
        {{"store_key", Int64.Type}, {"channel_code", type text},
         {"product_key", Int64.Type}})
in
    ChangedType2
