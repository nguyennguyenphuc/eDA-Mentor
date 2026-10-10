// fact_sales_line: Dòng bán: đặt kiểu, bỏ dòng trùng, thêm cột tiền dòng (gross, net).
// Nạp: Only Create Connection + Add this data to the Data Model
let
    Source = Excel.Workbook(
        File.Contents(SourceFolder & "Nexa_Sales_P1_Dataset.xlsx"), null, true),
    Navigation = Source{[Item = "tbl_fact_sales_line", Kind = "Table"]}[Data],
    ChangedType = Table.TransformColumnTypes(Navigation, {
            {"order_number", Int64.Type},
            {"line_item", Int64.Type},
            {"order_date", type date},
            {"delivery_date", type date},
            {"customer_key", Int64.Type},
            {"store_key", Int64.Type},
            {"channel_code", type text},
            {"product_key", Int64.Type},
            {"quantity", Int64.Type},
            {"unit_price", Currency.Type},
            {"promo_code", type text},
            {"discount_amount", Currency.Type},
            {"order_status", type text},
            {"currency_code", type text},
            {"unit_cost", Currency.Type}
        }),
    // 120 dòng trùng hoàn toàn (mọi cột giống nhau) -> giữ 1
    RemovedDuplicates = Table.Distinct(ChangedType),
    AddedGross = Table.AddColumn(RemovedDuplicates, "gross_amount",
        each [quantity] * [unit_price], Currency.Type),
    AddedNet = Table.AddColumn(AddedGross, "net_amount",
        each [gross_amount] - [discount_amount], Currency.Type)
in
    AddedNet
