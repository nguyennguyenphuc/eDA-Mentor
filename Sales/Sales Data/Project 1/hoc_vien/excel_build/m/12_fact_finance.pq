// fact_finance: Finance đã gắn store_key theo mã và ngày hiệu lực.
// Nạp: Only Create Connection + Add this data to the Data Model
let
    Source = finance_raw,
    MergedMap = Table.NestedJoin(Source, {"fin_store_code"},
        map_store_finance, {"fin_store_code"}, "map", JoinKind.LeftOuter),
    ExpandedMap = Table.ExpandTableColumn(MergedMap, "map",
        {"store_key", "valid_from", "valid_to"}),
    // giữ mã có hiệu lực tại ngày hạch toán (S38: GBR-03 -> GBR-03N từ 01/07/2019)
    FilteredValid = Table.SelectRows(ExpandedMap,
        each [posting_date] >= [valid_from] and [posting_date] <= [valid_to]),
    RemovedColumns = Table.RemoveColumns(FilteredValid, {"valid_from", "valid_to"}),
    // HQ-ADJ: store_key = null -> không thuộc cửa hàng nào
    ChangedType = Table.TransformColumnTypes(RemovedColumns,
        {{"store_key", Int64.Type}})
in
    ChangedType
