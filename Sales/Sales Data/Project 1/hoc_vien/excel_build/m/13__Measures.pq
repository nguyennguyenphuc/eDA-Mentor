// _Measures: Bảng một dòng để chứa mọi measure (ẩn cột ghi_chu trong Power Pivot).
// Nạp: Only Create Connection + Add this data to the Data Model
let
    Source = #table(type table [ghi_chu = text],
        {{"bảng chứa measure"}})
in
    Source
