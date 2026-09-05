// M Code
let
    Source = Excel.Workbook(File.Contents("C:\path\to\data.xlsx"), null, true),
    DimProduct = Source{[Item="Dim_Product",Kind="Sheet"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(DimProduct,{
        {"ProductID", type text},
        {"ProductName", type text},
        {"Category", type text},
        {"UnitPrice", type number},
        {"ReorderLevel", Int64.Type},
        {"LeadTime", Int64.Type}
    }),
    #"Removed Duplicates" = Table.Distinct(#"Changed Type", {"ProductID"})
in
    #"Removed Duplicates"