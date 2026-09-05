
// M Code in Power Query
let
    Source = Excel.Workbook(File.Contents("C:\path\to\data.xlsx"), null, true),
    FactInventory = Source{[Item="Fact_Inventory",Kind="Sheet"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(FactInventory,{
        {"Date", type date},
        {"QtyReceived", Int64.Type},
        {"QtyIssued", Int64.Type},
        {"OpeningStock", Int64.Type},
        {"ClosingStock", Int64.Type}
    }),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "NetMovement", 
        each [QtyReceived] - [QtyIssued], Int64.Type),
    #"Added Custom1" = Table.AddColumn(#"Added Custom", "TransactionType", 
        each if [QtyReceived] > 0 and [QtyIssued] = 0 then "GR"
             else if [QtyIssued] > 0 and [QtyReceived] = 0 then "GI"
             else "Adjustment", type text),
    #"Filtered Rows" = Table.SelectRows(#"Added Custom1", 
        each [QtyReceived] >= 0 and [QtyIssued] >= 0)
in
    #"Filtered Rows"