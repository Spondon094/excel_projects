// Power Query (M) code behind the "Dashboard_project" query.
// Set the Source folder path below to the folder that holds the LAPD CSV files (only CSV files in it).

section Section1;

shared Dashboard_project = let
    Source = Folder.Files("<PATH_TO_YOUR_CSV_FOLDER>"),
    #"Filtered Hidden Files1" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File from Dashboard_project", each #"Transform File from Dashboard_project"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File from Dashboard_project"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File from Dashboard_project", Table.ColumnNames(#"Transform File from Dashboard_project"(#"Sample File"))),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Table Column1",{{"Source.Name", type text}, {"DR_NO", Int64.Type}, {"Date Rptd", type datetime}, {"DATE OCC", type datetime}, {"TIME OCC", Int64.Type}, {"AREA", Int64.Type}, {"AREA NAME", type text}, {"Rpt Dist No", Int64.Type}, {"Part 1-2", Int64.Type}, {"Crm Cd", Int64.Type}, {"Crm Cd Desc", type text}, {"Mocodes", type text}, {"Vict Age", Int64.Type}, {"Vict Sex", type text}, {"Vict Descent", type text}, {"Premis Cd", Int64.Type}, {"Premis Desc", type text}, {"Weapon Used Cd", Int64.Type}, {"Weapon Desc", type text}, {"Status", type text}, {"Status Desc", type text}, {"Crm Cd 1", Int64.Type}, {"Crm Cd 2", Int64.Type}, {"Crm Cd 3", Int64.Type}, {"Crm Cd 4", type any}, {"LOCATION", type text}, {"Cross Street", type text}, {"LAT", type number}, {"LON", type number}}),
    #"Replaced Value" = Table.ReplaceValue(#"Changed Type",#datetime(2026, 11, 21, 0, 0, 0),#datetime(2026, 11, 21, 0, 0, 0),Replacer.ReplaceValue,{"Date Rptd"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value",{{"Date Rptd", type date}}),
    #"Replaced Value1" = Table.ReplaceValue(#"Changed Type1",#datetime(2026, 11, 21, 0, 0, 0),#datetime(2026, 11, 21, 0, 0, 0),Replacer.ReplaceValue,{"DATE OCC"}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Replaced Value1",{{"DATE OCC", type date}}),
    #"Replaced Value2" = Table.ReplaceValue(#"Changed Type2","     ","",Replacer.ReplaceText,{"LOCATION"}),
    #"Replaced Value3" = Table.ReplaceValue(#"Replaced Value2","  ","",Replacer.ReplaceText,{"LOCATION"}),
    #"Replaced Value4" = Table.ReplaceValue(#"Replaced Value3","  ","",Replacer.ReplaceText,{"LOCATION"}),
    #"Replaced Value5" = Table.ReplaceValue(#"Replaced Value4"," ","",Replacer.ReplaceText,{"LOCATION"})
in
    #"Replaced Value5";

shared #"Sample File" = let
    Source = Folder.Files("<PATH_TO_YOUR_CSV_FOLDER>"),
    Navigation1 = Source{0}[Content]
in
    Navigation1;

shared #"Sample File Parameter1" = #"Sample File" meta [IsParameterQuery=true, BinaryIdentifier=#"Sample File", Type="Binary", IsParameterQueryRequired=true];

shared #"Transform Sample File from Dashboard_project" = let
    Source = Csv.Document(#"Sample File Parameter1",[Delimiter=",", Columns=28, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])
in
    #"Promoted Headers";

[ FunctionQueryBinding = "{""exemplarFormulaName"":""Transform Sample File from Dashboard_project""}" ]
#"Transform File from Dashboard_project" = let
    Source = (#"Sample File Parameter1") => let
        Source = Csv.Document(#"Sample File Parameter1",[Delimiter=",", Columns=28, Encoding=1252, QuoteStyle=QuoteStyle.None]),
        #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true])
    in
        #"Promoted Headers"
in
    Source;