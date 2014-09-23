namespace Ric.FormatLib

open System
open System.IO
open System.Text
open System.Collections.Generic
open Template
open Parser
open Finance.Ast
open Finance.Parser 
open Finance.Lexer
open Microsoft.FSharp.Text.Lexing
open Microsoft.FSharp.Text.Parsing
open Microsoft.Office.Interop.Excel

[<Class>]
type HorizontalFormat() =
    inherit AFormat() 

    let mutable template = HFile(List.Empty)
    let mutable content = Seq.empty

    // Content of the Nda
    member this.Content 
        with get() = content
        and set cont = content <- cont

    // Template
    member this.Template
        with get() = template
        and set temp = template <- temp

    //-----------------------------------------------------------------------
    //
    //                   ABSTRACT CLASS IMPLEMENTATION
    //
    //-----------------------------------------------------------------------
    
    //-----------------------------------------------------------------------
    // Load Template
    //
    // Options:
    //      -.txt
    //      -.csv
    //      -.xls
    //
    //-----------------------------------------------------------------------
    override this.LoadTemplate (pathOrTemplate) =
        try
            match pathOrTemplate with
            | :? HFile      as template -> this.ChooseTemplate <| template
            | :? string     as path     -> this.LoadTemplateFromFile path
            | _                         -> failwith "wrong template value"
        with
            | _     -> failwith "Loading failed"

    member private this.LoadTemplateFromFile (path: string) =
        let LoadingTemplateFunction = 
            match path with
            | path when path.ToLower().EndsWith(".csv") -> this.LoadTemplateCsv
            | _                                         -> this.LoadTemplateTxt
        path
        |> LoadingTemplateFunction

    member private this.LoadTemplateCsv path = 
        try
            ','
            |> this.LoadTemplateFile path
        with
            | _ -> failwith "Error in loading template from CSV file";

    member private this.LoadTemplateTxt path = 
        try
            '\t'
            |> this.LoadTemplateFile path
        with
            | _ -> failwith "Error in loading template from txt file";

    member private this.LoadTemplateFile path delimiter =
        let fileContent = path |> File.ReadLines 
        this.Template <- HFile( 
            (fileContent
            |> Seq.take 1
            |> Seq.map (fun line -> Titles(Array.toList(line.Split delimiter)))
            |> Seq.toList) @
            (fileContent
            |> Seq.skip 1
            |> Seq.map (fun line -> HLine(Array.toList(line.Split delimiter)))
            |> Seq.toList))

    // member private this.LoadTemplateXls path = 
    // TODO
    

    //-----------------------------------------------------------------------
    // Load file
    //
    // Options:
    //      -.txt
    //      -.csv
    //      -.xls
    //      -.xslx
    //
    //-----------------------------------------------------------------------
    override this.Load (path: string) =
        try
            let LoadingFunction =
                match path with
                | path when path.ToLower().EndsWith(".csv") -> this.LoadCsv
                | path when path.ToLower().EndsWith(".xls") -> this.LoadXls
                | path when path.ToLower().EndsWith(".xlsx") -> this.LoadXls
                | _                                         -> this.LoadTxt
            path
            |> LoadingFunction
        with
            | _     -> failwith "Loading failed"

    member private this.LoadCsv path = 
        this.LoadFile path ','

    member private this.LoadTxt path = 
        this.LoadFile path '\t'

    member private this.LoadFile path delimiter =
        this.Content <- path
                        |> File.ReadLines
                        |> Seq.map (fun line -> Array.toSeq(line.Split delimiter))
        this.Content

    member private this.LoadXls path = 
        let toGrid groupSize items =
            let noOfItems = items |> Seq.length
            let noOfGroups = int(Math.Ceiling(float(noOfItems) / float(groupSize)))
            seq { for groupNo in 0..groupSize-1 do
                    yield seq {
                        for tableNoInGroup in 0..groupSize-1 do
                            let absoluteIndex = int((groupNo * groupSize) + tableNoInGroup)
                            if (absoluteIndex < noOfItems) then
                                let toYield = items |> Seq.nth absoluteIndex
                                yield toYield.ToString()    
                    }
            }
        let app = ApplicationClass(Visible = false)
        let book = app.Workbooks.Open path
        let sheet = book.Worksheets.[1] :?> _Worksheet
        let vals = sheet.UsedRange.Value2 :?> obj[,]
        let length = sheet.UsedRange.Columns.Count

        this.Content <- vals
                        |> Seq.cast<string>
                        |> toGrid length

        this.Content
        

        

    //-----------------------------------------------------------------------
    // Sends email to recipients with generated file
    //
    //-----------------------------------------------------------------------
    override this.Send recipients =
        ()

    //-----------------------------------------------------------------------
    // Saving result file
    //
    // options:
    //      -.txt
    //      -.csv
    //      -.xls
    //
    //-----------------------------------------------------------------------
    override this.Save (path: string) =
        try
            let SavingFunction = 
                match path with
                | path when path.ToLower().EndsWith(".csv") -> this.SaveCsv
                | path when path.ToLower().EndsWith(".xls") 
                        || path.ToLower().EndsWith(".xlsx") -> this.SaveXls
                | _                                         -> this.SaveTxt
            path
            |> SavingFunction
        with
            | _     -> failwith "Saving failed"

    member private this.SaveCsv path =
       this.SaveFile path ","

    member private this.SaveTxt path = 
        this.SaveFile path "\t"

    member private this.SaveFile (path: string) delimiter =
        let wr = new StreamWriter(path)
        this.Content 
        |> Seq.iter (fun (line: string seq) -> wr.WriteLine(line
                                                |> Seq.toList
                                                |> List.fold (fun r s -> r + s + delimiter) ""))
        wr.Close()


    member private this.SaveXls (path: string) =
        let GetColumnName (columnNb: int) =
            let rec GetColumnNamePart dividend columnName  = 
                match dividend with
                    | dividend when dividend > 0 -> let modulo = ((dividend - 1) % 26)
                                                    GetColumnNamePart ((int)((dividend - modulo) / 26)) (Convert.ToChar(65 + modulo).ToString() + columnName)
                    | _   -> columnName
            in GetColumnNamePart columnNb ""

        let app = new ApplicationClass(Visible = false) 

        let workbook = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
        let worksheet = (workbook.Worksheets.[1] :?> Worksheet)

        this.Content 
        |> Seq.iteri (fun i (line: string seq) -> let lineTab = line |> Seq.toArray
                                                  worksheet.Range("A" + (i + 1).ToString(), GetColumnName(lineTab.Length) + (i + 1).ToString()).Value2 <- lineTab)
        worksheet.SaveAs(path)
        workbook.Close(false, false)
        app.Quit()

    //-----------------------------------------------------------------------
    // GetContent
    //
    //-----------------------------------------------------------------------
    override this.GetContent () =
        this.Content


    //-----------------------------------------------------------------------
    //
    //                             METHODS
    //
    //-----------------------------------------------------------------------

    //-----------------------------------------------------------------------
    // Generates Content from template line
    //
    //-----------------------------------------------------------------------
    member private this.AddToTemplateLine (line: string list) =
        this.Content <- Seq.append this.Content  [(line
                                                |> List.map (fun linepart -> this.ParseMyValue(linepart))
                                                |> List.toSeq)]

    //-----------------------------------------------------------------------
    // Generates file from template and properties
    //
    //-----------------------------------------------------------------------
    override this.Generate () = 
        this.Current <- 0
        this.Content <- Seq.empty

        this.GenerateTitles()
        |> ignore

        for prop in this.Prop do
            this.GenerateLines()

    member public this.GenerateLines () = 
        let GetInsideList = function
            | Titles(lineList)-> ()
            | HLine(lineList)  -> this.AddToTemplateLine lineList
 
        match this.Template with
        | HFile(part)   -> part |> List.map GetInsideList |> ignore

        this.Current <- this.Current + 1

    member public this.GenerateTitles () = 
        let GetInsideList = function
            | Titles(lineList)-> this.AddToTemplateLine lineList
            | HLine(lineList)  -> ()
        match this.Template with
        | HFile(part)   -> part |> List.map GetInsideList

    //-----------------------------------------------------------------------
    // Choose Template
    //
    //-----------------------------------------------------------------------
    member public this.ChooseTemplate templateName =
        this.Template <- templateName