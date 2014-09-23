namespace Ric.FormatLib

open System
open System.IO
open System.Text
open System.Collections.Generic
open Parser
open Finance.Ast
open Finance.Parser 
open Finance.Lexer  
open Microsoft.FSharp.Text.Lexing
open Microsoft.FSharp.Text.Parsing
open Microsoft.Office.Interop.Excel

[<Class>]
type VerticalFormat() =
    inherit AFormat() 

    let mutable template = VFile(List.empty)
    let mutable content = Seq.empty

    // Content of the file
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
        this.Content

    (*
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
    *)

    //-----------------------------------------------------------------------
    // Load Template
    //
    // Options:
    //      -.txt
    //      -.csv
    //      -.xls
    //      -.xlsx
    //
    //-----------------------------------------------------------------------
    override this.LoadTemplate (pathOrTemplate) =
        try
            match pathOrTemplate with
            | :? VFile      as template -> this.ChooseTemplate <| template
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

    member public this.TestLineCsv (line: string[]) =
        let ContainsOnlySeparator linePart = 
            linePart
            :> seq<char>
            |> Seq.forall (fun c -> c = '-')

        match line.[0] with
        | ""    -> Empty
        | _     -> 
            match line.[1] with
            | "" when  ContainsOnlySeparator line.[0] = true  -> Separator(line.[0].Length)
            | "" when  ContainsOnlySeparator line.[0] = false -> Title(line.[0])
            | _     -> Line(line.[0], line.[1])

    member public this.TestLineTxt (line: string[]) =
        let ContainsOnlySeparator linePart = 
            linePart
            :> seq<char>
            |> Seq.forall (fun c -> c = '-')

        match line.Length with
        | 1 when line.[0] = ""                          -> Empty
        | 1 when ContainsOnlySeparator line.[0] = true  -> Separator(line.[0].Length)
        | 1                                             -> Title(line.[0].Trim())
        | _ when line.[1] = ""                          -> NoValueLine(line.[0].Trim())
        | _                                             -> Line(line.[0].Trim(), line.[1].Trim())
    

    member private this.LoadTemplateCsv path = 
        ()

    (*
        try
            let FmList =
                path
                |> File.ReadLines
                |> Seq.map (fun line -> line.Split ';')
                |> Seq.map this.TestLineCsv
                |> Seq.toList
            this.Template <- [Body(FmList)]
            ()
        with
            | _ -> failwith "Error in loading template from .csv";

    *)

    member private this.LoadTemplateTxt path = 
        ()

    (*
        try
            let FmList =
                path
                |> File.ReadLines
                |> Seq.map (fun line -> line.Split ':')
                |> Seq.map this.TestLineTxt
                |> Seq.toList
            this.Template <- [Body(FmList)]
            ()
        with
            | _ -> failwith "Error in loading template from .csv";
    *)

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
    //      -.xlsx
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
        let WriteLine (line: string list) = 
            match line with
                | line when line.Length.Equals(1) -> line.Head
                | line when line.Length > 1     -> (line.Item 0) + ":    " + (line.Item 1)
                | _                             -> ""          
        this.Content 
        |> Seq.iter (fun (line: string seq) -> wr.WriteLine(line 
                                                            |> Seq.toList 
                                                            |> WriteLine))
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

    member private this.FillKeyInLine (key: string) =
        key + (String.init (32 - key.Length) (fun i -> " "))

    
    //
    // Generates string from FmLine
    //
    member private this.GenerateLine = function
        | Empty             -> seq [""]
        | Separator(nb)     -> seq [String.init nb (fun i -> "-")]
        | Title(title)      -> seq [title]
        | NoValueLine(key)  -> seq [this.FillKeyInLine key; ""]
        | Line(key, value)  -> seq [(this.FillKeyInLine key); (this.ParseMyValue value)]


    member private this.AddToTemplateLine (line: VLine) =
        this.Content <- Seq.append this.Content [this.GenerateLine line]

    //-----------------------------------------------------------------------
    // GetContent
    //
    //-----------------------------------------------------------------------
    override this.GetContent () =
        this.Content

    override this.SetContent newContent =
        this.Content <- newContent
    
    













    //-----------------------------------------------------------------------
    // Generates file from template and properties
    //
    //-----------------------------------------------------------------------
    override this.Generate () = 
        this.Current <- 0
        this.Content <- List.Empty

        this.GenerateTitles()

        for prop in this.Prop do
            this.GenerateLines()

        this.GenerateFooters()

    //-----------------------------------------------------------------------
    // Add line to Nda from template and properties
    //
    //-----------------------------------------------------------------------
    member public this.GenerateLines () = 
        let GetInsideList = function
            | Head(lineList) 
            | Footer(lineList)-> ()
            | Body(lineList)  -> lineList 
                                 |> List.iter this.AddToTemplateLine
 
        match this.Template with
        | VFile(part)  -> part
                          |> List.map GetInsideList
                          |> ignore

        this.Current <- this.Current + 1

    //-----------------------------------------------------------------------
    // Add line to Nda from template and properties
    //
    //-----------------------------------------------------------------------
    member public this.GenerateTitles () = 
        let GetInsideList = function
            | Head(lineList)  -> lineList 
                                 |> List.iter this.AddToTemplateLine
            | Footer(lineList)
            | Body(lineList)  -> ()
 
        match this.Template with
        | VFile(part)   -> part 
                           |> List.map GetInsideList
                           |> ignore

    member public this.GenerateFooters () = 
        ()

    //-----------------------------------------------------------------------
    // Choose Template
    //
    //-----------------------------------------------------------------------
    member public this.ChooseTemplate templateName =
        this.Template <- templateName

        