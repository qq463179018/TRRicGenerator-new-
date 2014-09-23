namespace Finance.Lib

open System.IO
open System.Text
open FmTemplate
open Parser
open Finance.Ast
open Finance.Parser 
open Finance.Lexer  
open Microsoft.FSharp.Text.Lexing
open Microsoft.FSharp.Text.Parsing

[<Class>]
type VerticalFormat() =
    inherit AFormat() 

    let mutable template = List.empty
    let mutable content = List.empty

    // Template
    member this.Template 
        with get() = template
        and set temp = template <- temp

    // Content of the Fm
    member this.Content 
        with get() = content
        and set cont = content <- cont


    //
    //Abstract class Implementation
    //
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
    
    //
    // Load Template from CSV file
    //
    member private this.LoadCsv path = 
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

    //
    // Load Template from TXT file
    //
    member private this.LoadTxt path = 
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

    //
    //Load Template
    //
    // Options:
    //      -.txt
    //      -.csv
    //      -.xls
    //
    override this.Load (path: string) =
        match path with
        | path when path.ToLower().EndsWith(".csv") -> this.LoadCsv path
        | path when path.ToLower().EndsWith(".txt") -> this.LoadTxt path
        | _                                         -> failwith "unsupported file type"

    //
    // Sends email to recipients with the generated FM
    //
    override this.Send recipients =
        ()

    //
    // Saving to file
    //
    member private this.SaveTxt (path: string) =
        let wr = new StreamWriter(path)
        this.Content 
        |> List.rev
        |> List.iter (fun (line: string) -> wr.WriteLine(line))
        wr.Close()

    member private this.SaveCsv (path: string) =
        let FormatCsvLine (lineTab: string[]) =
            match lineTab.Length with
            | 1 -> lineTab.[0].Trim() + ";"
            | _ -> lineTab.[0].Trim() + ";" + lineTab.[1].Trim()

        let wr = new StreamWriter(path)
        this.Content 
        |> List.rev
        |> List.iter (fun (line: string) -> 
                                line.Split ':'
                                |> FormatCsvLine
                                |> (fun (line: string) -> wr.WriteLine(line))
                                )
        wr.Close()
    
    //
    // Saving result file
    //
    // options:
    //      -.txt
    //      -.csv
    //      -.xls
    //

    override this.Save (path: string) =
        match path with
        | path when path.ToLower().EndsWith(".csv") -> this.SaveCsv path
        | path when path.ToLower().EndsWith(".txt") -> this.SaveTxt path
        | _                                         -> failwith "unsupported file type"
     
    member private this.FillKeyInLine (key: string) =
        key + (String.init (32 - key.Length) (fun i -> " "))

    //
    // Generates string from FmLine
    //
    member private this.GenerateLine = function
        | Empty             -> ""
        | Separator(nb)     -> String.init nb (fun i -> "-")
        | Title(title)      -> title
        | NoValueLine(key)  -> this.FillKeyInLine key + ":"
        | Line(key, value)  -> (this.FillKeyInLine key) + ":    " + (this.ParseMyValue value)


    //
    // Generates Fm from template and properties
    //
    (*
    member public this.Generate () = 
        let AddToTemplate (line: FmLine) =
            this.Content <- ((this.GenerateLine line) :: (this.Content))

        let GetInsideList = function
            | FmHead(lineList) 
            | FmBody(lineList) 
            | FmFooter(lineList) -> lineList 
                                    |> List.iter AddToTemplate
 
        this.Template 
        |> List.map GetInsideList
        |> List.rev
    *)

    member private this.AddToTemplateLine (line: VLine) =
        this.Content <- ((this.GenerateLine line) :: (this.Content))

    //-----------------------------------------------------------------------
    // Generates file from template and properties
    //
    //-----------------------------------------------------------------------
    override this.Generate () = 
        this.Current <- 0
        this.Content <- List.Empty

        this.GenerateTitles()
        |> ignore

        for prop in this.Prop do
            this.GenerateLines()

    member public this.Generate template = 
        this.ChooseTemplate template
        this.Generate()

    override this.GenerateAndSave (template: string) (path) = 
        this.Generate template
        this.Save path

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
 
        this.Template 
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
 
        this.Template 
        |> List.map GetInsideList

    //
    // Choose Template
    //
    member public this.ChooseTemplate templateName =
        this.Template <- this.SelectTemplate templateName

    member private this.SelectTemplate (templateName: string) =
        match templateName with
        | "TwTemplate" -> TwTemplate
        | _ -> failwith "Wrong template"

        