﻿namespace Finance.Lib

open System
open System.IO
open System.Text
open System.Collections.Generic
open NdaTemplate
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

    let mutable template = List.empty
    let mutable content = List.empty

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
    // Load Template from CSV file
    //
    //-----------------------------------------------------------------------
    member private this.LoadCsv path = 
        try
            this.Template <-
                path
                |> File.ReadLines
                |> Seq.map (fun line -> HLine(Array.toList(line.Split ',')))
                |> Seq.toList
            ()
        with
            | _ -> failwith "Error in loading template from .csv";

    static member LoadCsv path = 
        try
            path
            |> File.ReadLines
            |> Seq.map (fun line -> HLine(Array.toList(line.Split ',')))
            |> Seq.toList
        with
            | _ -> failwith "Error in loading template from .csv";

    //-----------------------------------------------------------------------
    //Load Template
    //
    // Options:
    //      -.txt
    //      -.csv
    //      -.xls
    //
    //-----------------------------------------------------------------------
    override this.Load (path: string) =
        match path with
        | path when path.ToLower().EndsWith(".csv") || path.ToLower().EndsWith(".txt")
                -> this.LoadCsv path
        | _     -> failwith "unsupported file type"

    //-----------------------------------------------------------------------
    // Sends email to recipients with the generated FM
    //
    //-----------------------------------------------------------------------
    override this.Send recipients =
        ()

    //-----------------------------------------------------------------------
    // Saving to file
    //
    //-----------------------------------------------------------------------
    member private this.SaveCsv (path: string) =
        let wr = new StreamWriter(path)
        this.Content 
        |> List.rev
        |> List.iter (fun (line: string) -> wr.WriteLine(line.Replace(" \",\"", "\",\"").Replace("\",\"", ",")))
        wr.Close()

    member private this.SaveXls (path: string) =
        let GetColumnName (columnNb: int) =
            let rec GetColumnNamePart dividend columnName  = 
                match dividend with
                    | dividend when dividend > 0 -> 
                                                    let modulo = ((dividend - 1) % 26)
                                                    GetColumnNamePart ((int)((dividend - modulo) / 26)) (Convert.ToChar(65 + modulo).ToString() + columnName)
                    | _   -> columnName
            in GetColumnNamePart columnNb ""



        let app = new ApplicationClass(Visible = false) 

        let workbook = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
        let worksheet = (workbook.Worksheets.[1] :?> Worksheet)

        this.Content 
        |> List.rev
        |> List.iteri (fun i (line: string) ->  let lineTab1 = line.Replace("\",\"",",")
                                                let lineTab = line.Split ','
                                                worksheet.Range("A" + (i + 1).ToString(), GetColumnName(lineTab.Length) + (i + 1).ToString()).Value2 <- lineTab)
        
        worksheet.SaveAs(path)
        workbook.Close(false, false)
        app.Quit()

    member private this.SaveTxt (path: string) =
        let wr = new StreamWriter(path)
        this.Content 
        |> List.rev
        |> List.iter (fun (line: string) -> wr.WriteLine(line.Replace(" \",\"", "\",\"").Replace("\",\"", "\t")))
        wr.Close()
    
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
        match path with
        | path when path.ToLower().EndsWith(".csv") -> this.SaveCsv path
        | path when path.ToLower().EndsWith(".xls") -> this.SaveXls path
        | path when path.ToLower().EndsWith(".txt") -> this.SaveTxt path
        | _     -> failwith "unsupported file type"

    //-----------------------------------------------------------------------
    //
    //                             METHODS
    //
    //-----------------------------------------------------------------------

    //-----------------------------------------------------------------------
    // Generates string from FmLine
    //
    //-----------------------------------------------------------------------
    member private this.GenerateLine line =
        line
        |> List.fold (fun str x -> str + ((this.ParseMyValue(x.ToString())) + "\",\"")) ""


    member private this.AddToTemplateLine (line: string list) =
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
        ()

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
            | Titles(lineList)-> ()
            | HLine(lineList)  -> this.AddToTemplateLine lineList
 
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
            | Titles(lineList)-> this.AddToTemplateLine lineList
            | HLine(lineList)  -> ()
 
        this.Template 
        |> List.map GetInsideList

    //-----------------------------------------------------------------------
    // Choose Template
    //
    //-----------------------------------------------------------------------
    member public this.ChooseTemplate templateName =
        this.Template <- this.SelectTemplate templateName

    member private this.SelectTemplate (templateName: string) =
        match templateName with
        | "ThTemplate"  -> ThFm
        | "ThDomChain"  -> DomChain
        | "ThQa"        -> QaAdd
        | "ThForIdn"    -> ForIdn
        | "ThIaAdd"     -> IaAdd
        | "ThWrtAddCw"  -> WrtAddCw
        | "ThWrtAdd"    -> WrtAdd
        | "ThCwNvdr"    -> CwNvdr
        | "ThCw"        -> Cw
        | "CnQaAddCNord"-> QaAddCNord
        | "CnQaAddCNord2"-> QaAddCNord2
        | "CnQaAddCNord3"-> QaAddCNord3
        | "CnQaAddCNord4"-> QaAddCNord4
        | "CnQaChg"     -> QaChg
        | "CnBgChg"     -> BgChg
        | "CnIaAddFutDat"-> IaAddFutDat
        | "CnLotAdd"    -> LotAdd
        | "CnQaAddFutDat"-> QaAddFutDat
        | "CnQaChgFtd"  -> QaChgFtd
        | "CnTickAdd"   -> TickAdd
        | "CnIdnAddSS"  -> IdnAddSS
        | "CnIdnAddSZ"  -> IdnAddSZ
        | "TwIdnCb"     -> TwIdnCb
        | "TwCbBulk"    -> TwCbBulk
        | "TwFutureStep11"   -> TwFutureStep11
        | "TwFutureStep12"   -> TwFutureStep12
        | "TwFutureStep13"   -> TwFutureStep13
        | "TwFutureStep14"   -> TwFutureStep14
        | "TwFutureStep15"   -> TwFutureStep15
        | "TwFutureStep16"   -> TwFutureStep16
        | "TwFutureStep17"   -> TwFutureStep17
        | "TwFutureStep18"   -> TwFutureStep18
        | "TwFutureTeSsf"    -> TwFutureTeSsf
        | "TwFutureTeSpd"    -> TwFutureTeSpd
        | "TwFutureTeAlias"  -> TwFutureTeAlias
        | "TwFutureTeChain"  -> TwFutureTeChain
        | "TwFutureTeSup"    -> TwFutureTeSup
        | _                  -> failwith "Wrong template"