namespace Finance.Lib

open Finance.Ast
open System.IO
open Parser
open System.Text
open Microsoft.FSharp.Text.Lexing
open Microsoft.FSharp.Text.Parsing
open Finance.Parser 
open Finance.Lexer
open System.Collections.Generic

[<AbstractClass>]
type AFormat() = 

    let mutable prop = List.empty
    let mutable current = 0

    // Properties
    member this.Prop
        with get() = prop 
        and set properties = prop <- properties

    // Current position in properties
    member this.Current 
        with get() = current
        and set cur = current <- cur

    abstract member Save :          string      -> unit
    abstract member Load :          string      -> unit
    abstract member Generate :      unit        -> unit
    abstract member ParseMyValue :  string      -> string
    abstract member GenerateAndSave : string    -> string -> unit
    abstract member Send :          List<string> -> unit
    abstract member AddProp :       Dictionary<string, string> -> unit

    //-----------------------------------------------------------------------
    // Function calling the Lexer
    //
    // Return the parsed string
    //
    //-----------------------------------------------------------------------
    default this.ParseMyValue value =
        try
            let lexbuf = LexBuffer<_>.FromString value
            let GetExprList prog =
                match prog with
                | ParsedLine(exprList) -> exprList

            let rec LexToString lex =
                let rec GetResult lex1 (sb: StringBuilder) =
                    match lex1 with
                    | head :: tail  ->
                        let expression = GetStringFromExpr (head: Expr) (this.Prop.[this.Current])
                        sb.Append(expression: string) |> ignore
                        GetResult tail sb
                    | []            -> sb.ToString()
                GetResult lex <| new StringBuilder("")
            LexToString (GetExprList <| start Finance.Lexer.token lexbuf)

        with
            | _ -> failwith value

    //-----------------------------------------------------------------------
    // Add Property
    //
    //-----------------------------------------------------------------------
    default this.AddProp (props: Dictionary<string, string>) =
        this.Prop <- this.Prop @ [props]