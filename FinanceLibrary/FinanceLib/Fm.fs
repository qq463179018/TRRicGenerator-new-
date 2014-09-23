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
type Fm(format: Format) =
    inherit AFile()

    do
        base.ChooseFormat(format)

    new() =
        Fm(Vertical)