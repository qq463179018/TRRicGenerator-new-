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
type Geda() =
    inherit AFile()