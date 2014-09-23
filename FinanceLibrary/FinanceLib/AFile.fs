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
type AFile() = 
    [<DefaultValue>]
    val mutable format: AFormat

    abstract member ChooseFormat : Format -> unit
    abstract member Save :          string      -> unit
    abstract member Load :          string      -> unit
    abstract member Send :          List<string> -> unit
    abstract member AddProp :       Dictionary<string, string> -> unit
    abstract member Generate :      unit        -> unit
    abstract member GenerateAndSave : string    -> string -> unit

    default this.ChooseFormat (formattype: Format) =
        this.format <-
            match formattype with
            | Vertical    -> new VerticalFormat() :> AFormat
            | Horizontal  -> new HorizontalFormat() :> AFormat
            | _             -> failwith "Wrong argument"

    default this.Save (path: string) =
        this.format.Save path

    default this.Load (path: string) =
        this.format.Load path

    default this.Send (recipients: List<string>) =
        this.format.Send recipients

    default this.AddProp (prop: Dictionary<string, string>) = 
        this.format.AddProp prop

    default this.Generate () = 
        this.format.Generate()

    default this.GenerateAndSave (template: string) (path: string) =
     this.format.GenerateAndSave template path

