namespace Finance.Lib

open System.Collections.Generic

type AFormat = 
    class
        val prop : Dictionary<string, string> list
        val current : int

        abstract member Save :          string      -> unit
        abstract member Load :          string      -> unit
        abstract member Send :          List<string> -> unit
        abstract member ParseMyValue :  string      -> string
        abstract member AddProp :       Dictionary<string, string> -> unit
        abstract member Generate :      unit        -> unit
        abstract member GenerateAndSave : string    -> string -> unit
    end

