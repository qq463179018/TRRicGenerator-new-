namespace Ric.FormatLib

open System.IO
open Finance.Ast
open Finance.Parser 
open Finance.Lexer  
open Microsoft.FSharp.Text.Lexing
open Microsoft.FSharp.Text.Parsing
open System.Collections.Generic

module Parser =

    //------------------------------------------------------------------------
    // Template Functions
    //
    //  Equals:
    //  Startswith:
    //  EndsWith:
    //
    //-----------------------------------------------------------------------
    let GetCondResult funcName (param1: string) (param2: string) =
        match funcName with
        | "Equals"      -> if param1 = param2 then "true" else "false"
        | "Length"      -> if param1.Length.ToString() = param2 then "true" else "false"
        | "Startswith"  -> if param1.StartsWith(param2) then "true" else "false"
        | "Endswith"    -> if param1.EndsWith(param2) then "true" else "false"
        | _             -> "true"
    
    //
    // And checkin
    //
    let CheckAndCond = function
        | ("true", "true")  -> "true"
        | _                 -> "false"
    
    //
    // Or checkin
    //
    let CheckOrCond = function
        | ("true", _)
        | (_, "true")  -> "true"
        | _            -> "false"

    //
    // Parsing of the result from the lexer
    // Recursive function
    //
    let GetStringFromExpr value (properties: Dictionary<string, string>) =
        try
            let rec GetStringFromExprRec value (properties: Dictionary<string, string>) =
                match value with
                | Val(line)                 -> line
                | Parent(line)              -> GetStringFromExprRec line properties
                | Int(nb)                   -> nb.ToString()
                | Float(nb)                 -> nb.ToString()
                | Concat(expr1, expr2)      -> (GetStringFromExprRec expr1 properties) + (GetStringFromExprRec expr2 properties)
                | SubVal(line, nb1, nb2)    -> (GetStringFromExprRec line properties).Substring(nb1, nb2)
                | Var(line)                 -> properties.Item(GetStringFromExprRec line properties)
                | Or(cond1, cond2)          -> CheckOrCond ((GetStringFromExprRec cond1 properties), (GetStringFromExprRec cond2 properties))
                | And(cond1, cond2)         -> CheckAndCond ((GetStringFromExprRec cond1 properties), (GetStringFromExprRec cond2 properties))
                | IfThen(cond, ifcond)      ->
                    match (GetStringFromExprRec cond properties) with
                    | "true"    -> GetStringFromExprRec ifcond properties
                    | _         -> ""
                | IfThenElse(cond, ifcond, elsecond) -> 
                    match (GetStringFromExprRec cond properties) with
                    | "true"    -> GetStringFromExprRec ifcond properties
                    | "false"   -> GetStringFromExprRec elsecond properties
                    | _         -> ""
                | Function(funcName, par1, par2) -> GetCondResult funcName (GetStringFromExprRec par1 properties) (GetStringFromExprRec par2 properties)
            GetStringFromExprRec value properties
        with
            | _ -> failwith "Error in parsing in Ast Reading"