namespace Finance.Ast

type Expr = 
    | Val     of string 
    | SubVal  of Expr * int * int
    | Or      of Expr * Expr
    | And     of Expr * Expr
    | Int     of int
    | Float   of float
    | Var     of Expr
    | VarDate of Expr * Expr
    | Parent  of Expr
    | IfThen  of Expr * Expr
    | IfThenElse  of Expr * Expr * Expr
    | Function    of string * Expr * Expr
    | Concat      of Expr * Expr

type ParsedLine = ParsedLine of Expr list