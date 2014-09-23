namespace Ric.FormatLib

//
// Generic types
//

type Nb = int
type Key = string
type Value = string

//
// Vertical File
//

type VLine =
    | Empty
    | Separator     of Nb
    | Title         of Value
    | NoValueLine   of Key
    | Line          of Key * Value

type VPart =
    | Head    of VLine list
    | Body    of VLine list
    | Footer  of VLine list

//
// Horizontal File
//

type HPart =
    | Titles of Key list
    | HLine   of Value list

type HFile = HFile of HPart list

//
// Other File
//

//type File =
//   | Geda
//   | Nda
//   | Fm

type Format = 
    | Vertical
    | Horizontal
    | Raw