// Learn more about F# at http://fsharp.org

open System

type var = int
type pow = var * int
type term = int * List<pow>
type poly = List<term>
type expr = Div of poly * poly

let rec factorial x = 
  match x with 
  | 0 -> 1
  | n -> n * factorial (n-1)

let 

[<EntryPoint>]
let main argv =
    printfn "Hello World from F#!"
    (factorial 6)

