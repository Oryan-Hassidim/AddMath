module FuncPower =
    let rec pow f n =
        match n with
        | n when n <= 0 -> System.IndexOutOfRangeException() |> raise
        | 1 -> f
        | _ -> f >> pow f (n - 1)

    let inline ( ** ) f n = pow f n


let doubleIt x = 2 * x

open FuncPower
// For more information see https://aka.ms/fsharp-console-apps
printfn "Hello from F#"
printfn $"{(doubleIt ** 10) 1}"
