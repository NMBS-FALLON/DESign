// Learn more about F# at http://fsharp.org
// See the 'F# Tutorial' project for more help.

[<AutoOpen>]
module Model =
    type Model =
        { Count : int
          StepSize : int }
    let init () =
        { Count = 0
          StepSize = 1 }

[<AutoOpen>]
module Messages =
    type Msg =
        | Increment
        | Decrement
        | SetStepSize of int

[<AutoOpen>]
module Update =
    let update msg m =
        match msg with
        | Increment -> { m with Count = m.Count + m.StepSize }
        | Decrement -> { m with Count = m.Count - m.StepSize }
        | SetStepSize i -> { m with StepSize = i }

[<AutoOpen>]
module Bindings =
    open Elmish.WPF
    open Model


    let bindings model dispatch =
        [
            "CounterValue" |> Binding.oneWay (fun m -> m.Count)
            "Increment" |> Binding.cmd (fun m -> Increment)
            "Decrement" |> Binding.cmd (fun m -> Decrement)
            "StepSize" |> Binding.twoWay
                (fun m -> float m.StepSize)
                (fun newVal m -> int newVal |> SetStepSize)
        
        ]

open AutoDESign.Views
open System
open Elmish
open Elmish.WPF

[<EntryPoint; STAThread>]
let main argv = 
    Program.mkSimple init update bindings
    |> Elmish.WPF.Program.runWindow (MainWindow())
