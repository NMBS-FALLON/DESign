namespace DESign_Bot_FS_Tools

#if INTERACTIVE
//#r "../packages/Deedle.1.2.5/lib/net40/Deedle.dll"
#r "Microsoft.Office.Interop.Excel.dll"
//System.Environment.CurrentDirectory <- @"C:\Users\darien.shannon\Documents\Code\F#\FSharp\NMBS_TOOLS\NMBS_TOOLS\bin\Debug"
#endif

open System
open System.IO
open Microsoft.Office.Interop.Excel
open System.Runtime.InteropServices
open DESign_Bot_FS_Tools.ArrayExtensions
open System.Text.RegularExpressions

[<AutoOpen>]
module HelperFunctions =

    let (|Regex|_|) pattern input =
        let m = Regex.Match(input, pattern)
        if m.Success then Some(List.tail [ for g in m.Groups -> g.Value ])
        else None

    let nullableToOption<'T> value =
        match (box value) with
        | null  -> None
        | value when value = (box "") -> None
        | _ -> Some ((box value) :?> 'T)

    let toDouble (s : obj) =
            match (box s) with
            | v when v = (box "") -> 0.0
            | _ -> Convert.ToDouble(s)

module BOM =
   
    [<AbstractClass>]
    type OWSJ(mark : string, length : float, tcxl : float, tcxr : float, notesString : string Option) =
        member this.Mark = mark
        member this.Length = length
        member this.TCXL = tcxl
        member this.TCXR = tcxr
        member this.NotesString = notesString

        member this.LoadNoteIDs =       
            let getLoadNotes (note : string) =
                if note.Contains("(") then
                    let loadNoteStart = note.IndexOf("(")
                    let loadNotes = note.Substring(loadNoteStart)
                    let loadNotes = loadNotes.Split([|"("; ","; ")"|], StringSplitOptions.RemoveEmptyEntries)
                    let loadNotes = loadNotes |> List.ofArray
                    loadNotes |> List.map (fun (s : string) -> s.Trim())
                else
                    []
        
            match this.NotesString with
                | Some notes -> getLoadNotes notes
                | None -> []

        member this.SpecialNoteIDs =
            let getSpecialNotes (note : string) =
                if note.Contains("[") then
                    let specialNotesStart = note.IndexOf("[")
                    let specialNotesEnd = note.IndexOf("]")
                    let specialNotes = note.Substring(specialNotesStart, specialNotesEnd + 1)
                    let specialNotes = specialNotes.Split([|"["; ","; "]"|], StringSplitOptions.RemoveEmptyEntries)
                    let specialNotes = specialNotes |> List.ofArray
                    specialNotes |> List.map (fun (s: string) -> s.Trim())
                else
                    []
        
            match this.NotesString with
                | Some notes -> getSpecialNotes notes
                | None -> []

    type Load =
        {
        Type : string
        Category : string
        Position : obj
        Load1Value : obj
        Load1DistanceFt : obj
        Load1DistanceIn : obj
        Load2Value : obj
        Load2DistanceFt : obj
        Load2DistanceIn : obj
        Ref : obj
        LoadCases : int list
        }

        member this.LoadCasesString =
            match this.LoadCases with
            | [] -> ""
            | _ -> 
                this.LoadCases
                |> List.map string
                |> List.reduce (fun s1 s2 -> s1 + "," + s2)
            
        static member create(loadType, category, position, load1Value, load1DistanceFt, load1DistanceIn, load2Value, load2DistanceFt, load2DistanceIn, ref, loadcases) =
            {Type = loadType; Category = category; Position = position; Load1Value = load1Value;
                Load1DistanceFt = load1DistanceFt; Load1DistanceIn = load1DistanceIn; Load2Value = load2Value;
                Load2DistanceFt = load2DistanceFt; Load2DistanceIn = load2DistanceIn; Ref = ref; LoadCases = loadcases}


    type Joist(mark, description : string, length, tcxl, tcxr, notesString) =
        inherit OWSJ(mark, length, tcxl, tcxr, notesString)
        member this.Description = description

        member this.UDL =
            let TL, LL =
                match this.Description with
                | Regex @"(\d+\.?\d*)/(\d+\.?\d*)" [tl; ll] -> float tl,float ll
                | _ -> 10000.0, 5000.0
            let DL = TL - LL
            Load.create("U", "CL", "TC", DL,
                            null, null, null, null, null, null, [3])

    type Panel =
        {
        Number : int
        LengthFt : float
        LengthIn : float
        }

    type GirderGeometry =
        {
        Mark : String
        A_Ft : float
        A_In : float
        B_Ft : float
        B_In : float
        Panels : Panel list
        }

        member this.NumPanels =
            1 + List.length (this.Panels)

        member this.getPanelDim (panelNumber : int) =
            let mutable ft = this.A_Ft
            let mutable inch = this.A_In
            let mutable i = 0
            while i < panelNumber - 1 do
                ft <- ft + this.Panels.[i].LengthFt
                inch <- inch + this.Panels.[i].LengthIn
                i <- i + 1
            ft <- ft + (inch / 12.0) - ((inch / 12.0) % 1.0)
            inch <- ((inch / 12.0) % 1.0) * 12.0
            (ft, inch)

    type _AdditionalJoist =
        {
        LocationFt : obj
        LocationIn : obj
        Load : float
        }

        member this.ToLoad() =
            {
            Type = "C"
            Category = "CL"
            Position = "TC"
            Load1Value = this.Load * 1000.0
            Load1DistanceFt = this.LocationFt
            Load1DistanceIn = this.LocationIn
            Load2Value = null
            Load2DistanceFt = null
            Load2DistanceIn = null
            Ref = null
            LoadCases = []
            }

    type AdditionalJoist =
        {
        Mark : string
        AdditionalJoists : Load list
        }



    type Girder(mark, description: string, length, tcxl, tcxr, notesString, geometry : GirderGeometry, additionalJoistLoads : Load list) =
        inherit OWSJ(mark, length, tcxl, tcxr, notesString)
        member this.Description = description
        member this.Geometry = geometry
        member this.AdditionalJoistLoads = additionalJoistLoads
        
             

    type Note =
        {
        NoteID : string
        Text : string
        }

    type LoadNote =
        {
        LoadID : string
        Load : Load
        }

    type BOM =
        {
        GeneralNotes : Note list
        SpecialNotes : Note list
        LoadNotes : LoadNote list
        Joists : Joist list
        Girders : Girder list
        SDS : float
        }


module ExtractBOM = 
    open BOM

    let getBOM (bomWorkBook: Workbook) =

        let workSheetNames = [for sheet in bomWorkBook.Worksheets -> (sheet :?> Worksheet).Name]

        let loads =
            let loadSheetNames = workSheetNames |> List.filter (fun name -> name.Contains("L ("))
            if (List.isEmpty loadSheetNames) then
                []
            else
                let arrayList =
                    seq [for loadSheet in loadSheetNames do
                            let sheet = (bomWorkBook.Worksheets.[loadSheet] :?> Worksheet)
                            let loads = sheet.Range("A14","M55").Value2 :?> obj [,]
                            yield loads]
                let loadsAsArray = Array2D.joinMany (Array2D.joinByRows) arrayList

                let getLoads (a2D : obj[,]) =

                    let getLoadCases (loadCaseString: string) =
                        let loadNotes = loadCaseString.Split([|","|], StringSplitOptions.RemoveEmptyEntries)
                        let loadNotes = loadNotes
                                        |> List.ofArray
                                        |> List.map (fun string -> System.Int32.Parse(string))
                        loadNotes

                    let getLoadFromArraySlice (a : obj []) =
                        {
                        Type = string a.[1]
                        Category = string a.[2]
                        Position = a.[3]
                        Load1Value = a.[5]
                        Load1DistanceFt =  a.[6]
                        Load1DistanceIn = a.[7]
                        Load2Value = a.[8]
                        Load2DistanceFt = a.[9]
                        Load2DistanceIn = a.[10]
                        Ref = a.[11]
                        LoadCases = getLoadCases (string a.[12])
                        }


                    let mutable startRowIndex = Array2D.base1 a2D 
                    let endIndex = (a2D |> Array2D.length1) - (if startRowIndex = 0 then 1 else 0)
                    let startColIndex = Array2D.base2 a2D
                    let loadNotes = 
                        let mutable loadID = ""
                        [for currentIndex = startRowIndex to endIndex do
                            if a2D.[currentIndex, startColIndex + 1] <> null && a2D.[currentIndex, startColIndex + 1] <> (box "") then
                                if a2D.[currentIndex, startColIndex] <> null && a2D.[currentIndex, startColIndex] <> (box "") then
                                    loadID <- (string a2D.[currentIndex, startColIndex]).Trim()
                                yield {LoadID = loadID; Load = getLoadFromArraySlice a2D.[currentIndex, *]}]
                    loadNotes
                getLoads loadsAsArray

        let getNotes (a2D : obj[,]) =
        
            let mutable startRowIndex = Array2D.base1 a2D 
            let endIndex = (a2D |> Array2D.length1) - (if startRowIndex = 0 then 1 else 0)
            let startColIndex = Array2D.base2 a2D
            let notes : Note list =
                let mutable currentIndex = startRowIndex
                [while currentIndex <= endIndex do
                    let mutable noteNumber = ""
                    let mutable note = ""
                    let mutable additionalLines = 0
                    if a2D.[currentIndex, startColIndex] <> null && a2D.[currentIndex, startColIndex] <> (box "") then
                        noteNumber <- string a2D.[currentIndex, startColIndex]
                        note <- string a2D.[currentIndex, startColIndex + 1]
                        while (currentIndex + additionalLines + 1 < endIndex && (a2D.[currentIndex + additionalLines + 1, startColIndex] = null || a2D.[currentIndex + additionalLines + 1,startColIndex] = (box ""))) do
                            additionalLines <- additionalLines + 1
                            note <- String.concat " " [note; string a2D.[currentIndex + additionalLines, startColIndex + 1]]

                        yield
                            {
                            NoteID = noteNumber
                            Text = note
                            }
                    currentIndex <- currentIndex + 1 + additionalLines]
            notes

        let generalNotes =
            let generalNotesNames = workSheetNames |> List.filter (fun name -> name.Contains("P ("))
            if (List.isEmpty generalNotesNames) then
                []
            else
                let arrayList =
                    seq [for noteSheet in generalNotesNames do
                            let sheet = (bomWorkBook.Worksheets.[noteSheet] :?> Worksheet)
                            yield sheet.Range("A8", "H47").Value2 :?> obj [,]]
                let notesAsArray = Array2D.joinMany (Array2D.joinByRows) arrayList
                let notes = getNotes notesAsArray
                notes

        let isLiveLoadNote (note: string) =
                Regex.IsMatch(note, "[LS] *= *(\d+\.?\d*) *([Kk%]) *")

        let liveLoadUNO =
            let liveLoadNotes = generalNotes |> List.filter (fun note -> isLiveLoadNote note.Text)
            match liveLoadNotes with
            | [] -> "L = 0.0K"
            | _ -> liveLoadNotes.[0].Text

        let isSlopeLoadNote (note: string) =
            Regex.IsMatch(note, "SP *: *(\d+\.?\d*)/(\d+\.?\d*)")



        let sds =
            let isSDSNote (note : string) =
                Regex.IsMatch(note, "SDS *= *(\d+\.?\d*) *")
            let sdsNotes = generalNotes |> List.filter (fun note -> isSDSNote note.Text)
            let sds =
                match sdsNotes with
                | [] -> 100.0
                | _ -> 
                    let sdsNote = sdsNotes.[0]
                    match sdsNote.Text with
                    | Regex @"SDS *= *(\d+\.?\d*) *" [sds] -> float sds
                    | _ -> 100.0
            sds
                    
        

        let specialNotes =
            let specialNoteSheetNames = workSheetNames |> List.filter (fun name -> name.Contains("N ("))
            if (List.isEmpty specialNoteSheetNames) then
                []
            else
                let arrayList =
                    seq [for specialNoteSheet in specialNoteSheetNames do
                            let sheet = (bomWorkBook.Worksheets.[specialNoteSheet] :?> Worksheet)
                            yield sheet.Range("A13", "J51").Value2 :?> obj [,]]
                let notesAsArray = Array2D.joinMany (Array2D.joinByRows) arrayList
                let notes = getNotes notesAsArray
                notes

        let slopeSpecialNotes =
            (specialNotes |> List.filter (fun note -> isSlopeLoadNote note.Text))

        let liveLoadSpecialNotes =
            (specialNotes |> List.filter (fun note -> isLiveLoadNote note.Text))

        

        let joists =
            let joistSheetNames = workSheetNames |> List.filter (fun name -> name.Contains("J ("))
            if (List.isEmpty joistSheetNames) then
                []
            else
                let arrayList =
                    seq [for joistSheet in joistSheetNames do
                            let sheet = (bomWorkBook.Worksheets.[joistSheet] :?> Worksheet)
                            if (sheet.Range("A21").Value2 :?> string) = "MARK" then
                                yield sheet.Range("A23","GF40").Value2 :?> obj [,]
                            else
                                yield sheet.Range("A16", "GF45").Value2 :?> obj [,]]

                let joistsAsArray = Array2D.joinMany (Array2D.joinByRows) arrayList

                let getJoists (a2D : obj [,]) =
        
                    let mutable startRowIndex = Array2D.base1 a2D 
                    let endIndex = (a2D |> Array2D.length1) - (if startRowIndex = 0 then 1 else 0)
                    let startColIndex = Array2D.base2 a2D
                    let joists : Joist list =
                        [for currentIndex = startRowIndex to endIndex do
                            if a2D.[currentIndex, startColIndex] <> null && a2D.[currentIndex, startColIndex] <> (box "") then
                                let mark = string a2D.[currentIndex, startColIndex]
                                let description = string a2D.[currentIndex, startColIndex + 2]
                                let notesString = nullableToOption<string> a2D.[currentIndex, startColIndex + 26]
                                let length = 
                                    let feet = toDouble (a2D.[currentIndex, startColIndex + 3])
                                    let inches = toDouble (a2D.[currentIndex, startColIndex + 4])
                                    feet + inches/12.0

                                // **** NEEED TO IMPLEMENT TCXL AND TCXR **** //
                                let tcxl = 0.0
                                let tcxr = 0.0
                                yield Joist(mark, description, length, tcxl, tcxr, notesString)]
                    joists

                getJoists joistsAsArray

        let girders =
            let girderSheetNames = workSheetNames |> List.filter (fun name -> name.Contains("G ("))
            if (List.isEmpty girderSheetNames) then
                []
            else
                let arrayList1 =
                    seq [for girderSheet in girderSheetNames do
                            let sheet = (bomWorkBook.Worksheets.[girderSheet] :?> Worksheet)
                            if (sheet.Range("A26").Value2 :?> string) = "MARK" then
                                yield sheet.Range("A28","AA45").Value2 :?> obj [,]
                            else
                                yield sheet.Range("A14", "AA45").Value2 :?> obj [,]]
                let arrayList2 =
                    seq [for sheet in bomWorkBook.Worksheets do
                            let sheet = (sheet :?> Worksheet)
                            if sheet.Name.Contains("G (") then
                                yield sheet.Range("AB14","BG45").Value2 :?> obj [,]]
                let girdersAsArray = Array2D.joinMany (Array2D.joinByRows) arrayList1
                let additionalJoistsAsArray = Array2D.joinMany (Array2D.joinByRows) arrayList2

                let getGirders (sheet1 : obj [,], sheet2 : obj [,]) =
        
                    let getGirderGeometry (a2D : obj[,]) =
                        let startIndex = Array2D.base1 a2D 
                        let endIndex = (a2D |> Array2D.length1) - (if startIndex = 0 then 1 else 0)
                        let startColIndex = Array2D.base2 a2D
                        let mutable row = startIndex
        
        
                        [while row <= endIndex do
                            let numPanels =
                                let mutable panelCounter = 0
                                let markCell = a2D.[row, startColIndex]
                                let panelNoCell = a2D.[row, startColIndex + 4]

                                if (markCell <> null && markCell <> box "") then
                                    panelCounter <- 1
                                    let mutable continueCounting = true
                                    while continueCounting = true do
                                        row <- row + 1
                                        let markCell = a2D.[row, startColIndex]
                                        let panelNoCell = a2D.[row, startColIndex + 4]
                                        if (markCell = null || markCell = box "") &&
                                           (panelNoCell <> null && panelNoCell <> box "") then
                                            panelCounter <- panelCounter + 1
                                        else
                                            continueCounting <- false
                                            row <- row - 1
                                panelCounter
           

                            let panels =
                                [for j = 1 to numPanels do
                                    let mark = string (a2D.[row- numPanels + 1, startColIndex])
                                    let n = Convert.ToInt32 (a2D.[row- numPanels + j, startColIndex + 4])
                                    let feet = Convert.ToDouble (a2D.[row- numPanels + j, startColIndex + 5])
                                    let inch = Convert.ToDouble (a2D.[row- numPanels + j, startColIndex + 6])
                                    for k = 1 to Convert.ToInt32 (a2D.[row- numPanels + j, startColIndex + 4]) do
                                        yield
                                            {
                                            Number = 1
                                            LengthFt = Convert.ToDouble (a2D.[row- numPanels + j, startColIndex + 5])
                                            LengthIn = Convert.ToDouble (a2D.[row- numPanels + j, startColIndex + 6])
                                            }]
                            if numPanels <> 0 then
                                let mark = string (a2D.[row- numPanels + 1, startColIndex])
                                let aFt = Convert.ToDouble (a2D.[row- numPanels + 1, startColIndex + 2])
                                let aIn = Convert.ToDouble (a2D.[row- numPanels + 1, startColIndex + 3])
                                let bFt = Convert.ToDouble (a2D.[row, startColIndex + 7])
                                let bIn = Convert.ToDouble (a2D.[row, startColIndex + 8])
                                yield
                                     {
                                     Mark = string (a2D.[row- numPanels + 1, startColIndex])
                                     A_Ft = Convert.ToDouble (a2D.[row- numPanels + 1, startColIndex + 2])
                                     A_In = Convert.ToDouble (a2D.[row- numPanels + 1, startColIndex + 3])
                                     B_Ft = Convert.ToDouble (a2D.[row, startColIndex + 7])
                                     B_In = Convert.ToDouble (a2D.[row, startColIndex + 8])
                                     Panels = panels
                                     }
                            row <- row+ 1] 

                    let getAdditionalJoistsFromArraySlice (a : obj [])  =
                        let mutable col = 16
                        [while col <= 28 do
                            if (a.[col + 2] <> null && a.[col + 2] <> (box "")) then
                                if (a.[col] <> null && a.[col] <> (box "")) || (a.[col + 1] <> null && a.[col + 1] <> (box "")) then
                                    let additionalJoist =
                                        {
                                        LocationFt = string a.[col]
                                        LocationIn = string a.[col + 1]
                                        Load = Convert.ToDouble(a.[col + 2])
                                        }
                                    yield additionalJoist.ToLoad()
                            col <- col + 4]

                    let getAdditionalJoistsFromArray (a2D : obj [,]) =
                        let mutable startRowIndex = Array2D.base1 a2D
                        let colIndex = Array2D.base2 a2D
                        let endIndex = (a2D |> Array2D.length1) - (if startRowIndex = 0 then 1 else 0)
                        let additionalJoists : AdditionalJoist list =
                            [for currentIndex = startRowIndex to endIndex do
                                if a2D.[currentIndex, colIndex] <> null && a2D.[currentIndex, colIndex] <> (box "") then
                                    yield
                                        {
                                        Mark = string a2D.[currentIndex, colIndex]
                                        AdditionalJoists = getAdditionalJoistsFromArraySlice a2D.[currentIndex, *]
                                        } ]
                        additionalJoists

                    let getPanelDim (panel : int) (girderGeom : GirderGeometry) =
                        let mutable ft = girderGeom.A_Ft
                        let mutable inch = girderGeom.A_In
                        let mutable i = 0
                        while i < panel - 1 do
                            ft <- ft + girderGeom.Panels.[i].LengthFt
                            inch <- inch + girderGeom.Panels.[i].LengthIn
                            i <- i + 1
                        ft <- ft + (inch / 12.0) - ((inch / 12.0) % 1.0)
                        inch <- ((inch / 12.0) % 1.0) * 12.0
                        (ft, inch)  

                    let mutable startIndex = Array2D.base1 sheet1
                    let colIndex = Array2D.base2 sheet1
                    let endIndex = (sheet1 |> Array2D.length1) - (if startIndex = 0 then 1 else 0)
                    let allGirderGeometry = getGirderGeometry sheet2
                    let additionalJoists = getAdditionalJoistsFromArray sheet2
                    let girders : Girder list =
                        [for currentIndex = startIndex to endIndex do
                            if sheet1.[currentIndex, colIndex] <> null && sheet1.[currentIndex, colIndex] <> (box "") then
                                let mark = string sheet1.[currentIndex, colIndex]
                                let geometry = allGirderGeometry |> List.find (fun geom -> geom.Mark = mark)
                                let mark = mark
                                let description = string sheet1.[currentIndex, colIndex + 2]
                                let length = toDouble(sheet1.[currentIndex, colIndex + 3]) +
                                                toDouble(sheet1.[currentIndex, colIndex + 4]) / 12.0
                                let tcxl = toDouble(sheet1.[currentIndex, colIndex + 6]) +
                                            toDouble(sheet1.[currentIndex, colIndex + 7]) / 12.0
                                let tcxr = toDouble(sheet1.[currentIndex, colIndex + 9]) +
                                            toDouble(string sheet1.[currentIndex, colIndex + 10]) / 12.0
                                let notesString =  nullableToOption<string> sheet1.[currentIndex, colIndex + 25]
                        
                                let additionalJoistsOnGirder = additionalJoists |> List.filter (fun addJoist -> addJoist.Mark = mark)
                                let additionalLoads =
                                    [for addJoist in additionalJoistsOnGirder do
                                        for load in addJoist.AdditionalJoists do
                                            let mutable locationFt = load.Load1DistanceFt
                                            let mutable locationIn = load.Load1DistanceIn
            
                                            if (string locationFt) = "P" then
                                                let panel = System.Int32.Parse((string locationIn).Replace("#", ""))
                                                let ft, inch = geometry.getPanelDim panel
                                                locationFt <- box ft
                                                locationIn <- box inch                   
                                            yield {load with Load1DistanceFt = locationFt; Load1DistanceIn = locationIn}] 
                                let additionalJoists = additionalLoads
                                yield Girder(mark, description, length, tcxl, tcxr, notesString, geometry, additionalJoists)]
                    girders

                getGirders (girdersAsArray, additionalJoistsAsArray)               

        {
        GeneralNotes = generalNotes
        SpecialNotes = specialNotes
        LoadNotes = loads
        Joists = joists
        Girders = girders
        SDS = sds
        }





module Modifiers =
    open BOM

    let seperateSeismic (bom : BOM) =
        
        let joistSDS (joist : Joist) =
            let uniformSDS = 0.14 * bom.SDS * System.Convert.ToDouble(joist.UDL.Load1Value)
            Load.create ("U", "SM", "TC", uniformSDS, null, null, null, null, null, null, [3])

        let joistLC3Loads (joist : Joist) =
                bom.LoadNotes
                |> List.filter (fun load -> joist.LoadNoteIDs |> List.contains load.LoadID )
                |> List.map (fun note -> note.Load)
                |> List.filter (fun load -> load.Category <> "WL" && load.Category <> "SM" && (load.LoadCases = [] || (load.LoadCases |> List.contains 1)))
                |> List.map (fun load -> {load with LoadCases = [3]})
                |> List.append [joist.UDL; joistSDS joist]

        ()

    let adjustSinglePitchedOWSJs (bom : BOM) =

        let endDepths (owsj : OWSJ) =
        
            let slopeSpecialNote =
                    [for slopeSpecialNote in bom.SpecialNotes do
                            if List.contains slopeSpecialNote.NoteID owsj.SpecialNoteIDs then
                                yield slopeSpecialNote.Text]
            let endDepths =
                match slopeSpecialNote with
                | [] -> 0.0, 0.0
                | _ ->
                    match slopeSpecialNote.[0] with
                    | Regex @"SP *: *(\d+\.?\d*)/(\d+\.?\d*)" [leDepth; reDepth] -> (float leDepth, float reDepth)
                    | _ -> 0.0, 0.0 

            endDepths

        ()