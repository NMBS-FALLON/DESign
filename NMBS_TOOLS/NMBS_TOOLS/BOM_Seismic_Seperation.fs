namespace NMBS_Tools.BOM_Seismic_Seperation

module Seperator =
    #if INTERACTIVE
    //#r "../packages/Deedle.1.2.5/lib/net40/Deedle.dll"
    #r "Microsoft.Office.Interop.Excel.dll"
    //System.Environment.CurrentDirectory <- @"C:\Users\darien.shannon\Documents\Code\F#\FSharp\NMBS_TOOLS\NMBS_TOOLS\bin\Debug"
    #endif

    open System
    open System.IO
    open Microsoft.Office.Interop.Excel
    open System.Runtime.InteropServices
    open NMBS_Tools.ArrayExtensions

    type Load =
        {
        Type : string;
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

        member this.LoadCaseString =
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

    type LoadNote =
        {
        LoadNumber : string
        Load : Load
        }

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

        
    (*
    let testArray = array2D [[box "G43"; box "B"; box 5; box 8.5; box 5; box 8; box 0; box 7; box 8.5 ];
                             [box "G44"; box "B"; box 5; box 8.5; box 5; box 8; box 0; box 7; box 8.5 ];
                             [box "G45"; box "B"; box 5; box 8.5; box 5; box 8; box 0; box 7; box 8.5 ];
                             [box "G46"; box "B"; box 7; box 3.125; box 1; box 8; box 0; box 7; box 8.5 ];
                             [box null; box null; box null; box null; box 1; box 5; box 6; box 7; box 8.5 ];
                             [box null; box null; box null; box null; box 1; box 5; box 6; box 7; box 8.5 ];
                             [box null; box null; box null; box null; box 1; box 5; box 6; box 7; box 8.5 ];
                             [box ""; box "B"; box 5; box 8.5; box 1; box 10; box 0; box 7; box 8.5 ];
                             [box "G47"; box "B"; box 5; box 8.5; box 5; box 8; box 0; box 7; box 8.5 ];
                             [box ""; box "B"; box 5; box 8.5; box 5; box 8; box 0; box 7; box 8.5 ];
                             [box ""; box "B"; box 5; box 8.5; box 5; box 8; box 0; box 7; box 8.5 ];
                             [box ""; box "B"; box 5; box 8.5; box 5; box 8; box 0; box 7; box 8.5 ];
                             [box "G48"; box "B"; box 5; box 8.5; box 5; box 8; box 0; box 7; box 8.5 ];
                             [box null; box null; box null; box null; box null; box null; box null; box null; box null ];
                             [box null; box null; box null; box null; box null; box null; box null; box null; box null ];
                             [box null; box null; box null; box null; box null; box null; box null; box null; box null ]]

   let test2 = getGirderGeometry testArray
   let test3 = test2 |> List.filter (fun geom -> geom.Mark = "G46")
   let test4 = getPanelDim 4 (test3.[0])

   *)

   



            
        

    let getLoadNotes (note : string) =
        let loadNoteStart = note.IndexOf("(")
        let loadNotes = note.Substring(loadNoteStart)
        let loadNotes = loadNotes.Split([|"("; ","; ")"|], StringSplitOptions.RemoveEmptyEntries)
        let loadNotes = loadNotes |> List.ofArray
        loadNotes |> List.map (fun (s : string) -> s.Trim())
    
    type Joist =
        {
        Mark : string
        JoistSize : string
        LoadNoteString : string option
        }

        member this.LoadNoteList =
            match (this.LoadNoteString) with
            | Some notes -> getLoadNotes notes
            | None -> []

        member this.UDL =
            let size = this.JoistSize
            if size.Contains("/") then
                let sizeAsArray = size.Split( [|"LH"; "K"; "/"|], StringSplitOptions.RemoveEmptyEntries)
                let TL = float sizeAsArray.[1]
                let LL = float sizeAsArray.[2]
                let DL = TL - LL
                Some(Load.create("U", "CL", "TC", DL,
                             null, null, null, null, null, null, [3]))
            else
                None       

        member this.Sds sds =
            match this.UDL with
            | Some udl -> 
                let sds = 0.14 * sds * System.Convert.ToDouble(udl.Load1Value)
                Some (Load.create ("U", "SM", "TC", sds,
                              null, null, null, null, null, null, [3]))
            | None -> None

        member this.LC3Loads (loadNotes :LoadNote list) sds =
            match this.UDL, (this.Sds sds) with
            | Some udl, Some sds ->
                loadNotes
                |> List.filter (fun note -> this.LoadNoteList |> List.contains note.LoadNumber)
                |> List.map (fun note -> note.Load)
                |> List.filter (fun load -> load.Category <> "WL" && load.Category <> "SM" && (load.LoadCases = [] || (load.LoadCases |> List.contains 1)))
                |> List.map (fun load -> {load with LoadCases = [3]})
                |> List.append [udl; sds]
            | _ -> []


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

        member this.ToLoad2 (gGeom : GirderGeometry) =
            let mutable locationFt = this.LocationFt
            let mutable locationIn = this.LocationIn
            
            if (string this.LocationFt) = "P" then
                let panel = System.Int32.Parse((string this.LocationIn).Replace("#", ""))
                let ft, inch = getPanelDim panel gGeom
                locationFt <- box ft
                locationIn <- box inch                   
            {
            Type = "C"
            Category = "CL"
            Position = "TC"
            Load1Value = this.Load * 1000.0
            Load1DistanceFt = locationFt
            Load1DistanceIn = locationIn
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

    type Girder =
        {
        Mark : string
        GirderSize : string
        OverallLengthFt : float
        OverallLengthIn : float
        TcxlLengthFt : float
        TcxlLengthIn : float
        TcxrLengthFt : float
        TcxrLengthIn : float
        LoadNoteString : string option
        AdditionalJoists : Load list
        GirderGeometry : GirderGeometry
        }

        member this.LoadNoteList =
            match (this.LoadNoteString) with
            | Some notes -> Some (getLoadNotes notes)
            | None -> None

        member this.BaseLength =
            (this.OverallLengthFt + this.OverallLengthIn/12.0) -
             (this.TcxlLengthFt + this.TcxlLengthIn/12.0) -
              (this.TcxrLengthFt + this.TcxlLengthIn/12.0)

        member this.UDL_PDL =
            let size = this.GirderSize
            let sizeAsArray = size.Split( [|"G"; "BG"; "VG"; "N"; "K"|], StringSplitOptions.RemoveEmptyEntries)
            let load = sizeAsArray.[2].Split([|"/"|], StringSplitOptions.RemoveEmptyEntries)
            let minSpace =
                let geometry = this.GirderGeometry
                let aSpace = geometry.A_Ft + geometry.A_In / 12.0
                let bSpace = geometry.B_Ft + geometry.B_In / 12.0
                let minPanelSpace = List.min (geometry.Panels |> List.map (fun geom -> geom.LengthFt + geom.LengthIn / 12.0))
                Math.Min(Math.Min(aSpace, bSpace), minPanelSpace)
            if (Array.length load) = 2 then
                let TL = float load.[0]
                let LL = float load.[1]
                let DL = 1000.0 * (TL - LL)
                let UDL = DL / minSpace
                UDL, DL
                 
            else
                failwith "GirderSize is not in correct format"

        member this.SDS sds =
            let udl, _ = this.UDL_PDL
            let SDS = udl * 0.14 * sds
            Load.create("U", "SM", "TC", SDS,
                          null, null, null, null, null, null, [3])

        member this.DeadLoads =
            let _,dl = this.UDL_PDL
            let geom = this.GirderGeometry
            [for i = 1 to geom.NumPanels do
                let distanceFt, distanceIn = getPanelDim i geom
                yield
                    Load.create("C", "CL", "TC", dl, distanceFt, distanceIn,
                                 null, null, null, null, [3]) ]
            

        member this.LC3Loads (loadNotes :LoadNote list) sds =
                let additionalJoistLoads =
                    this.AdditionalJoists
                    |> List.map (fun load -> {load with LoadCases = [3]})
                loadNotes
                |> List.filter (fun note ->
                    match this.LoadNoteList with
                    | Some loadNoteList -> loadNoteList |> List.contains note.LoadNumber
                    | None -> false)
                |> List.map (fun note -> note.Load)
                |> List.filter (fun load -> load.Category <> "WL" && load.Category <> "SM" && (load.LoadCases = [] || (load.LoadCases |> List.contains 1)))
                |> List.map (fun load -> {load with LoadCases = [3]})
                |> List.append [this.SDS sds]
                |> List.append this.DeadLoads
                |> List.append additionalJoistLoads


            



    module CleanBomInfo =

        let nullableToOption<'T> value =
            match (box value) with
            | null  -> None
            | value when value = (box "") -> None
            | _ -> Some ((box value) :?> 'T)


        module CleanLoads =

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

            let getLoadNotesFromArray (a2D : obj[,]) =
                let mutable startRowIndex = Array2D.base1 a2D 
                let endIndex = (a2D |> Array2D.length1) - (if startRowIndex = 0 then 1 else 0)
                let startColIndex = Array2D.base2 a2D
                let loadNotes = 
                    let mutable loadNumber = ""
                    [for currentIndex = startRowIndex to endIndex do
                        if a2D.[currentIndex, startColIndex + 1] <> null && a2D.[currentIndex, startColIndex + 1] <> (box "") then
                            if a2D.[currentIndex, startColIndex] <> null && a2D.[currentIndex, startColIndex] <> (box "") then
                                loadNumber <- (string a2D.[currentIndex, startColIndex]).Trim()
                            yield {LoadNumber = loadNumber; Load = getLoadFromArraySlice a2D.[currentIndex, *]}]
                loadNotes

        module CleanJoists =

            let getJoistsFromArray (a2D : obj [,]) =
                let mutable startRowIndex = Array2D.base1 a2D 
                let endIndex = (a2D |> Array2D.length1) - (if startRowIndex = 0 then 1 else 0)
                let startColIndex = Array2D.base2 a2D
                let joists : Joist list =
                    [for currentIndex = startRowIndex to endIndex do
                        if a2D.[currentIndex, startColIndex] <> null && a2D.[currentIndex, startColIndex] <> (box "") then
                            yield
                                {
                                Mark = string a2D.[currentIndex, startColIndex]
                                JoistSize = string a2D.[currentIndex, startColIndex + 2]
                                LoadNoteString = nullableToOption<string> a2D.[currentIndex, startColIndex + 26]
                                }]
                joists

        module CleanGirders =

            let toDouble (s : obj) =
                match (box s) with
                | v when v = (box "") -> 0.0
                | _ -> Convert.ToDouble(s)

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

            let getGirders (sheet1 : obj [,], sheet2 : obj [,]) =
                let mutable startIndex = Array2D.base1 sheet1
                let colIndex = Array2D.base2 sheet1
                let endIndex = (sheet1 |> Array2D.length1) - (if startIndex = 0 then 1 else 0)
                let allGirderGeometry = getGirderGeometry sheet2
                let girders : Girder list =
                    [for currentIndex = startIndex to endIndex do
                        if sheet1.[currentIndex, colIndex] <> null && sheet1.[currentIndex, colIndex] <> (box "") then
                            let mark = string sheet1.[currentIndex, colIndex]
                            let geometry = allGirderGeometry |> List.find (fun geom -> geom.Mark = mark)
                            yield
                                {
                                Mark = mark
                                GirderSize = string sheet1.[currentIndex, colIndex + 2]
                                OverallLengthFt = toDouble(sheet1.[currentIndex, colIndex + 3])
                                OverallLengthIn = toDouble(sheet1.[currentIndex, colIndex + 4])
                                TcxlLengthFt = toDouble(sheet1.[currentIndex, colIndex + 6])
                                TcxlLengthIn = toDouble(sheet1.[currentIndex, colIndex + 7])
                                TcxrLengthFt = toDouble(sheet1.[currentIndex, colIndex + 9])
                                TcxrLengthIn = toDouble(string sheet1.[currentIndex, colIndex + 10])
                                LoadNoteString =  nullableToOption<string> sheet1.[currentIndex, colIndex + 25]
                                AdditionalJoists = []
                                GirderGeometry = geometry
                                }]
                girders



            let addAdditionalJoistLoadsToGirders (girders: Girder list, additionalJoists : AdditionalJoist list) =
                [for girder in girders do
                    let additionalJoistsOnGirder = additionalJoists |> List.filter (fun addJoist -> addJoist.Mark = girder.Mark)
                    let additionalLoads =
                        [for addJoist in additionalJoistsOnGirder do
                            for load in addJoist.AdditionalJoists do
                                let mutable locationFt = load.Load1DistanceFt
                                let mutable locationIn = load.Load1DistanceIn
            
                                if (string locationFt) = "P" then
                                    let panel = System.Int32.Parse((string locationIn).Replace("#", ""))
                                    let ft, inch = getPanelDim panel girder.GirderGeometry
                                    locationFt <- box ft
                                    locationIn <- box inch                   
                                yield {load with Load1DistanceFt = locationFt; Load1DistanceIn = locationIn}] 
                    let additionalJoists = girder.AdditionalJoists |> List.append additionalLoads
                    yield {girder with AdditionalJoists = additionalJoists}]
                    
    let saveWorkbook (title : string) (workbook : Workbook) =
            let title = title.Replace(".xlsm", " (IMPORT).xlsm")
            let title = title.Replace(".xlsx", " (IMPORT).xlsx")
            workbook.SaveAs(title)
    
    let getAllInfo reportPath getInfoFunction modifyWorkbookFunction (sds : float) =
        let tempExcelApp = new Microsoft.Office.Interop.Excel.ApplicationClass(Visible = false)
        tempExcelApp.DisplayAlerts = false |> ignore
        tempExcelApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable |> ignore
        //let bom = tempExcelApp.Workbooks.Open(bomPath)
        try 
            tempExcelApp.DisplayAlerts <- false
            let tempReportPath = System.IO.Path.GetTempFileName()      
            File.Delete(tempReportPath)
            File.Copy(reportPath, tempReportPath)
            let workbook = tempExcelApp.Workbooks.Open(tempReportPath)
            let info = getInfoFunction workbook
            modifyWorkbookFunction workbook info sds
            
            workbook |> saveWorkbook reportPath

            workbook.Close(false)
            Marshal.ReleaseComObject(workbook) |> ignore
            System.GC.Collect() |> ignore
            printfn "Finished processing %s." reportPath 
            printfn "Finished processing all files."
            info
        finally
            tempExcelApp.Quit()
            Marshal.ReleaseComObject(tempExcelApp) |> ignore
            System.GC.Collect() |> ignore            

    let getInfo (bom: Workbook) =

        let workSheetNames = [for sheet in bom.Worksheets -> (sheet :?> Worksheet).Name] 

        let loads =
            let loadSheetNames = workSheetNames |> List.filter (fun name -> name.Contains("L ("))
            if (List.isEmpty loadSheetNames) then
                []
            else
                let arrayList =
                    seq [for sheet in bom.Worksheets do
                            let sheet = (sheet :?> Worksheet)
                            if sheet.Name.Contains("L (") then
                                let loads = sheet.Range("A14","M55").Value2 :?> obj [,]
                                yield loads]   
                let loadsAsArray = Array2D.joinMany (Array2D.joinByRows) arrayList
                let loads = CleanBomInfo.CleanLoads.getLoadNotesFromArray loadsAsArray
                loads

        let joists =
            let joistSheetNames = workSheetNames |> List.filter (fun name -> name.Contains("J ("))
            if (List.isEmpty joistSheetNames) then
                []
            else
                let arrayList =
                    seq [for sheet in bom.Worksheets do
                            let sheet = (sheet :?> Worksheet)
                            if sheet.Name.Contains("J (") then
                                if (sheet.Range("A21").Value2 :?> string) = "MARK" then
                                    yield sheet.Range("A23","AA40").Value2 :?> obj [,]
                                else
                                    yield sheet.Range("A16", "AA45").Value2 :?> obj [,]]

                let joistsAsArray = Array2D.joinMany (Array2D.joinByRows) arrayList
                let joists = CleanBomInfo.CleanJoists.getJoistsFromArray joistsAsArray
                joists

        let girdersAndAdditionalJoists =
            let girderSheetNames = workSheetNames |> List.filter (fun name -> name.Contains("G ("))
            if (List.isEmpty girderSheetNames) then
                []
            else
                let arrayList1 =
                    seq [for sheet in bom.Worksheets do
                            let sheet = (sheet :?> Worksheet)
                            if sheet.Name.Contains("G (") then
                                if (sheet.Range("A26").Value2 :?> string) = "MARK" then
                                    yield sheet.Range("A28","AA45").Value2 :?> obj [,]
                                else
                                    yield sheet.Range("A14", "AA45").Value2 :?> obj [,]]
                let arrayList2 =
                    seq [for sheet in bom.Worksheets do
                            let sheet = (sheet :?> Worksheet)
                            if sheet.Name.Contains("G (") then
                                yield sheet.Range("AB14","BG45").Value2 :?> obj [,]]
                let girdersAsArray = Array2D.joinMany (Array2D.joinByRows) arrayList1
                let additionalJoistsAsArray = Array2D.joinMany (Array2D.joinByRows) arrayList2
                let girders = CleanBomInfo.CleanGirders.getGirders (girdersAsArray, additionalJoistsAsArray)
                
                let additionalJoists =
                    CleanBomInfo.CleanGirders.getAdditionalJoistsFromArray additionalJoistsAsArray

                (CleanBomInfo.CleanGirders.addAdditionalJoistLoadsToGirders (girders, additionalJoists))

        (joists, girdersAndAdditionalJoists, loads)


    type BomInfo = 
        {
        Joists : Joist list
        Girders : Girder list
        Loads : LoadNote list
        }

    let modifyWorkbookFunction (bom : Workbook) (bomInfo : Joist list * Girder list * LoadNote list) sds: Unit =

        bom.Unprotect()
        //for sheet in bom.Worksheets do
        //    let sheet = (sheet :?> Worksheet)
        //    sheet.Unprotect("AAABBBBBABA-")
        
        let workSheetNames = [for sheet in bom.Worksheets -> (sheet :?> Worksheet).Name]


        let switchSmToLc3 (a2D : obj [,]) =
            let startRow = Array2D.base1 a2D
            let endRow = (Array2D.length1 a2D) - (if startRow = 0 then 1 else 0)
            let startCol = Array2D.base2 a2D
            for currentIndex = startRow to endRow do
                let lc = (string a2D.[currentIndex, startCol + 12]).Trim()
                if a2D.[currentIndex, startCol + 2] = (box "SM") && (lc = "1" || lc = "") then
                    a2D.[currentIndex, startCol + 12] <- box "3"

        

        let changeSmLoadsToLC3() =
            let loadSheetNames = workSheetNames |> List.filter (fun name -> name.Contains("L ("))
            if (List.isEmpty loadSheetNames) then
                ()
            else
                for sheet in bom.Worksheets do
                    let sheet = (sheet :?> Worksheet)
                    if sheet.Name.Contains("L (") then
                        let loads = sheet.Range("A14","M55").Value2 :?> obj [,]
                        switchSmToLc3 loads
                        sheet.Range("A14", "M55").Value2 <- loads

        let addLoadNote (mark : string) (note : string) =
            if (mark.Length > 0 && note.Length > 0) then
                let loadNote = "S" + mark
                let insertLocation = note.IndexOf(")")
                let newNote = note.Substring(0, insertLocation) + ", " + loadNote + ")"
                newNote
            else
               ""

        let removeLL_FromGirder (mark : string) (designation: string) =
            if (mark.Length > 0 && designation.Length > 0) then
                let designationArray = designation.Split([|'/'; 'K'|], StringSplitOptions.RemoveEmptyEntries)
                let newDesignation =
                    if Array.length designationArray = 3 then
                        Some (designationArray.[0] + "K" + designationArray.[2])
                    else
                        None
                match newDesignation with
                | Some _ -> ()
                | None -> System.Windows.Forms.MessageBox.Show(sprintf "Mark %s is not in TL/LL format; please fix" mark) |> ignore; ()

                match newDesignation with
                | Some s -> s
                | _ -> designation
            else
                ""


        let addLC3LoadsToLoadNotes() =
            let joists, girders, loads = bomInfo
            let joistsWithLC3Loads = joists |> List.filter (fun joist -> List.isEmpty (joist.LC3Loads loads sds) = false)
            let joistSheetNames = workSheetNames |> List.filter (fun name -> name.Contains ("J ("))
            if (List.isEmpty joistSheetNames) then ()
            else
                for sheet in bom.Worksheets do
                    let sheet = (sheet :?> Worksheet)
                    if sheet.Name.Contains("J (") then
                        let array =
                            if (sheet.Range("A21").Value2 :?> string) = "MARK" then
                                sheet.Range("A23","AA40").Value2 :?> obj [,]
                            else
                                sheet.Range("A16", "AA45").Value2 :?> obj [,]
                        let startRowIndex = Array2D.base1 array
                        let endRowIndex = (array |> Array2D.length1) - (if startRowIndex = 0 then 1 else 0) 
                        let colIndex = Array2D.base2 array
                        
                        for i = startRowIndex to endRowIndex do
                            let joistMarksWithLC3Loads =
                                joistsWithLC3Loads |> List.map (fun joist -> joist.Mark)
                            let mark = string array.[i, colIndex]
                            if (joistMarksWithLC3Loads |> List.contains mark) then
                                array.[i, colIndex + 26] <- box (addLoadNote mark (string array.[i, colIndex + 26]))
                        if (sheet.Range("A21").Value2 :?> string) = "MARK" then
                            sheet.Range("A23","AA40").Value2 <- array
                        else
                            sheet.Range("A16", "AA45").Value2 <- array

            let girdersWithLC3Loads = girders |> List.filter (fun girder -> List.isEmpty (girder.LC3Loads loads sds) = false)
            let girderWorksheetNames = workSheetNames |> List.filter (fun name -> name.Contains ("G ("))
            if (List.isEmpty girderWorksheetNames) then ()
            else
                for sheet in bom.Worksheets do
                    let sheet = (sheet :?> Worksheet)
                    if sheet.Name.Contains("G (") then
                        let array =
                            if (sheet.Range("A26").Value2 :?> string) = "MARK" then
                                sheet.Range("A28","AA45").Value2 :?> obj [,]
                            else
                                sheet.Range("A14", "AA45").Value2 :?> obj [,]

                        let startRowIndex = Array2D.base1 array
                        let endRowIndex = (array |> Array2D.length1) - (if startRowIndex = 0 then 1 else 0) 
                        let colIndex = Array2D.base2 array

                        for i = startRowIndex to endRowIndex do
                            let girderMarksWithLC3Loads =
                                girdersWithLC3Loads |> List.map (fun girder -> girder.Mark)
                            let mark = string array.[i, colIndex]
                            array.[i, colIndex + 2] <- box (removeLL_FromGirder mark (string array.[i, colIndex + 2]))
                            if (girderMarksWithLC3Loads |> List.contains mark) then
                                array.[i, colIndex + 25] <- box (addLoadNote mark (string array.[i, colIndex + 25]))
                        if (sheet.Range("A26").Value2 :?> string) = "MARK" then
                            sheet.Range("A28","AA45").Value2 <- array
                        else
                            sheet.Range("A14", "AA45").Value2 <- array



        let addLC3Loads sds =

            let addLoadSheet() =
                let workSheetNames = [for sheet in bom.Worksheets -> (sheet :?> Worksheet).Name] 
                let indexOfLastLoadSheet, lastLoadSheetNumber =
                    let lastLoadSheetNumber = workSheetNames
                                              |> List.filter (fun sheet -> sheet.Contains("L ("))
                                              |> List.map (fun sheet -> System.Int32.Parse(sheet.Split([|"(";")"|], StringSplitOptions.RemoveEmptyEntries).[1]))
                                              |> List.max
                    (bom.Worksheets.[sprintf "L (%i)" lastLoadSheetNumber] :?> Worksheet).Index, lastLoadSheetNumber
                                 

                let blankLoadWorksheet = bom.Worksheets.["L_A"] :?> Worksheet
                blankLoadWorksheet.Visible <- Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible
                blankLoadWorksheet.Copy(bom.Worksheets.[indexOfLastLoadSheet + 1])
                blankLoadWorksheet.Visible <- Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden
                let newLoadSheet = (bom.Worksheets.[indexOfLastLoadSheet + 1]) :?> Worksheet
                newLoadSheet.Name <- "L (" + string(lastLoadSheetNumber + 1) + ")"
                newLoadSheet
                
            
            let joists, girders, loads = bomInfo
          
            let joistsWithLC3Loads = joists |> List.filter (fun joist -> List.isEmpty (joist.LC3Loads loads sds) = false)
            
            let mutable row = 1
            let mutable maxJoistIndex = List.length joistsWithLC3Loads

            let mutable newLoadSheet = addLoadSheet()
            let mutable array = newLoadSheet.Range("A14", "M55").Value2 :?> obj [,]            
            
            let mutable joistIndex = 0  
                   


            while joistIndex < maxJoistIndex do
                let joist = joistsWithLC3Loads.[joistIndex]

                if row + (List.length (joist.LC3Loads loads sds)) >= 42 then
                    newLoadSheet.Range("A14", "M55").Value2 <- array.Clone()
                    newLoadSheet <- addLoadSheet()
                    array <- newLoadSheet.Range("A14", "M55").Value2 :?> obj [,]
                    row <- 1
                    joistIndex <- joistIndex - 1
                else
                    array.[row, 1] <- box ("S" + joist.Mark)


                    for load in (joist.LC3Loads loads sds) do
                        array.[row, 2] <- box (load.Type)
                        array.[row, 3] <- box (load.Category)
                        array.[row, 4] <- load.Position
                        array.[row, 6] <- load.Load1Value
                        array.[row, 7] <- load.Load1DistanceFt
                        array.[row, 8] <- load.Load1DistanceIn
                        array.[row, 9] <- load.Load2Value
                        array.[row, 10] <- load.Load2DistanceFt
                        array.[row, 11] <- load.Load2DistanceIn
                        array.[row, 12] <- load.Ref
                        array.[row, 13] <- box (load.LoadCaseString)
                        row <- row + 1
                if joistIndex = maxJoistIndex - 1 then
                    newLoadSheet.Range("A14", "M55").Value2 <- array
                joistIndex <- joistIndex + 1

            let girdersWithLC3Loads = girders |> List.filter (fun girder -> List.isEmpty (girder.LC3Loads loads sds) = false)
            
            let mutable row = 1

            let mutable maxGirderIndex = List.length girdersWithLC3Loads

            let mutable newLoadSheet = addLoadSheet()

            let mutable array = newLoadSheet.Range("A14", "M55").Value2 :?> obj [,]            
            
            let mutable girderIndex = 0   

            while girderIndex < maxGirderIndex do
                let girder = girdersWithLC3Loads.[girderIndex]

                if row + (List.length (girder.LC3Loads loads sds)) >= 42 then
                    newLoadSheet.Range("A14", "M55").Value2 <- array.Clone()
                    newLoadSheet <- addLoadSheet()
                    array <- newLoadSheet.Range("A14", "M55").Value2 :?> obj [,]
                    row <- 1
                    girderIndex <- girderIndex - 1
                else
                    array.[row, 1] <- box ("S" + girder.Mark)
                    for load in (girder.LC3Loads loads sds) do
                        array.[row, 2] <- box (load.Type)
                        array.[row, 3] <- box (load.Category)
                        array.[row, 4] <- load.Position
                        array.[row, 6] <- load.Load1Value
                        array.[row, 7] <- load.Load1DistanceFt
                        array.[row, 8] <- load.Load1DistanceIn
                        array.[row, 9] <- load.Load2Value
                        array.[row, 10] <- load.Load2DistanceFt
                        array.[row, 11] <- load.Load2DistanceIn
                        array.[row, 12] <- load.Ref
                        array.[row, 13] <- box (load.LoadCaseString)
                        row <- row + 1
                if girderIndex = maxGirderIndex - 1 then
                    newLoadSheet.Range("A14", "M55").Value2 <- array
                girderIndex <- girderIndex + 1

        changeSmLoadsToLC3()
        addLC3LoadsToLoadNotes()
        addLC3Loads sds

    let getAllBomInfo bomPath sds =
        let (joists, girders, loads) = getAllInfo bomPath getInfo modifyWorkbookFunction sds  /// warning is OK since this will always return a list of three itmes
        {
        Joists = joists
        Girders = girders
        Loads = loads
        }
















        





                       



    



    

    


