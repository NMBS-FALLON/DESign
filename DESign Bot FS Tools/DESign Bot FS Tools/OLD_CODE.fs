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
    open System.Text.RegularExpressions



        
    let getLoadNotes (note : string) =
        if note.Contains("(") then
            let loadNoteStart = note.IndexOf("(")
            let loadNotes = note.Substring(loadNoteStart)
            let loadNotes = loadNotes.Split([|"("; ","; ")"|], StringSplitOptions.RemoveEmptyEntries)
            let loadNotes = loadNotes |> List.ofArray
            loadNotes |> List.map (fun (s : string) -> s.Trim())
        else
            []

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

    type Note =
        {
        Number : string
        Text : string
        }




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




    
    type Girder =

        member this.UDL_PDL (liveLoadUNO : string) (liveLoadSpecialNotes : Note List)=
            let size = this.GirderSize
            let sizeAsArray = size.Split( [|"G"; "BG"; "VG"; "N"; "K"|], StringSplitOptions.RemoveEmptyEntries)
            let load = sizeAsArray.[2]
            let minSpace =
                let geometry = this.GirderGeometry
                let aSpace = geometry.A_Ft + geometry.A_In / 12.0
                let bSpace = geometry.B_Ft + geometry.B_In / 12.0
                let minPanelSpace = List.min (geometry.Panels |> List.map (fun geom -> geom.LengthFt + geom.LengthIn / 12.0))
                List.min [aSpace; bSpace; minPanelSpace]

            let TL = float load

            let liveLoadSpecialNote =
                match this.SpecialNoteList with
                | Some list -> 
                    [for liveLoadSpecialNote in liveLoadSpecialNotes do
                         if List.contains liveLoadSpecialNote.Number list then
                             yield liveLoadSpecialNote.Text]
                | None -> []

            let liveLoad =
                match liveLoadSpecialNote with
                | [] -> liveLoadUNO
                | _ -> liveLoadSpecialNote.[0] 
                     

            let LL =
                match liveLoad with
                | Regex @" *[LS] *= *(\d+\.?\d*) *[Kk] *" [value] -> float value
                | Regex @" *[LS] *= *(\d+\.?\d*) *% *" [percent] ->
                    let fraction = float percent/100.0
                    TL*fraction
                | _ -> 0.0

            let DL = 1000.0 * (TL - LL)
            let UDL = DL / minSpace
            UDL, DL


        member this.SDS sds =
            let udl, _ = this.UDL_PDL this.LiveLoadUNO this.LiveLoadSpecialNotes
            let SDS = udl * 0.14 * sds
            Load.create("U", "SM", "TC", SDS,
                          null, null, null, null, null, null, [3])

        member this.DeadLoads =
            let _,dl = this.UDL_PDL this.LiveLoadUNO this.LiveLoadSpecialNotes
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

            let addLiveLoadInfoToGirders (girders: Girder list, liveLoadUNO : string, liveLoadSpecialNotes : Note list) =
                [for girder in girders do
                    yield {girder with LiveLoadUNO = liveLoadUNO; LiveLoadSpecialNotes = liveLoadSpecialNotes}]

                    
    let saveWorkbook (title : string) (workbook : Workbook) =
            let title = title.Replace(".xlsm", " (IMPORT).xlsm")
            let title = title.Replace(".xlsx", " (IMPORT).xlsx")
            workbook.SaveAs(title)
    
    let getAllInfo (reportPath:string) getInfoFunction modifyWorkbookFunctions =
        let tempExcelApp = new Microsoft.Office.Interop.Excel.ApplicationClass(Visible = false)
        tempExcelApp.DisplayAlerts = false |> ignore
        tempExcelApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable |> ignore
        //let mutable workbook = tempExcelApp.Workbooks.Add()
        
        //let bom = tempExcelApp.Workbooks.Open(bomPath)
        try 
            tempExcelApp.DisplayAlerts <- false
            let tempReportPath = System.IO.Path.GetTempFileName()      
            File.Delete(tempReportPath)
            File.Copy(reportPath, tempReportPath)
            let workbook = tempExcelApp.Workbooks.Open(tempReportPath) 
            let info = getInfoFunction workbook
            for modifyWorkbookFunction in modifyWorkbookFunctions do
                modifyWorkbookFunction workbook info
            
            workbook |> saveWorkbook reportPath

            printfn "Finished processing %s." reportPath 
            printfn "Finished processing all files."
            info
        finally
           // workbook.Close(false)
           // Marshal.ReleaseComObject(workbook) |> ignore
            System.GC.Collect() |> ignore
            tempExcelApp.Quit()
            Marshal.ReleaseComObject(tempExcelApp) |> ignore
            System.GC.Collect() |> ignore            

    module Modifiers =
        let seperateSeismic (bom : Workbook) (bomInfo : Joist list * Girder list * LoadNote list * float) : Unit =

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
    (*
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
    *)

            let addLC3LoadsToLoadNotes() =
                let joists, girders, loads, SDS = bomInfo
                let joistsWithLC3Loads = joists |> List.filter (fun joist -> List.isEmpty (joist.LC3Loads loads SDS) = false)
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

                let girdersWithLC3Loads = girders |> List.filter (fun girder -> List.isEmpty (girder.LC3Loads loads SDS) = false)
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
                                //array.[i, colIndex + 2] <- box (removeLL_FromGirder mark (string array.[i, colIndex + 2]))
                                if (girderMarksWithLC3Loads |> List.contains mark) then
                                    array.[i, colIndex + 25] <- box (addLoadNote mark (string array.[i, colIndex + 25]))
                            if (sheet.Range("A26").Value2 :?> string) = "MARK" then
                                sheet.Range("A28","AA45").Value2 <- array
                            else
                                sheet.Range("A14", "AA45").Value2 <- array



            let addLC3Loads()=

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
                
            
                let joists, girders, loads, SDS = bomInfo
          
                let joistsWithLC3Loads = joists |> List.filter (fun joist -> List.isEmpty (joist.LC3Loads loads SDS) = false)
            
                let mutable row = 1
                let mutable maxJoistIndex = List.length joistsWithLC3Loads

                let mutable newLoadSheet = addLoadSheet()
                let mutable array = newLoadSheet.Range("A14", "M55").Value2 :?> obj [,]            
            
                let mutable joistIndex = 0  
                   


                while joistIndex < maxJoistIndex do
                    let joist = joistsWithLC3Loads.[joistIndex]

                    if row + (List.length (joist.LC3Loads loads SDS)) >= 42 then
                        newLoadSheet.Range("A14", "M55").Value2 <- array.Clone()
                        newLoadSheet <- addLoadSheet()
                        array <- newLoadSheet.Range("A14", "M55").Value2 :?> obj [,]
                        row <- 1
                        joistIndex <- joistIndex - 1
                    else
                        array.[row, 1] <- box ("S" + joist.Mark)


                        for load in (joist.LC3Loads loads SDS) do
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

                let girdersWithLC3Loads = girders |> List.filter (fun girder -> List.isEmpty (girder.LC3Loads loads SDS) = false)
            
                let mutable row = 1

                let mutable maxGirderIndex = List.length girdersWithLC3Loads

                let mutable newLoadSheet = addLoadSheet()

                let mutable array = newLoadSheet.Range("A14", "M55").Value2 :?> obj [,]            
            
                let mutable girderIndex = 0   

                while girderIndex < maxGirderIndex do
                    let girder = girdersWithLC3Loads.[girderIndex]

                    if row + (List.length (girder.LC3Loads loads SDS)) >= 42 then
                        newLoadSheet.Range("A14", "M55").Value2 <- array.Clone()
                        newLoadSheet <- addLoadSheet()
                        array <- newLoadSheet.Range("A14", "M55").Value2 :?> obj [,]
                        row <- 1
                        girderIndex <- girderIndex - 1
                    else
                        array.[row, 1] <- box ("S" + girder.Mark)
                        for load in (girder.LC3Loads loads SDS) do
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
            addLC3Loads()

        let adjustSinglePitchJoists (bom : Workbook) (bomInfo : Joist list * Girder list * LoadNote list * float) : Unit =
            ()

    let seperateSeismic bomPath =
        getAllInfo bomPath getInfo [Modifiers.seperateSeismic]

    let seperateSeismicAndAdjustSinglePitches bomPath =
        getAllInfo bomPath getInfo [Modifiers.seperateSeismic; Modifiers.adjustSinglePitchJoists]

    let adjustSinglePitchJoists bomPath =
        getAllInfo bomPath getInfo [Modifiers.adjustSinglePitchJoists]
















        





                       



    



    

    


