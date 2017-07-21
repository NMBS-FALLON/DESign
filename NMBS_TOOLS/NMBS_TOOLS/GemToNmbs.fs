namespace NMBS_Tools.GemToNmbs

module Run = 
    #if INTERACTIVE
    #r "../packages/Deedle.1.2.5/lib/net40/Deedle.dll"
    #r "Microsoft.Office.Interop.Excel.dll"
    #r "../packages/FSharp.Configuration.1.3.0/lib/net45/FSharp.Configuration.dll"
    System.Environment.CurrentDirectory <- @"C:\Users\darien.shannon\Documents\Code\F#\FSharp\NMBS_TOOLS\NMBS_TOOLS\bin\Debug"
    #endif
    
    open System
    open Microsoft.Office.Interop.Excel
    open System.Runtime.InteropServices
    open System.IO
    open FSharp.Configuration


    type Joist =
        {
        Quantity : int
        Description : string
        BaseLengthFt : float
        BaseLengthIn : float
        BCExtension : bool
        PitchType : string
        Slope : float
        }

    let nullableToOption<'T> value =
        match (box value) with
        | null  -> None
        | value when value = (box "") -> None
        | _ -> Some ((box value) :?> 'T)
    
    let getInfoFunction (workBook: Workbook) =
        let takeoff = workBook.Worksheets.["Takeoff"] :?> Worksheet
        let takeoffArray = takeoff.Range("D4", "M2000").Value2 :?> obj[,]
        let startRow = Array2D.base1 takeoffArray
        let endRow = 
            match startRow with
            | 0 -> (Array2D.length1 takeoffArray) - 1
            | _ -> (Array2D.length1 takeoffArray)

        let startColumn = Array2D.base2 takeoffArray

        let splitBaseLength baseLength =
             let baseLength = if baseLength = null then "" else baseLength
             let baseLengthArray = baseLength.Split([|"."|], StringSplitOptions.RemoveEmptyEntries)
             let baseLengthFt = if baseLengthArray.Length < 1 then 0.0 else float (baseLengthArray.[0])
             let baseLengthIn = if baseLengthArray.Length < 2 then 0.0
                                else
                                    if baseLengthArray.[1] = "1" then 10.0
                                    else float (baseLengthArray.[1])
             (baseLengthFt, baseLengthIn)


        [for row = startRow to endRow do
             let (_, quantity) = Int32.TryParse(Convert.ToString(takeoffArray.[row, startColumn]))
               
             if quantity <> 0 then
                 let depth = Convert.ToString(takeoffArray.[row, startColumn + 1])
                 let series = Convert.ToString(takeoffArray.[row, startColumn + 2 ])
                 let designation = Convert.ToString(takeoffArray.[row, startColumn + 3])
                 let description = (depth + series + designation).Replace(" ", "")
                 let baseLength = Convert.ToString(takeoffArray.[row, startColumn + 4])
                 let baseLengthFt = fst (splitBaseLength baseLength)
                 let baseLengthIn = snd (splitBaseLength baseLength)
                 let bcExtension = Convert.ToString(takeoffArray.[row, startColumn + 7]).Contains("B")
                 let pitchType = Convert.ToString(takeoffArray.[row, startColumn + 8]).Trim()
                 let slope = Convert.ToDouble(takeoffArray.[row, startColumn + 9])
                 yield
                     {
                     Quantity = quantity
                     Description = description
                     BaseLengthFt = baseLengthFt
                     BaseLengthIn = baseLengthIn
                     BCExtension = bcExtension
                     PitchType = pitchType
                     Slope = slope
                     }]

    let getAllInfo (reportPath : string) (getInfoFunction : Workbook -> 'TOutput list) =
        let tempExcelApp = new Microsoft.Office.Interop.Excel.ApplicationClass(Visible = false)
        let info =
            let tempReportPath = System.IO.Path.GetTempFileName()
            File.Delete(tempReportPath)
            File.Copy(reportPath, tempReportPath)
            let workbook = tempExcelApp.Workbooks.Open(tempReportPath)
            let info = getInfoFunction workbook
            workbook.Close(false)
            Marshal.ReleaseComObject(workbook) |> ignore
            System.GC.Collect() |> ignore
            printfn "Finished processing %s." reportPath
            info
        tempExcelApp.Quit()
        Marshal.ReleaseComObject(tempExcelApp) |> ignore
        System.GC.Collect() |> ignore
        info

    type resources = ResXProvider<file="Resources.resx">
    
    let inputAllInfo (joists : Joist list) =
        printfn "Creating NMBS Takeoff; Please hold."
        let tempExcelApp = new Microsoft.Office.Interop.Excel.ApplicationClass(Visible = false)
        try
            let excelPath = System.IO.Path.GetTempFileName();
            System.IO.File.WriteAllBytes(excelPath, resources.``BLANK SALES BOM``)

            let bom = tempExcelApp.Workbooks.Open(excelPath)

            let addJoistSheet index name =                                       
                let blankJoistSheet = bom.Worksheets.["J(BLANK)"] :?> Worksheet     
                blankJoistSheet.Copy(bom.Worksheets.[index])
                let newJoistSheet = (bom.Worksheets.[index]) :?> Worksheet
                newJoistSheet.Name <- name
                newJoistSheet

            let coverIndex = (bom.Worksheets.["Cover"] :?> Worksheet).Index

            let mutable pageCount = 1
            let mutable joistSheet = addJoistSheet (coverIndex + pageCount) (sprintf "J (%i)" pageCount)

            let mutable row = 6
            let mutable markCount = 1
            for joist in joists do
                if row > 41 then
                    pageCount <- pageCount + 1
                    joistSheet <- addJoistSheet (coverIndex + pageCount) (sprintf "J (%i)" pageCount)
                    row <- 6
                joistSheet.Range("A" + row.ToString()).Value2 <- markCount.ToString()
                joistSheet.Range("B" + row.ToString()).Value2 <- joist.Quantity
                joistSheet.Range("C" + row.ToString()).Value2 <- joist.Description
                joistSheet.Range("D" + row.ToString()).Value2 <- joist.BaseLengthFt
                joistSheet.Range("E" + row.ToString()).Value2 <- joist.BaseLengthIn

                row <- row + 3
                markCount <- markCount + 1

            let savePath =
                let saveFile = new System.Windows.Forms.SaveFileDialog()
                saveFile.Filter <- "Excel files (*.xlsm)|*.xlsx"
                saveFile.Title <- "Save File"
                if (saveFile.ShowDialog()) = (System.Windows.Forms.DialogResult.OK) then
                    let path =
                        if saveFile.FileName.Contains(".xlsm") then saveFile.FileName
                        else saveFile.FileName + ".xlsm"
                    Some path
                else
                    None

            match savePath with
            | Some path ->bom.SaveAs(path)
            | None -> ()

            bom.Close()
            Marshal.ReleaseComObject(bom) |> ignore
            System.GC.Collect() |> ignore

        finally
            tempExcelApp.Quit()
            Marshal.ReleaseComObject(tempExcelApp) |> ignore
            System.GC.Collect() |> ignore    

    let InputAllInfo() =
        let reportPath =
            let openFile = new System.Windows.Forms.OpenFileDialog()
            openFile.Filter <- "Excel files (*.xls)|*.xls"
            openFile.Title <- "Select BOM"
            if (openFile.ShowDialog())= (System.Windows.Forms.DialogResult.OK) then
                Some openFile.FileName
            else
                None

        match reportPath with
        | Some reportPath -> inputAllInfo (getAllInfo reportPath getInfoFunction)
                             ()
        | None -> printfn "No BOM Selected."

    InputAllInfo()

    

                


