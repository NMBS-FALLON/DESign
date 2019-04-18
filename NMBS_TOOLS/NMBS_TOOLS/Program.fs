open NMBS_Tools.DSM_Analysis
open NMBS_Tools.EmployeeReports
open NMBS_Tools.CustomerReports
open NMBS_Tools.BOM_Seismic_Seperation
open NMBS_Tools.TCWidths
open NMBS_Tools.GemToNmbs
open NMBS_Tools.DESign_Automation
open System
open Microsoft.Office.Interop.Excel
open NMBS_Tools.XML

[<EntryPoint>]
[<STAThreadAttribute>]
let main argv = 


    //FeedbackReport.sendAllFeedbackToExcel()

    //EmployeeReport.createEmployeeReport()

    //CustomerReports.createCustomerAnalysis()
    
    // Seismic Seperator
    
    
   // printfn "Please enter Sds (then click enter): "
   // let sds = float (System.Console.ReadLine())

    let seperateSeismic =
        printfn "Seperate Sesimic? (Y/N)"
        let mutable receivedCorrectInput = false
        let mutable seperateSeismic = false
        while receivedCorrectInput = false do
            let input = Console.ReadLine()
            match input with
            | "Y" | "y" -> receivedCorrectInput <- true; seperateSeismic <- true
            | "N" | "n" -> receivedCorrectInput <- true; seperateSeismic <- false
            | _ -> printfn "Incorrect Input; please try again. Seperate Sesimic? (Y/N)"
        seperateSeismic

    let checkUplift =
        printfn "Check Inward Pressure? (Y/N)"
        let mutable receivedCorrectInput = false
        let mutable checkUplift = false
        while receivedCorrectInput = false do
            let input = Console.ReadLine()
            match input with
            | "Y" | "y" -> receivedCorrectInput <- true; checkUplift <- true
            | "N" | "n" -> receivedCorrectInput <- true; checkUplift <- false
            | _ -> printfn "Incorrect Input; please try again. Check Inward Pressure? (Y/N)"
        checkUplift

    let seperatorInfo =
        {
        Seperator.SeperatorInfo.SeperateSeismic = seperateSeismic
        Seperator.SeperatorInfo.CheckInwardPressureOnGirders = checkUplift
        }

    if seperatorInfo.SeperateSeismic || seperatorInfo.CheckInwardPressureOnGirders then   
        let reportPath =
            let openFile = new System.Windows.Forms.OpenFileDialog()
            openFile.Filter <- "Excel files|*.xlsm"
            openFile.Title <- "Select BOM"
            if (openFile.ShowDialog())= (System.Windows.Forms.DialogResult.OK) then
                let fileName = openFile.FileName
                (*if fileName.Contains(".xlsm") then
                    let tempExcelApp = new Microsoft.Office.Interop.Excel.ApplicationClass(Visible = false)
                    try 
                        tempExcelApp.DisplayAlerts <- false
                        let tempReportPath = System.IO.Path.GetTempFileName()      
                        System.IO.File.Delete(tempReportPath)
                        System.IO.File.Copy(fileName, tempReportPath)
                        let workbook = tempExcelApp.Workbooks.Open(tempReportPath)    
                        let fileName = fileName.Replace(".xlsm", ".xlsx")
                        workbook.SaveAs(fileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal)
                        workbook.Close(false)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook) |> ignore
                        System.GC.Collect() |> ignore
                    finally
                        tempExcelApp.Quit()
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(tempExcelApp) |> ignore
                        System.GC.Collect() |> ignore
                        *)
                //Some (fileName.Replace(".xlsm", ".xlsx"))
                Some fileName
            else
                None

        match reportPath with
        | Some reportPath -> Seperator.seperateSeismic reportPath seperatorInfo |> ignore
                             ()
        | None -> printfn "No BOM Selected."
        
    

        //CreateReport.TCAnalysis()

        //Run.InputAllInfo()

        //NMBS_Tools.DESign_Automation.InputErfosAndDeflection.goToJoistList()
    
        //NMBS_Tools.XML.XML.xmlTest()

        //NMBS_Tools.DeckDesignCatalog.getShearValues() |> ignore

    printfn "Complete!"
    printfn "Click [ENTER] to exit."

    let s = System.Console.ReadLine()
    0 // return an integer exit code
