open NMBS_Tools.DSM_Analysis
open NMBS_Tools.EmployeeReports
open NMBS_Tools.CustomerReports
open NMBS_Tools.BOM_Seismic_Seperation
open NMBS_Tools.TCWidths
open NMBS_Tools.GemToNmbs
open System
open Microsoft.Office.Interop.Excel

[<EntryPoint>]
[<STAThreadAttribute>]
let main argv = 

    //FeedbackReport.sendAllFeedbackToExcel()

    //EmployeeReport.createEmployeeReport()

    //CustomerReports.createCustomerAnalysis()
    
    (*
    printfn "Please enter Sds (then click enter): "
    let sds = float (System.Console.ReadLine())

    
    let reportPath =
        let openFile = new System.Windows.Forms.OpenFileDialog()
        openFile.Filter <- "Excel files|*.xlsx"
        openFile.Title <- "Select BOM"
        if (openFile.ShowDialog())= (System.Windows.Forms.DialogResult.OK) then
            let fileName = openFile.FileName
            if fileName.Contains(".xlsm") then
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
            Some (fileName.Replace(".xlsm", ".xlsx"))
        else
            None

    match reportPath with
    | Some reportPath -> Seperator.getAllBomInfo reportPath sds |> ignore
                         ()
    | None -> printfn "No BOM Selected."
        
    *)

    //CreateReport.TCAnalysis()

    Run.InputAllInfo()

    printfn "Complete!"
    printfn "Click enter to exit."

    let s = System.Console.ReadLine()
    0 // return an integer exit code
