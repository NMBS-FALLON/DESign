namespace NMBS_Tools.DESign_Automation

module InputErfosAndDeflection =
    #if INTERACTIVE
    System.Environment.CurrentDirectory <- @"C:\Users\darien.shannon\Desktop\DESign\DESign\NMBS_TOOLS\NMBS_TOOLS\bin\Debug"
    #r @"C:\Users\darien.shannon\Desktop\DESign\DESign\NMBS_TOOLS\NMBS_TOOLS\bin\Debug\AutoItX3.Assembly.dll"
    #endif

    open AutoIt
    open System.Windows.Forms

    AutoItX.AutoItSetOption("WinTitleMatchMode", 2) |> ignore
    AutoItX.AutoItSetOption("MouseCoordMode", 0) |> ignore

    let waitForActiveWindow (windowTitle:string) =
        let i = 1
        while i = 1 do
            if AutoItX.WinActive(windowTitle) = 1 then
                i = 0 |> ignore


    let goToJoistList() =
        let madeItThrough = false
        let test = AutoItX.WinExists("Joist Design")
        let test2 = AutoItX.WinActive("Joist Design")
        if AutoItX.WinExists("Joist Design") = 0 then
            MessageBox.Show("JEDI IS NOT OPEN; EXITING SCRIPT") |> ignore
        else
            AutoItX.WinActivate("Joist Design")  |> ignore
            AutoItX.Sleep(300)
            if AutoItX.WinActive("Joist Design") = 0 then
                MessageBox.Show("PLEASE ACTIVATE JEDI") |> ignore
            waitForActiveWindow("Joist Design")



    

    



