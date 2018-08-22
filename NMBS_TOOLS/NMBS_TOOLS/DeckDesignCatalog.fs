namespace NMBS_Tools

    #if INTERACTIVE
    #r "Microsoft.Office.Interop.Excel.dll"
    #endif

module DeckDesignCatalog =

    type SupportFastener =
       | ArcSpotWelds
     //  | ArcSpotWeldsDDM03
     //  | ArcSeamWelds
       | Screws
     //  | ScrewsDDM03
       | PowerDriven

       member this.Value =
           match this with
              | ArcSpotWelds       -> "Arc Spot Welds"
            //  | ArcSpotWeldsDDM03  -> "Arc Spot Welds (DDM03)"
           //   | ArcSeamWelds       -> "Arc Seam Welds"
              | Screws             -> "Screws"
           //   | ScrewsDDM03        -> "Screws (DDM03)"
              | PowerDriven        -> "Power Driven"

       member this.Options =
           match this with
            | ArcSpotWelds      -> ["0.5"]
          //  | ArcSpotWeldsDDM03 -> ["0.5"]
          //  | ArcSeamWelds      -> ["0.250 x 1.00" (*; "0.375 x 1.25"*)]
            | Screws            -> ["#12"; "#14"]
          //  | ScrewsDDM03       -> ["#12"; "#14"]
            | PowerDriven       -> [(*"Hilti X-ENP-19L15";*) "Hilti ENP2K, X-EDN19, X-EDNK22 or X-HSN24"]

       static member List =
            [ArcSpotWelds; (*ArcSpotWeldsDDM03; ArcSeamWelds;*) Screws; (*ScrewsDDM03;*) PowerDriven]


    type SideLapFastener =
        | ArcSpotWelds
     //   | ArcSpotWeldsDDM03
        | ArcSeamWelds
        | FilletWelds
        | Screws
     //   | ScrewsDDM03
        | TopArcSeamSideLapWelds
        | ButtonPunch

        member this.Value =
            match this with
            | ArcSpotWelds       -> "Arc Spot Welds"
         //   | ArcSpotWeldsDDM03  -> "Arc Spot Welds (DDM03)"
            | ArcSeamWelds       -> "Arc Seam Welds"
            | FilletWelds        -> "Fillet Welds"
            | Screws             -> "Screws"
        //    | ScrewsDDM03        -> "Screws (DDM03)"
            | TopArcSeamSideLapWelds -> "Top Arc Seam Side Lap Welds"
            | ButtonPunch        -> "Button Punch"

        member this.Options =
            match this with
            | ArcSpotWelds       -> ["0.5" (*; "0.625"; "0.75"; "1" *)]
       //     | ArcSpotWeldsDDM03  -> ["0.5"; "0.625"; "0.75"; "1"]
            | ArcSeamWelds       -> []
            | FilletWelds        -> ["1.5"]
            | Screws             -> ["#10"; "#12"]
       //     | ScrewsDDM03        -> ["#10"; "#12"]
            | TopArcSeamSideLapWelds -> ["1.5"]
            | ButtonPunch        -> ["Generic"]


    type DeckProfile =
        | B
        | BI
        | N

        member this.Value =
            match this with
            | B -> "B"
            | BI -> "BI"
            | N -> "N"

        member this.Gauges =
            let standard = ["22"; "20"; "18"; "16"]
            match this with
            | B -> standard
            | BI -> standard
            | N -> standard

        member this.Yields =
            let standard = ["33"; "40"; "50"; "80"]
            match this with
            | B -> standard
            | BI -> standard
            | N -> standard

        member this.SidelapFasteners =
            let nestableTypes = [SideLapFastener.ArcSpotWelds; (*ArcSpotWeldsDDM03;*) ArcSeamWelds; FilletWelds; Screws (*; ScrewsDDM03*)]
            let interlockingTypes = [SideLapFastener.TopArcSeamSideLapWelds; ButtonPunch; Screws; (*ScrewsDDM03*)]
            match this with
            | B -> nestableTypes
            | BI -> interlockingTypes
            | N -> nestableTypes

        member this.SupportPattern =
            let _36 = ["36/9"; "36/7"; "36/5"]
            let _24 = ["24/4"]
            match this with
            | B -> _36
            | BI -> _36
            | N -> _24


    type DeckConfiguration = 
        {
            deckProfile : string
            deckGauge : string
            deckYield : string
            supportFastener : string
            supportFastenerOption : string
            sideLapFastener : string
            sideLapFastenerOption : string
            supportPattern : string
            spanFt : string
            spanIn : string
            sideLapSpace : string
        }

 (*   let allDeckConfigurations =
        let deckTypes = [B; BI; N]
        let spans =
            [for ft in 2..1..14 do
                for inch in [0;6] do
                    yield (string ft, string inch) ]

        let sideLapSpaces = [for space in 36..-3..3 do yield string space ] 


        [for deckType in deckTypes do
            for gauge in deckType.Gauges do
                for deckYield in deckType.Yields do
                    for supportFastener in SupportFastener.List do
                        for supportFastenerOption in supportFastener.Options do
                            for sideLapFastener in deckType.SidelapFasteners do
                                for sideLapFastenerOption in sideLapFastener.Options do
                                    for supportPattern in deckType.SupportPattern do
                                        for span in spans do
                                            for sideLapSpace in sideLapSpaces do
                                                yield
                                                    {
                                                    deckProfile = deckType.Value
                                                    deckGauge = gauge
                                                    deckYield = deckYield
                                                    supportFastener = supportFastener.Value
                                                    supportFastenerOption = supportFastenerOption
                                                    sideLapFastener = sideLapFastener.Value
                                                    sideLapFastenerOption = sideLapFastenerOption
                                                    supportPattern = supportPattern
                                                    spanFt = fst span
                                                    spanIn = snd span
                                                    sideLapSpace = sideLapSpace
                                                    } ]
                                                    *)
    type DeckConfigurationWithShear =
        {
            deckProfile : string
            deckGauge : string
            deckYield : string
            supportFastener : string
            supportFastenerOption : string
            sideLapFastener : string
            sideLapFastenerOption : string
            supportPattern : string
            span : string
            sideLapSpace : string
            shearE : int
            shearW : int
            shearOther : int
            shearBuckling : int
        }


    open Microsoft.Office.Interop.Excel
    open System.Text.RegularExpressions

    let (|Regex|_|) pattern input =
        let m = Regex.Match(input, pattern)
        if m.Success then Some(List.tail [ for g in m.Groups -> g.Value ])
        else None

        
    let getShearValues() =
        // DONT FORGET TO SET B16 TO "EFFECTIVE DIAMETER" FOR ARC SPOT WELDS
        // DONT FORGET TO SET L28 TO "SPACING"


        let tempExcelApp = new Microsoft.Office.Interop.Excel.ApplicationClass(Visible = false)
        let workbook = tempExcelApp.Workbooks.Open("C:\\Users\\darien.shannon\\Desktop\\Random Projects\\Deck Values\\Deck.xlsm")
        let workSheet = workbook.Sheets.["DIAPHRAGM"] :?> Worksheet
        workSheet.Range("L28", "L28").Value2 <- "Spacing"

        

        let deckConfigurationsWithShear = 

            let deckProfiles = [B; BI; N]
            let spans =
                [for ft in 2..1..14 do
                    for inch in [0;6] do
                        yield (string ft, string inch) ]

            let sideLapSpaces = [for space in 36..-3..3 do yield string space ]

            let mutable performanceCount = 0

            [for deckProfile in deckProfiles do
                workSheet.Range("F1", "F1").Value2 <- deckProfile.Value
                for deckGauge in deckProfile.Gauges do
                    workSheet.Range("F5", "F5").Value2 <- deckGauge
                    for deckYield in deckProfile.Yields do
                        workSheet.Range("F7", "F7").Value2 <- deckYield
                        for supportFastener in SupportFastener.List do
                            workSheet.Range("C15", "C15").Value2 <- supportFastener.Value
                            if supportFastener.Value = "Arc Spot Welds" then
                                    workSheet.Range("A16", "A16").Value2 <- "Effective Diameter:"
                            for supportFastenerOption in supportFastener.Options do
                                workSheet.Range("C16", "C16").Value2 <- supportFastenerOption
                                for sideLapFastener in deckProfile.SidelapFasteners do
                                    workSheet.Range("C18", "F18").Value2 <- sideLapFastener.Value
                                    for sideLapFastenerOption in sideLapFastener.Options do
                                        workSheet.Range("C19", "C19").Value2 <- sideLapFastenerOption
                                        for supportPattern in deckProfile.SupportPattern do
                                            workSheet.Range("Y25", "Y25").Value2 <- supportPattern

                                            let omegaEText = string (workSheet.Range("AO33", "AO33").Text)
                                            let omegaWText = string (workSheet.Range("AO34", "AO34").Text)
                                            let omegaOtherText = string (workSheet.Range("AO35", "AO35").Text)
                                            let omegaBucklingText = string (workSheet.Range("AO36", "AO36").Text)

                                            let omegaValue omegaValueText=
                                                let omegaValue =
                                                    match omegaValueText with
                                                    | Regex @".*[:] (\d+\.?\d*)" [value] -> float value
                                                    | _ -> -1.0
                                                omegaValue

                                            let omegaE = omegaValue omegaEText
                                            let omegaW = omegaValue omegaWText
                                            let omegaOther = omegaValue omegaOtherText
                                            let omegaBuckling = omegaValue omegaBucklingText

                                            for initialSpan in [("3","0");("10","0")] do
                                                workSheet.Range("C39", "C39").Value2 <- fst initialSpan
                                                workSheet.Range("E39", "E39").Value2 <- snd initialSpan
                                                for initialSideLapCriteria in [("36","3");("5","1")] do
                                                    workSheet.Range("B41", "B41").Value2 <- fst initialSideLapCriteria
                                                    workSheet.Range("L29", "L29").Value2 <- snd initialSideLapCriteria
                                                    let sndCell = if (fst initialSideLapCriteria) = "36" then "BN22" else "BN16"

                                                    let shearArray = workSheet.Range("AZ11", sndCell).Value2 :?> obj [,]

                                                    let indexFstRow = Array2D.base1 shearArray
                                                    let indexLastRow = if indexFstRow = 0 then
                                                                          (Array2D.length1 shearArray) - 1
                                                                       else
                                                                           (Array2D.length1 shearArray) 

                                                    let indexFstCol = Array2D.base2 shearArray
                                                    let indexLastCol = if indexFstCol = 0 then
                                                                          (Array2D.length2 shearArray) - 1
                                                                       else
                                                                           (Array2D.length2 shearArray)
                                                                         
                                                    for row in (indexFstRow + 1)..indexLastRow do
                                                        let sideLapSpacing = shearArray.[row,indexFstCol]
                                                        
                                                        for col in (indexFstCol + 1)..indexLastCol do
                                                            let span = shearArray.[indexFstCol, col]

                                                            let nominalShearText = shearArray.[row, col]
             
                                                            let nominalShear =
                                                                match (string nominalShearText) with
                                                                | Regex @"(\d+\.?\d*)" [value] -> float value
                                                                | _                            -> -1.0
                                                            
                                                            let shearValue (omegaValue:float) = int (System.Math.Floor (nominalShear / omegaValue))



                                                            let deckConfigurationWithShears =                                                     
                                                                {
                                                                    deckProfile = deckProfile.Value
                                                                    deckGauge = deckGauge
                                                                    deckYield = deckYield
                                                                    supportFastener = supportFastener.Value
                                                                    supportFastenerOption = supportFastenerOption
                                                                    sideLapFastener = sideLapFastener.Value
                                                                    sideLapFastenerOption = sideLapFastenerOption
                                                                    supportPattern = supportPattern
                                                                    span =  string span
                                                                    sideLapSpace = string sideLapSpacing
                                                                    shearE = shearValue omegaE
                                                                    shearW = shearValue omegaW
                                                                    shearOther = shearValue omegaOther
                                                                    shearBuckling = shearValue omegaBuckling
                                                                }

                                                            performanceCount <- performanceCount + 1
                                                            printfn "%i" performanceCount

                                                            yield deckConfigurationWithShears]
                                                    
        let numRows = List.length deckConfigurationsWithShear
        let numCols = 14

        let resultArray : obj[,]= Array2D.create numRows numCols null

        for row in 0..(numRows - 1) do
            let deckConfiguration = deckConfigurationsWithShear.[row]
            resultArray.[row,0] <- box deckConfiguration.deckProfile
            resultArray.[row,1] <- box deckConfiguration.deckGauge
            resultArray.[row,2] <- box deckConfiguration.deckYield
            resultArray.[row,3] <- box deckConfiguration.supportFastener
            resultArray.[row,4] <- box deckConfiguration.supportFastenerOption
            resultArray.[row,5] <- box deckConfiguration.supportPattern
            resultArray.[row,6] <- box deckConfiguration.sideLapFastener
            resultArray.[row,7] <- box deckConfiguration.sideLapFastenerOption
            resultArray.[row,8] <- box deckConfiguration.sideLapSpace
            resultArray.[row,9] <- box deckConfiguration.span
            resultArray.[row,10] <- box deckConfiguration.shearE
            resultArray.[row,11] <- box deckConfiguration.shearW
            resultArray.[row,12] <- box deckConfiguration.shearOther
            resultArray.[row,13] <- box deckConfiguration.shearBuckling


        let resultSheet = workbook.Sheets.["RESULT"] :?> Worksheet

        let bottomRightCell = "O" + (string ((Array2D.length1 resultArray) + 1))

        resultSheet.Range("A1", bottomRightCell).Value2 <- resultArray

        workbook.Save()




        






                                        


    
