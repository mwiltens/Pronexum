<#
 .Synopsis
  Aan de hand van de opgegeven status, wordt er bepaald of er aan de voorwaarden wordt voldaan, om het script te starten

 .Description
  Aan de hand van de opgegeven status, wordt er bepaald of er aan de voorwaarden wordt voldaan, om het script te starten
  
 .Parameter Status
 In deze variabele staat de opgegeven status

 .Example
   # De globale variabele wordt op True gezet. Er wordt voldaan aan de voorwaarden om het script verder uit te voeren
   Update_Status $true

 .Example
   # De globale variabele wordt op False gezet. Er wordt niet voldaan aan de voorwaarden om het script verder uit te voeren
   Update_Status $False

#>
Function Update_Status{
    Param ($Status)
    If ($Global:PreReqResult -ne $False ){
        $Global:PreReqResult = $Status
    }
}
