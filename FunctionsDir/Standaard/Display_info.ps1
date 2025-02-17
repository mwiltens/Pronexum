<#
 .Synopsis
  Plaats een omschrijving met de opgegeven waarde op het scherm

 .Description
  Plaats een omschrijving met de opgegeven waarde op het scherm
  
 .Parameter Omschrijving
  De omschrijving die op het scherm moet worden getoond

 .Parameter Waarde
  De waarde die achter de omschrijving moet worden getoond

 .Parameter Color
  De kleur die de de getoond tekst moet gaan krijgen


 .Example
   # De omschrijving "Time start" wordt op het scherm getoond met de bijbehoorde waarde in het groen
   Display_info -Omschrijving "Time start" -Waarde $time -Color Green

#>
Function Display_Info{
    Param ($Omschrijving, $Waarde, [ValidateSet("Black","Blue" ,"Cyan" ,"DarkBlue","DarkCyan","DarkGray","DarkGreen","DarkMagenta","DarkRed","DarkYellow","Gray","Green","Magenta","White","Red","Yellow")]$Color)  
    $Functionname = "Display_Info"
    Try{
        $Message = "$Omschrijving".padright($tab," ") + ": $Waarde"
        Write-log -Category host -message $Message -color $Color
    }Catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "   
        Write-log -Category Error -message "[$Functionname] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $($_.Exception.message)"   
    }
}
