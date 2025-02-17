<#
 .Synopsis
  Converteert het aantal inches naar het aantal points

 .Description
  Converteert het aantal inches naar het aantal points
  
 .Parameter Inches
  De waarde in inches


 .Example
   # Geeft aan het aantap punten weer aan PrefferedWidth
   $Kolom.PreferredWidth = InchesToPoints $Kolombreedte

#>
Function InchesToPoints{
	param($Inches)
    $FunctionName = "InchesToPoints"
    Try{
	    return [string][math]::Round(($Kolombreedte * 72))
    }Catch{
        Write-log -Category Error -message "[$FunctionName] : Unknown error. "
        Write-log -Category Error -message "[$FunctionName] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$FunctionName] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$FunctionName] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$FunctionName] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$FunctionName] : Errormessage : $($_.Exception.message)"   
    }}
