<#
 .Synopsis
  Past de opgegeven stijl toe

 .Description
  Past de opgegeven stijl toe
  
 .Parameter Object
  Is het Object, waarop de bewerking wordt toegepast

 .Parameter Style
  De Style die de selectie moet gaan krijgen

.Example
   # Pas de style "Kop 1" toe op de opgegeven regel
   Style -Object $Word -Style "Kop 1"
#>
Function Style{
	Param ($Object, $Style)
    $FunctionName = "Style"
    Try{
        $Style = CheckStyle $Style
	    $Object.Selection.Style = $Style	
        $Object.Selection.MoveDown($wdLine, 1) | Out-Null
    }Catch{
        Write-log -Category Error -message "[$FunctionName] : Unknown error. "
        Write-log -Category Error -message "[$FunctionName] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$FunctionName] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$FunctionName] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$FunctionName] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$FunctionName] : Errormessage : $($_.Exception.message)"   
    }
}