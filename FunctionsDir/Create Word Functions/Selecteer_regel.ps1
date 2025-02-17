<#
 .Synopsis
  Selecteer een regel

 .Description
  Er wordt een regel in het Word-document geselecteerd
  
 .Parameter Object
  Is het Object, waarop de bewerking wordt toegepast

.Example
   # Selecteer een regel 
   Selecteer_Regel -Object $Word 


#>

Function Selecteer_regel { 
    Param ($Object)
    $FunctionName = "Selecteer_regel"
    Try{
        $Object.Selection.EndKey($wdLine,$wdExtend) | Out-Null
    }Catch{
        Write-log -Category Error -message "[$FunctionName] : Unknown error. "
        Write-log -Category Error -message "[$FunctionName] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$FunctionName] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$FunctionName] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$FunctionName] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$FunctionName] : Errormessage : $($_.Exception.message)"   
    }}


