<#
 .Synopsis
  Maakt een inhoudsopgave aan

 .Description
  Maakt een inhoudsopgave aan in MS-Word
  
 .Parameter Object
  Is het Object, waarop de bewerking wordt toegepast

 .Example
   # Aanmaken van een inhoudsopgave in Word.
    Aanmaken_Inhoudsopgave -Object Word

#>
Function Aanmaken_Inhoudsopgave{
	Param ($Object, $Style, $UpperHeadingLevel=3, $LowerHeadingLevel = 1)
    $FunctionName = "Aanmaken_Inhoudsopgave"
    Try{
        $Selection = $Object.Selection
        $Range = $Selection.Range
        $GLOBAL:TOC = $Global:Doc.TablesOfContents.Add($Range) #,$UpperHeadingLevel,$LowerHeadingLevel) | outnull
        $Object.Selection.TypeParagraph()| Out-Null
    }Catch{
        Write-log -Category Error -message "[$FunctionName] : Unknown error. "
        Write-log -Category Error -message "[$FunctionName] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$FunctionName] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$FunctionName] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$FunctionName] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$FunctionName] : Errormessage : $($_.Exception.message)"   
    }}