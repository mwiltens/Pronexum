Function TypeTekst{
    param($Object,$bold=$False, $Italic=$False, $Underline=$False, $Fontsize, $Fontname,  $Tekst,  $Enter=$False)
    $FunctionName = "TypeTekst"
    Write-log Verbose -Message "`$Tekst = `'$Tekst`'"
    Write-log Verbose -Message "`$bold = `'$bold`'"
    Write-log Verbose -Message "`$Italic = `'$Italic`'"
    Write-log Verbose -Message "`$Underline = `'$Underline`'"
    Try{
        If ($Bold){
            $Object.Selection.Font.Bold = $wdToggle
        }
        If ($Italic){
            $Object.Selection.Font.Italic = $WDToggle
        }
        If ($Underline){
            If($Object.Selection.Font.Underline -eq $WdUnderlineNone){
                $Object.Selection.Font.Underline = $WdUnderlineSingle
            }else{
                $Object.Selection.Font.Underline = $WdUnderlineNone
            }
        }
        If ($FontName){
            $Object.Selection.Font.Name = $FontName
        }
        If ($FontSize){
            $Object.Selection.Font.Size = $FontSize
        }
            $Object.Selection.TypeText("$Tekst")
            $Object.Selection.Font.Italic = $False
            $Object.Selection.Font.Bold = $False
                $Object.Selection.Font.Underline = $WdUnderlineNone
        if ($Enter){
            $Object.Selection.TypeParagraph()
        }
    }Catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "
        #Write-log -Category Error -message "[$Functionname] : Targetname   : $(_.CategoryInfo.targetname)"
        #Write-log -Category Error -message "[$Functionname] : Fullname     : $(_.exception.gettype.Fullname)"
        #Write-log -Category Error -message "[$Functionname] : Type fout    : $(_.CategoryInfo.category)"
        #Write-log -Category Error -message "[$Functionname] : Position     : $(_.invocationinfo.positionmessage)"
        #Write-log -Category Error -message "[$Functionname] : Errormessage : $(_.Exception.message)"
    }
}