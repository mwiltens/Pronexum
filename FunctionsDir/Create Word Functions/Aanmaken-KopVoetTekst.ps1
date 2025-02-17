<#
 .Synopsis
  Aanmaken van een kop of een voettekst

 .Description
  Aanmaken van een kop of een voettekst
  
 .Parameter Object


 .Parameter Type
  In Type wordt opgegeven of het een kop of een voettekst moet worden


 .Parameter Picture
  In Picture wordt het image opgegeven wat wordt toegevoegd aan de koptekst

 .Example
   # Aanmaken van een koptekst waar het jpg bestand Edulas.jpg aan wordt toegevoegd
    Aanmaken-KopVoetTekst -Object $Word -Type Koptekst -Picture "E:\Data\Sjablonen\Edulas.jpg"
    Aanmaken-KopVoetTekst -Object $Word -Type Voettekst 

 .Example
   # Aanmaken van een voettekst 
    Aanmaken-KopVoetTekst -Object $Word -Type Voettekst 

#>
Function Aanmaken-KopVoetTekst{
    Param ($Object,[validateset ("Koptekst","Voettekst")]$Type,$Picture)
    $FunctionName = "Aanmaken-KopVoetTekst"
    $wdAlignParagraphRight = 2
    $wdAlignParagraphLeft = 0
    switch ($Type)
    {
        "Koptekst" 
            {
                Try{
                    $Object.ActiveWindow.ActivePane.View.SeekView = $wdSeekCurrentPageHeader
                    Maak_Kolommen -Object $OBject -Rijen 1 -Kolommen 2 -Borderlines False -BorderBottom $true
                    If (!([string]::IsNullOrEmpty($Picture))){
                        Move-Cursor -Object $Object -Direction Right -Unit $wdCell
                        $Object.Selection.ParagraphFormat.Alignment = $wdAlignParagraphRight
                        $object.Selection.InlineShapes.AddPicture($Picture) | Out-Null
                    }
                    $Object.ActiveWindow.ActivePane.View.SeekView = $wdSeekMainDocument
                }Catch{
                    Write-log -Category Error -message "[$FunctionName] : Error bij het aanmaken van een koptekst."
                    Write-log -Category Error -message "[$FunctionName] : Targetname   : $($_.CategoryInfo.targetname)"
                    Write-log -Category Error -message "[$FunctionName] : Fullname     : $($_.exception.gettype().Fullname)"
                    Write-log -Category Error -message "[$FunctionName] : Type fout    : $($_.CategoryInfo.category)"
                    Write-log -Category Error -message "[$FunctionName] : Position     : $($_.invocationinfo.positionmessage)"
                    Write-log -Category Error -message "[$FunctionName] : Errormessage : $($_.Exception.message)"   
                }            }
        "Voettekst"
            {
                Try{
                    $Object.ActiveWindow.ActivePane.View.SeekView = $wdSeekCurrentPageFooter
                    Maak_Kolommen -Object $OBject -Rijen 1 -Kolommen 2 -Borderlines False -BorderTop $true
                    $Selection = $Global:WORD.Selection
                    $Range = $Selection.Range
                    typetekst -tekst "" -Object $WORD -bold $False -Fontsize 9 -Fontname "Times New Roman" -Enter $False
                    $Global:Doc.fields.add($Range, 29, "\* FirstCap \p ", $true)  | Out-Null
                    Move-Cursor -Object $Object -Direction Right -Unit $wdCell
                    $Object.Selection.ParagraphFormat.Alignment = $wdAlignParagraphRight
                    #$object.selection.TypeParagraph()
                    typetekst -tekst "Bladzijde " -Object $WORD -bold $False -Fontsize 9 -Fontname "Times New Roman" -Enter $False
                    $Selection = $Global:WORD.Selection
                    $Range = $Selection.Range
                    $Global:Doc.fields.add($Range, 33, "", $true) | Out-Null
                    typetekst -tekst " van " -Object $WORD -bold $False -Fontsize 9 -Fontname "Times New Roman" -Enter $False
                    $Selection = $Global:WORD.Selection
                    $Range = $Selection.Range
                    $Global:Doc.fields.add($Range, 26, "", $true) | Out-Null

                    $wdFieldFileName = 29
                    $wdFieldNumPages = 26
                    $wdFieldPage = 33
                    $wdFieldCreateDate = 21
                    #Maak_Kolommen -Object $OBject -Rijen 1 -Kolommen 2 -Borderlines True                 
                    #Move-Cursor -Object $Object -Direction Right -Unit $wdCell
                    #$Object.Selection.ParagraphFormat.Alignment = $wdAlignParagraphRight
                    #Move-Cursor -Object $Object -Direction EndKey -Unit $wdStory        
                    $Object.ActiveWindow.ActivePane.View.SeekView = $wdSeekMainDocument
                }Catch{
                    Write-log -Category Error -message "[$FunctionName] : Error bij het aanmaken van een Voettekst."
                    Write-log -Category Error -message "[$FunctionName] : Targetname   : $($_.CategoryInfo.targetname)"
                    Write-log -Category Error -message "[$FunctionName] : Fullname     : $($_.exception.gettype().Fullname)"
                    Write-log -Category Error -message "[$FunctionName] : Type fout    : $($_.CategoryInfo.category)"
                    Write-log -Category Error -message "[$FunctionName] : Position     : $($_.invocationinfo.positionmessage)"
                    Write-log -Category Error -message "[$FunctionName] : Errormessage : $($_.Exception.message)"   
                }
            }
    }
}