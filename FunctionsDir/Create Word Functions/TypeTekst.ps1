<#
 .Synopsis
  Voegt tekst toe aan het MS-Word document

 .Description
  Voegt tekst toe aan het MS-Word document 
  
 .Parameter Object
  Geef het Object op, welke gebruikt wordt om een tekst te bewerken

 .Parameter Tekst
  Bevat de tekst, die in MS-Word wordt toegevoegd

 .Parameter Enter
  Voegt een newline commando toe aan de tekst

 .Parameter Bold
  De tekst wordt 'vet' afgedrukt

 .Parameter Italic
  De tekst wordt 'Italic' afgedrukt

 .Parameter Underline
  De tekst wordt 'Onderstreept' afgedrukt

 .Parameter Fontsize
  De tekst krijgt de opgegeven lettergrootte mee

 .Parameter Fontname
  De tekst krijgt de opgegeven Font mee

 .Example
   # voegt de tekst "Hello World" toe aan het opgegeven object en vervolgens wordt er een Enter toegevoegd, zodat er naar de volgende regel wordt gegaan.
   Typetekst -tekst "Hello World" -Object $WORD -Enter $true

   # voegt de tekst "Hello World" toe aan het opgegeven object en vervolgens wordt er een Enter toegevoegd, zodat er naar de volgende regel wordt gegaan. De tekst wordt 'Vet' afgedrukt
   Typetekst -tekst "Hello World" -Object $WORD -bold $true -Enter $True

   # voegt de tekst "Hello World" toe aan het opgegeven object en vervolgens wordt er een Enter toegevoegd, zodat er naar de volgende regel wordt gegaan. De tekst wordt 'Italic' afgedrukt 
   Typetekst -tekst "Hello World" -Object $WORD -Italic $True  -Enter $True

   # voegt de tekst "Hello World" toe aan het opgegeven object en vervolgens wordt er een Enter toegevoegd, zodat er naar de volgende regel wordt gegaan. De tekst wordt 'Onderstreept' afgedrukt
   Typetekst -tekst "Hello World" -Object $WORD -Underline $True -Enter $True

   # voegt de tekst "Hello World" toe aan het opgegeven object en vervolgens wordt er een Enter toegevoegd, zodat er naar de volgende regel wordt gegaan. De tekst wordt 'Vet', Italic en onderstreept afgedrukt
   Typetekst -tekst "Hello World" -Object $WORD -bold $true -Italic $true -Underline $true -Enter $true

   # voegt de tekst "Hello World" toe aan het opgegeven object en vervolgens wordt er een Enter toegevoegd, zodat er naar de volgende regel wordt gegaan. De tekstgrootte is 15
   Typetekst -tekst "Hello World" -Object $WORD -Fontsize 15 -Enter $true

   # voegt de tekst "Hello World" toe aan het opgegeven object en vervolgens wordt er een Enter toegevoegd, zodat er naar de volgende regel wordt gegaan. Het font van de tekst is "Times New Roman"
   Typetekst -tekst "Hello World" -Object $WORD -Fontname "Times New Roman"  -Enter $true. 

#>
Function TypeTekst{
    param($Object,$Bold=$False, $Italic=$False, $Underline=$False, $Fontsize, $Fontname,  $Tekst,  $Enter=$False)
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
        Write-log -Category Error -message "[$Functionname] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $($_.Exception.message)"   
    }
}