Function Opmaak{
	Param($Bold, $Italic, $Underline)
	#Deze functie bepaald de opmaak van een stuk tekst
	<# Start variabelen#>
	#$Bold Als deze variabele de waarde 1 heeft, wordt de tekst als Vet opgemaakt
	#$Italic$ Als deze variabele de waarde 1 heeft, wordt de tekst als Italic opgemaakt
	#$Underline Als deze variabele de Waarde 1 heeft, wordt de tekst met Underline opgemaakt
	#$Word In deze globale variabele staat de DDE-Link naar Word	
	#$wdToggle In deze variabele staat de code, die Wordt nodig heeft om Italic uit te voeren
	#$wdUnderline In deze variabele staat de code, die Wordt nodig heeft om Underline uit te voeren
	#$wdUnderlineSingle In deze variabele staat de code, die Wordt nodig heeft om de Underline op te maken
	#$wdUnderlineNone In deze variabele staat de code, die Wordt nodig heeft om de Underline op te maken
	<# End variabelen#>
    If ($Bold -ne ""){
        $WORD.Selection.Font.Bold = $wdToggle
    }
    If ($Italic  -ne ""){
        $WORD.Selection.Font.Italic = $wdToggle
    }
    If ($Underline -ne ""){
        If ($WORD.Selection.Font.Underline -eq $wdUnderlineNone){
            $WORD.Selection.Font.Underline = $wdUnderlineSingle
        }Else{
            $WORD.Selection.Font.Underline = $wdUnderlineNone
        }
    }
}
