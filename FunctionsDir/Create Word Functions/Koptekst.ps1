Function Koptekst{
	Param ($Tekst, $Heading)
	#Deze functie schrijft een koptekst
	<# Start variabelen#>
	#$Word In deze variabele staat de DDE-Link naar Word
	#$Tekst In deze variabele staat de tekst die in Word moet worden geprint
	#$Heading In deze variabele staat de style die de koptekst moet krijgen
	#$wdLine In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$wdExtend In deze variabele staat de waarde die Word nodig is om een bepaalde bewerking uit te voeren
	<# End variabelen#>
    TypeTekst $Tekst  1 0 0 1
	$WORD.Selection.MoveUp($wdLine, 1) | Out-Null
	$WORD.Selection.EndKey($wdLine, $wdExtend) | Out-Null
    Style $heading
}