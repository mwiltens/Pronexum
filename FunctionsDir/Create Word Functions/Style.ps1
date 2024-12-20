Function Style{
	Param ($Stijl)
	#Deze functie bepaald de opmaak van een stuk tekst
	<# Start variabelen#>
	#$Stijl In deze variabele staat Hoe de tekst moet worden opgemaakt
	#$Word In deze globale variabele staat de DDE-Link naar Word	
	#$wdLine In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren	
	<# End variabelen#>
    $Stijl = CheckStyle $Stijl
	$WORD.Selection.Style = $Stijl	
    $WORD.Selection.MoveDown($wdLine, 1) | Out-Null
}