Function ZetKantlijn{
	Param ($Kantlijn)
	#Deze functie stelt de kantlijn in
	<# Start variabelen#>
	#$kantlijn In deze variabele staat de positie waarde kantlijn moet komen te staan
	#$Word In deze globale variabele staat de DDE-Link naar Word	
	#$Pos In deze variabele staat de positie in points
	<# End variabelen#>
	$Pos = InchesToPoints $Kantlijn		
	$Word.Selection.Paragraphs.LeftIndent = InchesToPoints $Kantlijn		
    $WORD.Selection.Paragraphs.FirstLineIndent = -1 * (InchesToPoints $Kantlijn	)
}