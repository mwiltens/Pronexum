Function ZoekBookMark{
	Param($BookMark)
	#Deze functie zoek een bookmark in het document
	<# Start variabelen#>
	#$Word In deze globale variabele staat de DDE-Link naar Word	
	#$BookMark In deze variabele staat de bookmark, die in Word moet worden opgezocht
	<# End variabelen#>
	$Word.Selection.GoTo(-1,0,0, $bookmark) | Out-Null
    $Word.Selection.Find.ClearFormatting | Out-null
}