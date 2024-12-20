Function Maak_Kolommen{
	Param ($Rijen, $Kolommen)
	#Deze functie maakt in Word een tabel aan
	<# Start variabelen#>
	#$Word In deze variabele staat de DDE-Link naar Word
	#$Rijen In deze variabele staat het aantal rijen die de tabel moet bevatten
	#$Kolommen In deze variabele staat het aantal kolommen die de tabel moet bevatten
	#$wdWord9TableBehavior In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$wdAutoFitFixed In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$Table In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	<# End variabelen#>
	$Word.ActiveDocument.Tables.Add($word.Selection.Range, $Rijen, 3, $wdWord9TableBehavior, $wdAutoFitFixed) | Out-Null
	$table = $Word.ActiveDocument.Tables.item(1)	
	$table.borders.enable = $true	
   	$table.Borders.InsideLineStyle = 0
   	$table.Borders.OutsideLineStyle = 0	
	$Kolom = $table.columns.item(1)
	$Kolom.PreferredWidth = InchesToPoints 1.05
	$Kolom = $table.columns.item(2)
	$Kolom.PreferredWidth = InchesToPoints 0.13
	$Kolom = $table.columns.item(3)
	$Kolom.PreferredWidth = InchesToPoints 4.92
}