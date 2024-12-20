Function ZetKolomWidth{
	Param($Kolom, $Width)
	# Start variabelen
	#$Kolom In deze variabele staat de kolom welke moet worden aangepast in de tabel
	#$Width In deze variabele staat de breedte die moet worden opgegeven van de kolom
	#$Word In deze globale variabele staat de DDE-Link naar Word	
	#$Width1 In deze variabele staat breedte in points voor de kolom-breedte
	#$wdAdjustNone In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$wdExtend In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$wdLine In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	<# End variabelen#>
	$word.Selection.MoveDown($wdLine, 1, $wdExtend)	  | out-null
	$Width1 = InchesToPoints $width
	$WORD.Selection.Tables.Item(1).Columns.item($kolom).setwidth($Width1, $wdAdjustNone)
}