Function Opsommingtekens{
	#Deze functie zet de opsommingstekens aan
	<# Start variabelen#>
	#$Word In deze globale variabele staat de DDE-Link naar Word	
	#$wdBulletGallery In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren	
	<# End variabelen#>
	$Word.ListGalleries.Item($wdBulletGallery).ListTemplates.item(1).ListLevels.item(1).NumberPosition = InchesToPoints 0
	$Word.ListGalleries.Item($wdBulletGallery).ListTemplates.item(1).ListLevels.item(1).TextPosition = InchesToPoints 0.5
	$Word.Selection.Range.ListFormat.ApplyListTemplateWithLevel($word.ListGalleries.item($wdBulletGallery).ListTemplates.item(1))
}