Function Maak_Legenda{
	#Deze functie maakt binnnen Word een legenda aan
	<# Start variabelen#>
	#$Parameter1 In deze variabele staat parameter1_beschrijving
	#$Parameter2 In deze variabele staat parameter2_beschrijving
	#$Word In deze variabele staat de DDE-Link naar Word
	#$wdCharacter In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$wdFieldEmpty In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$wdLine In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$wdPageBreak In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	<# End variabelen#>
    Maak_Kolommen 3
    $word.Selection.MoveRight($wdCharacter, 2) | Out-Null
    $word.Selection.Font.Size = 16
    $word.Selection.Font.Name = "Arial"
    TypeTekst "Historie" 1 1 "" 0
	$WORD.activedocument.Tables.Add | Out-Null
	$word.Selection.MoveRight($wdCharacter, 2) | Out-Null
    $word.Selection.MoveDown($wdLine, 1)	 | Out-Null    
    $word.Selection.Font.Size = 8
    $word.Selection.Font.Name = "Arial"
	opmaak "" 1 ""	
    $Word.Selection.Fields.Add($word.Selection.Range, $wdFieldEmpty, "CREATEDATE  \@ ""d-MM-yyyy"" ", $True) | Out-Null
    $word.Selection.TypeText("/")
    $word.Selection.Fields.Add($word.Selection.Range, $wdFieldEmpty, "USERINITIALS  \* Upper ", $True) | Out-Null
    $word.Selection.MoveRight($wdCharacter, 2) | Out-Null
    $word.Selection.Font.Size = 8
    $word.Selection.Font.Name = "Arial"
    $word.Selection.TypeText("Initieel document")
    $word.Selection.MoveDown($wdLine, 1) | Out-Null
    $word.Selection.MoveLeft($wdCharacter, 2) | Out-Null
	$word.Selection.MoveDown($wdLine, 2) | Out-Null
    $word.Selection.Font.Size = 16
    $word.Selection.Font.Name = "Arial"
	opmaak "" 0 ""	
    typetekst "Inhoud" 1 1 "" 1
    $word.Selection.MoveDown($wdLine, 1) | Out-Null
    $word.Selection.Fields.Add($word.Selection.Range, $wdFieldEmpty, "TOC \o ""1-5"" \h \z \u ", $True) | Out-Null
	$word.Selection.TypeParagraph | Out-Null
	$word.Selection.TypeParagraph | Out-Null
    $word.Selection.InsertBreak($wdPageBreak) | Out-Null

}