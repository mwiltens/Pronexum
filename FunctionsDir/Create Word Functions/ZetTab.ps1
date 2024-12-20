Function ZetTab{
	Param ($TabPos,$clear = $False)
	#Deze functie zet de tabpositie
	<# Start variabelen#>
	#$Word In deze globale variabele staat de DDE-Link naar Word	
	#$TabPos In deze variabele staat de positie in, waar de tabpositie komt te staan. Deze waarde wordt in Inches opgegeven
	#$Clear In deze variabele staat of de tabposities moeten worden verwijderd
	#$Pos In deze variabele staat positie in Points.
	<# End variabelen#>
	$Pos = InchesToPoints $TabPos
	If ($clear -eq $true){		
		($word.Selection.ParagraphFormat.TabStops.item($Pos).clear())  | Out-Null
		$WORD.ActiveDocument.DefaultTabStop= Inchestopoints 0.49
	}else{
    	($word.Selection.ParagraphFormat.TabStops.Add($Pos ,0,0))  | Out-Null
	}
}