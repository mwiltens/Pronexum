Function LinesToPoints{
	Param ($IntLines)
	#Deze functie berekend het aantal points, die bij het opgegeven aantal lines hoort
	<# Start variabelen#>
	#$IntLines In deze variabele staat het aantal lines, wat omgezet moet worden naar het aantal points
	#$Parameter2 In deze variabele staat parameter2_beschrijving
	#$Points In deze variabele staat het aantal points, dat berekend wordt in deze functie
	<# End variabelen#>
	$Points = $IntLines * 12
	return $Points
}