Function Add-Table { 
	param ( [int]$row = 2, [int]$col = 5,$Object) 
	$Word.selection.TypeParagraph() | Out-Null
	$global:paragraph = $WORD.Content.Paragraphs.Add() 	
    $range = $paragraph.Range 
    $global:table = $WORD.activedocument.Tables.Add($word.Selection.Range,$row,$col) 
	$table.AutoFormat(3)
}