$document = "C:\Data\scripts\Sjabloon.dotm" # in deze variable staat de naam van het document wat door Word wordt gebruikt
$msoLanguageIDDutch = 1043 # Word variabele
$msoLanguageIDEnglishUS = 1033 # Word variabele
$wdAdjustNone = 0 # Word variabele
$wdAutoFitFixed = 0 # Word variabele
$wdAlignParagraphCenter = 1 # Word variabele
$wdBorderLeft = -2 # Word variabele
$wdBorderRight = -4 # Word variabele
$wdBorderTop = -1 # Word variabele
$wdBorderBottom = -3 # Word variabele
$wdBorderHorizontal = -5 # Word variabele
$wdBorderVertical = -6 # Word variabele
$wdBorderDiagonalDown = -7 # Word variabele
$wdBorderDiagonalUp = -8 # Word variabele
$wdBulletGallery = 1 # Word variabele
$wdCell = 12 # Word variabele
$wdCharacter = 1 # Word variabele
$wdColumn = 9 # Word variabele
$wdExtend = 1 # Word variabele
$wdFieldEmpty = -1 # Word variabele
$wdFormatXMLDocument = 12 # Word variabele
$wdGoToBookmark = -1 # Word variabele
$wdLine=5 # Word variabele
$wdLineStyleNone = 0 # Word variabele
$wdNumberParagraph = 1 # Word variabele
$wdPageBreak = 7 # Word variabele
$wdRow = 10 # Word variabele
$wdStory = 6 # Word variabele
$wdTable = 15 # Word variabele
$wdToggle=9999998  # Word variabele
$wdUnderlineNone =0 # Word variabele
$wdUnderlineSingle =1 # Word variabele
$Wdword = 2 # Word variabele
$wdWord9TableBehavior = 1 # Word variabele

$wdPropertyTitle = 1    # Word variabele
$wdSeekMainDocument = 0 # Word variabele
$wdSeekCurrentPageHeader = 9 # Word variabele
$wdSeekCurrentPageFooter =10 # Word variabele

$Quote = [Char]34 #Variabele om een Quote weer te geven
$Tab = [Char]9 #Variabele om een tab weer te geven
Function Add-Table { 
	param ( [int] $row = 2, [int] $col = 5) 
	#Deze functie voegt een tabel toe in Word
	<# Start variabelen#>
	#$Row In deze variabele staat het aantal rijen, die een tabel moet bevatten
	#$Ccol In deze variabele staat het aantal kolommen, die een tabel moet bevatten
	#$Word In deze variabele staat de DDE-Link naar Word
	#$Paragraph In deze variabele staat het commando om een regel verder te gaan
	#$Range In deze variabele staat het bereid van word
	#$Table In deze variabele staat de waarde om een tabel te maken
	<# End variabelen#>
	$Word.selection.TypeParagraph() | Out-Null
	$global:paragraph = $WORD.Content.Paragraphs.Add() 	
    $range = $paragraph.Range 
    $global:table = $WORD.activedocument.Tables.Add($word.Selection.Range,$row,$col) 
	$table.AutoFormat(3)
}
Function CheckStyle{
	Param ($Stijl)
	#Deze functie bepaald de stijl voor de tekst opmaak aan de hand van de taal die geinstalleerd is op de computer
	<# Start variabelen#>
	#$Word In deze variabele staat de DDE-Link naar Word
	#$Stijl In deze variabele staat de stijl die gecontroleerd moet worden
	#CheckStyle In deze variabele staat welke Style er wordt gebruikt voor de opmaak van het document
	#$msLanguageIDDutch In deze variabele staat de code voor de Nederlandse taal
	#$msoLanguageIDEnglishUS In deze variabele staat de code voor de Engelse taal
	<# End variabelen#>
	$CheckStyle = $Stijl 
    Switch ($word.Application.Language){
		$msoLanguageIDDutch {
            #MsgBox "The user interface language is Dutch."
            #Rem Dim StrHeading1, StrHeading2, StrHeading3, StrHeading4, StrHeading5, StrHeading6
            Switch ($Stijl){
                "Kop 1"{
                    $CheckStyle = "Kop 1"
				}
                "Kop 2"{
                    $CheckStyle = "Kop 2"
				}
                "Kop 3"{
                    $CheckStyle = "Kop 3"
				}
                "Kop 4"{
                    $CheckStyle = "Kop 4"
				}
                "Kop 5"{
                    $CheckStyle = "Kop 5"
				}
                "Kop 6"{
                    $CheckStyle = "Kop 6"
				}
            }			
		}
		$msoLanguageIDEnglishUS {
            Switch ($Stijl){
                "Kop 1"{
                    $CheckStyle = "Heading 1"
				}
                "Kop 2"{
                    $CheckStyle = "Heading 2"
				}
                "Kop 3"{
                    $CheckStyle = "Heading 3"
				}
                "Kop 4"{
                    $CheckStyle = "Heading 4"
				}
                "Kop 5"{
                    $CheckStyle = "Heading 5"
				}
                "Kop 6"{
                    $CheckStyle = "Heading 6"
				}
            }
		}
	}
	Return $CheckStyle
}
Function InchesToPoints{
	param($inches)
	#Deze functie berekend het aantal points dat in een inch gaat.
	<# Start variabelen#>
	#$Inches In deze variabele staat het aantal inches dat moet worden omgezet in het aantal points
	#$Points In deze variabele staat het aantal points dat berekend wordt in deze functie
	<# End variabelen#>
	$Points = $inches * 72
	return $Points
}
Function CentimetersToPoints{
	Param ($IntCentimters)
	#Deze functie berekend het aantal points die de opgegeven aantal centimeters oplevert
	<# Start variabelen#>
	#$IntCentimters In deze variabele staat het aantal centimeters, wat omgezet moet worden naar het aantal points
	#$Points In deze variabele staat het aantal points, dat berekend wordt in deze functie
	<# End variabelen#>
	$Points = $IntCentimters * 28.35
	return $Points
}
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
Function Koptekst{
	Param ($Tekst, $Heading)
	#Deze functie schrijft een koptekst
	<# Start variabelen#>
	#$Word In deze variabele staat de DDE-Link naar Word
	#$Tekst In deze variabele staat de tekst die in Word moet worden geprint
	#$Heading In deze variabele staat de style die de koptekst moet krijgen
	#$wdLine In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$wdExtend In deze variabele staat de waarde die Word nodig is om een bepaalde bewerking uit te voeren
	<# End variabelen#>
    TypeTekst $Tekst  1 0 0 1
	$WORD.Selection.MoveUp($wdLine, 1) | Out-Null
	$WORD.Selection.EndKey($wdLine, $wdExtend) | Out-Null
    Style $heading
}
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
Function Maak_Kolommen1{
	Param ($Rijen, $Kolommen,$Shading = $false, $Center = $false, $Stylenone = $False)
	#Deze functie maakt een tabel in Word 
	<# Start variabelen#>
	#$Rijen In deze variabele staat het aantal rijen, die de tabel moet bevatten
	#$Kolommen In deze variabele staat het aantal kolommen, die de tabel moet bevatten
	#$Shading$ In deze variabele staat of er shading moet worden toegepast
	#$Center In deze variabele staat of de tekst gecentreerd moet worden
	#$Stylenone In deze variabele staat of de tabel lijnen moet krijgen
	#$Word In deze globale variabele staat de DDE-Link naar Word	
	#$wdWord9TableBehavior In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$wdAutoFitFixed In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$Table In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$wdAlignParagraphCenter In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$wdExtend In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$wdBorderTop In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$wdBorderLeft In deze variabele staat Var1_beschrijving
	#$wdBorderBottom In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$wdBorderRight In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$wdLine In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$wdBorderHorizontal In deze variabele staat Var1_beschrijving
	#$wdBorderVertical In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$wdBorderDiagonalDown In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$wdBorderDiagonalUp In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	#$wdBorderRight In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren
	<# End variabelen#>
	$Word.ActiveDocument.Tables.Add($word.Selection.Range, $Rijen, $Kolommen, $wdWord9TableBehavior, $wdAutoFitFixed) | Out-Null
	$table = $Word.ActiveDocument.Tables.item(1)	| Out-Null
	If ($Shading -eq $true){
		$word.Selection.MoveRight($wdWord, $kolommen, $wdExtend)  | out-null
		$WORD.Selection.Cells.Shading.BackgroundPatternColor = -587137025
		If ($Center -eq $true){
			$WORD.Selection.ParagraphFormat.Alignment = $wdAlignParagraphCenter
		}
		If ($stylenone -eq $true){
			$aantalrijen = $Rijen - 1
				$WORD.Selection.Movedown($wdline,1,$wdExtend)  | out-null
				$WORD.Selection.MoveRight($wdline,2,$wdExtend)  | out-null
    			$WORD.Selection.Borders.Item($wdBorderTop).LineStyle = $wdLineStyleNone
    			$WORD.Selection.Borders.Item($wdBorderLeft).LineStyle = $wdLineStyleNone
    			$WORD.Selection.Borders.Item($wdBorderBottom).LineStyle = $wdLineStyleNone
    			$WORD.Selection.Borders.Item($wdBorderRight).LineStyle = $wdLineStyleNone
    			$WORD.Selection.Borders.Item($wdBorderHorizontal).LineStyle = $wdLineStyleNone
    			$WORD.Selection.Borders.Item($wdBorderVertical).LineStyle = $wdLineStyleNone
    			$WORD.Selection.Borders.Item($wdBorderDiagonalDown).LineStyle = $wdLineStyleNone
    			$WORD.Selection.Borders.Item($wdBorderDiagonalUp).LineStyle = $wdLineStyleNone
				$WORD.Selection.Borders.Item($wdBorderRight).LineStyle = $wdLineStyleNone
		
		}
	}
}
Function OpenWord{
	param ($document)
	#Deze functie maakt een DDE-Link aan met MS-Word. 
	<# Start variabelen#>
	#$Word In deze globale variabele staat de DDE-Link naar Word	
	#$Document In deze variabele staat de naam van een Word-document wat geopend moet worden
	#$DOC In deze Globale variabele staat de naam van het Word-document, wat door Word is gekoppeld
	<# End variabelen#>
	$global:WORD = New-Object -comobject Word.Application
	$WORD.Visible = $True
	If ($document -eq "") {
		$global:DOC = $WORD.Documents.add()
	}else {
		($Word.documents.Open($document)) | out-null
	}
}
Function Opmaak{
	Param($Bold, $Italic, $Underline)
	#Deze functie bepaald de opmaak van een stuk tekst
	<# Start variabelen#>
	#$Bold Als deze variabele de waarde 1 heeft, wordt de tekst als Vet opgemaakt
	#$Italic$ Als deze variabele de waarde 1 heeft, wordt de tekst als Italic opgemaakt
	#$Underline Als deze variabele de Waarde 1 heeft, wordt de tekst met Underline opgemaakt
	#$Word In deze globale variabele staat de DDE-Link naar Word	
	#$wdToggle In deze variabele staat de code, die Wordt nodig heeft om Italic uit te voeren
	#$wdUnderline In deze variabele staat de code, die Wordt nodig heeft om Underline uit te voeren
	#$wdUnderlineSingle In deze variabele staat de code, die Wordt nodig heeft om de Underline op te maken
	#$wdUnderlineNone In deze variabele staat de code, die Wordt nodig heeft om de Underline op te maken
	<# End variabelen#>
    If ($Bold -ne ""){
        $WORD.Selection.Font.Bold = $wdToggle
    }
    If ($Italic  -ne ""){
        $WORD.Selection.Font.Italic = $wdToggle
    }
    If ($Underline -ne ""){
        If ($WORD.Selection.Font.Underline -eq $wdUnderlineNone){
            $WORD.Selection.Font.Underline = $wdUnderlineSingle
        }Else{
            $WORD.Selection.Font.Underline = $wdUnderlineNone
        }
    }
}
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
Function OpsommingtekensUit{
	#Deze functie zet de opsommingstekens uit
	<# Start variabelen#>
	#$Word In deze globale variabele staat de DDE-Link naar Word	
	#$wdBulletGallery In deze variabele staat de waarde die Word nodig is om een bewerking uit te voeren	
	<# End variabelen#>
	$WORD.selection.Range.ListFormat.RemoveNumbers($wdNumberParagraph)
	$WORD.Selection.ParagraphFormat.LeftIndent = inchestopoints 0.5	
}
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
Function TypeTekst{
	Param($Tekst,$Bold, $Italic, $Underline, $Enter)
	#Deze functie schrijft tekst in een word document
	<# Start variabelen#>
	#$Tekst In deze variabele staat de tekst die geprint moet worden in het Word document
	#$Bold Als deze variabele de waarde 1 heeft, wordt de tekst als Vet opgemaakt
	#$Italic$ Als deze variabele de waarde 1 heeft, wordt de tekst als Italic opgemaakt
	#$Underline Als deze variabele de Waarde 1 heeft, wordt de tekst met Underline opgemaakt
	#$Enter Als deze variabele de waarde 1 heeft, wordt er een Enter toegevoegd aan de teskt
	#$Word In deze globale variabele staat de DDE-Link naar Word	
	#$wdToggle In deze variabele staat de code, die Wordt nodig heeft om Italic uit te voeren
	#$wdUnderline In deze variabele staat de code, die Wordt nodig heeft om Underline uit te voeren
	#$wdUnderlineSingle In deze variabele staat de code, die Wordt nodig heeft om de Underline op te maken
	#$wdUnderlineNone In deze variabele staat de code, die Wordt nodig heeft om de Underline op te maken
	<# End variabelen#>
    If ($Bold -eq 1){
        $WORD.Selection.Font.Bold = $wdToggle
    }
    If ($Italic  -eq 1){
        $WORD.Selection.Font.Italic = $wdToggle
    }
    If ($Underline -eq 1){
        If ($WORD.Selection.Font.Underline -eq $wdUnderlineNone){
            $WORD.Selection.Font.Underline = $wdUnderlineSingle
        }Else{
            $WORD.Selection.Font.Underline = $wdUnderlineNone
        }
    }
	$Word.selection.typetext("$Tekst")
    $WORD.Selection.Font.Bold = $False
    $WORD.Selection.Font.Italic = $False
    $WORD.Selection.Font.Underline = $wdUnderlineNone
	If ($Enter -eq 1){		
		$WORD.Selection.TypeParagraph()
	}
}
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
### Aanpassen document properties
	$binding = "System.Reflection.BindingFlags" -as [type]
	$builtinProperties = $WORD.ActiveDocument.BuiltInDocumentProperties
	[Array]$getArgs = "Title"
	$builtinPropertiesType = $builtinProperties.GetType()
	$builtinProperty = $builtinPropertiesType.InvokeMember( `
	  "Item", $binding::GetProperty, `
	  $null, $builtinProperties, $getArgs)
	$builtinPropertyType = $builtinProperty.GetType()
	[Array]$setArgs = $ScriptbaseName
	$builtinPropertyType.InvokeMember( `
	  "Value", $binding::SetProperty, `
	  $null, $builtinProperty, $setArgs)
### Fontsize aanpassen
$word.Selection.Font.Size = 12
$word.selection.Font.name = $fontname
####Aanmaken kolommen

	Maak_kolommen1 2 $Kolommen $False $False $False
#### Table End

		$word.Selection.EndKey($wdStory)  | out-null
		$WORD.Selection.TypeParagraph()  | out-null
#### Update word document
	$WORD.ActiveWindow.ActivePane.View.SeekView = $wdSeekCurrentPageFooter
	$word.Selection.WholeStory()
	$word.Selection.Fields.Update()    | out-null
	$WORD.ActiveWindow.ActivePane.View.SeekView = $wdSeekMainDocument
	$word.Selection.WholeStory()
	$word.Selection.Fields.Update()| out-null
	$WORD.activedocument.save()  | out-null
	$WORD.Selection.Homekey($wdStory) | out-null
