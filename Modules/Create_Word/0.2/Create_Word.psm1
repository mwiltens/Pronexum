Function Add-Table { 
	param ( [int]$row = 2, [int]$col = 5,$Object) 
	$Word.selection.TypeParagraph() | Out-Null
	$global:paragraph = $WORD.Content.Paragraphs.Add() 	
    $range = $paragraph.Range 
    $global:table = $WORD.activedocument.Tables.Add($word.Selection.Range,$row,$col) 
	$table.AutoFormat(3)
}

Function CentimetersToPoints{
	Param ($IntCentimeters)
	$Points = $IntCentimeters * 28.35
	return $Points
}

Function CheckStyle{
	Param ($Stijl,$Object)
	$CheckStyle = $Stijl 
    Switch ($Object.Application.Language){
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

Function InchesToPoints{
	param($Inches)
	$Points = $Inches * 72
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

Function LinesToPoints{
	Param ($IntLines)
	$Points = $IntLines * 12
	return $Points
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


Function OpenWord{
    param ($Document)
    $Functionname = "OpenWord"
    Try{
        Write-host "[$Functionname] : Execute : `$GLOBAL:WORD = New-Object -comobject Word.Application"
        $GLOBAL:WORD = New-Object -comobject Word.Application
        Write_log -Category Verbose -Message "[$Functionname] : Execute : `$WORD.Visible = $True"
        $WORD.Visible = $True
        If ($Document -eq ""){
            Write_log -Message "[$Functionname] : Execute : `$GLOBAL:DOC = $WORD.Documents.add()"
            $GLOBAL:DOC = $WORD.Documents.add()
        }else {
            Write_log -Message "[$Functionname] : Execute : `($Word.Documents.Open($Document)) | out-null"
            ($Word.Documents.Open($Document)) | out-null
        }

    }catch{
        Write-log -Category Error -message "[Functionname] : Unknown error. "
        #Write-log -Category Error -message "[Functionname] : Targetname   : $(_.CategoryInfo.targetname)"
        #Write-log -Category Error -message "[Functionname] : Fullname     : $(_.exception.gettype.Fullname)"
        #Write-log -Category Error -message "[Functionname] : Type fout    : $(_.CategoryInfo.category)"
        #Write-log -Category Error -message "[Functionname] : Position     : $(_.invocationinfo.positionmessage)"
        #Write-log -Category Error -message "[Functionname] : Errormessage : $(_.Exception.message)"
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
    param($Object,$bold=$False, $Italic=$False, $Underline=$False, $Fontsize, $Fontname,  $Tekst,  $Enter=$False)
    $FunctionName = "TypeTekst"
    Write-log Verbose -Message "`$Tekst = `'$Tekst`'"
    Write-log Verbose -Message "`$bold = `'$bold`'"
    Write-log Verbose -Message "`$Italic = `'$Italic`'"
    Write-log Verbose -Message "`$Underline = `'$Underline`'"
    Try{
        If ($Bold){
            $Object.Selection.Font.Bold = $wdToggle
        }
        If ($Italic){
            $Object.Selection.Font.Italic = $WDToggle
        }
        If ($Underline){
            If($Object.Selection.Font.Underline -eq $WdUnderlineNone){
                $Object.Selection.Font.Underline = $WdUnderlineSingle
            }else{
                $Object.Selection.Font.Underline = $WdUnderlineNone
            }
        }
        If ($FontName){
            $Object.Selection.Font.Name = $FontName
        }
        If ($FontSize){
            $Object.Selection.Font.Size = $FontSize
        }
            $Object.Selection.TypeText("$Tekst")
            $Object.Selection.Font.Italic = $False
            $Object.Selection.Font.Bold = $False
                $Object.Selection.Font.Underline = $WdUnderlineNone
        if ($Enter){
            $Object.Selection.TypeParagraph()
        }
    }Catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "
        #Write-log -Category Error -message "[$Functionname] : Targetname   : $(_.CategoryInfo.targetname)"
        #Write-log -Category Error -message "[$Functionname] : Fullname     : $(_.exception.gettype.Fullname)"
        #Write-log -Category Error -message "[$Functionname] : Type fout    : $(_.CategoryInfo.category)"
        #Write-log -Category Error -message "[$Functionname] : Position     : $(_.invocationinfo.positionmessage)"
        #Write-log -Category Error -message "[$Functionname] : Errormessage : $(_.Exception.message)"
    }
}

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
$wdSeekCurrentPageFooter = 10 # Word variabele
$Quote = [Char]34 #Variabele om een Quote weer te geven
$Tab = [Char]9 #Variabele om een tab weer te geven

Function Write-log{
    Param ([Validateset("Verbose","Debug","Error")]$Category, $Message)
    Write-host $Message
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

