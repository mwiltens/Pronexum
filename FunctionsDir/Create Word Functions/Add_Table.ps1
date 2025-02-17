<#
 .Synopsis
  Toevoegen van een tabel

 .Description
  Er wordt een tabel toegevoegd aan het Word document
  
 .Parameter Object
  Is het Object, waarop de bewerking wordt toegepast

 .Parameter Row
  Het aantal rijen die de tabel gaat bevatten

 .Parameter Column
  Het aantal kolommen die de tabel gaat bevatten

 .Parameter BackgroudColor
  De kleur van de achtergrond

 .Parameter 1

.Example
   # Maakt een tabel aan met 2 rijen en 3 kolommen. De achtergrond
   Add-Table -Object $WORD -row 2 -Column 3 -BackgroundColor 5


#>

Function Add-Table { 
    Param ($Object, [int]$Row = 2, [int]$Column = 5, $BackgroundColor) 
    #Param ($Object, [int]$Row = 2, [int]$Column = 5, [Validateset("3","5")]$BackgroundColor) 
    $Functionname = "Add-Table"
    Try{
	    $Object.selection.TypeParagraph() | Out-Null
	    #$global:paragraph = $WORD.Content.Paragraphs.Add() 	
        #$range = $paragraph.Range 
        $global:Table = $Object.activedocument.Tables.Add($Object.Selection.Range,$Row,$Column) 
	    $Table.AutoFormat($BackgroundColor)
        $Selection = $Object.Selection.Tables(1)
        If ($Selection.Style -ne "Tabelraster"){
            #$Selection.Style = "Tabelraster"
        }
        $Selection.ApplyStyleHeadingRows = $true
        $Selection.ApplyStyleLastRow = $False
        $Selection.ApplyStyleFirstColumn = $True
        $Selection.ApplyStyleLastColumn = $False
        $Selection.ApplyStyleRowBands = $True
        $Selection.ApplyStyleColumnBands = $False

        $Selection.Selection.Borders($wdBorderTop).LineStyle = 1
        $Selection.Selection.Borders($wdBorderTop).LineWidth = 1
        $Selection.Selection.Borders($wdBorderTop).Color = 0
    }Catch{
        Write-log -Category Error -message "[$FunctionName] : Unknown error. "
        Write-log -Category Error -message "[$FunctionName] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$FunctionName] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$FunctionName] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$FunctionName] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$FunctionName] : Errormessage : $($_.Exception.message)"   
    }
        <#
        $Selection.Selection.Borders($wdBorderBottom).LineStyle = 1
        $Selection.Selection.Borders($wdBorderBottom).LineWidth = 1
        $Selection.Selection.Borders($wdBorderBottom).Color = 0

        $Selection.Selection.Borders($wdBorderRight).LineStyle = 1
        $Selection.Selection.Borders($wdBorderRight).LineWidth = 1
        $Selection.Selection.Borders($wdBorderRight).Color = 0

        $Selection.Selection.Borders($wdBorderHorizontal).LineStyle = 1
        $Selection.Selection.Borders($wdBorderHorizontal).LineWidth = 1
        $Selection.Selection.Borders($wdBorderHorizontal).Color = 0
        $Selection.Selection.Borders($wdBorderVertical).LineStyle = 1
        $Selection.Selection.Borders($wdBorderVertical).LineWidth = 1
        $Selection.Selection.Borders($wdBorderVertical).Color = 0



            Selection.EscapeKey
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
    With Selection.Borders(wdBorderTop)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderLeft)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderBottom)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderRight)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderHorizontal)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderVertical)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With





        #$Object.Selection.EndKey($wdstory) | Out-Null
        #$Selection.Selection.Tables(1)
        #With $Selection.Selection.Tables(1)
        #    If .Style <> "Tabelraster" Then
        #        .Style = "Tabelraster"
        #    End If
        #    .ApplyStyleHeadingRows = True
        #    .ApplyStyleLastRow = False
        #    .ApplyStyleFirstColumn = True
        ##    .ApplyStyleLastColumn = False
        #    .ApplyStyleRowBands = True
        #    .ApplyStyleColumnBands = False
        #End With
    }catch{
        Write-log -Category Error -message "[$Functionname] : Fout bij het toevoegen van een tabel. "
        Write-log -Category Error -message "[$Functionname] : Targetname   : $Global:Error.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $Global:Error.exception.gettype.Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $Global:Error.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $Global:Error.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $Global:Error.Exception.message)"
    }
}

<#
$word = New-Object -comobject word.application
$word.Visible = $false
$doc = $word.Documents.Add()
$Selection = $word.Selection
$Selection.Style="Title"
$Selection.Font.Bold = 1
$Selection.ParagraphFormat.Alignment = 2
$Selection.TypeText("Expected Billing")
$Selection.TypeParagraph()
$Selection.Style="No Spacing"
$Selection.Font.Bold = 0
$Selection.TypeParagraph()
$doc.SaveAs([ref]$savepath) 
$doc.Close() 
$word.quit()
#>

<#
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=3, NumColumns:= _
        3, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    With Selection.Tables(1)
        If .Style <> "Tabelraster" Then
            .Style = "Tabelraster"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
    End With
    Selection.MoveRight Unit:=wdWord, Count:=3, Extend:=wdExtend
    Selection.Shading.Texture = wdTextureNone
    Selection.Shading.ForegroundPatternColor = wdColorAutomatic
    Selection.Shading.BackgroundPatternColor = 6299648
    Selection.TypeText Text:="Dit is een test"
    Selection.Shading.Texture = wdTextureNone
    Selection.Shading.ForegroundPatternColor = wdColorAutomatic
    Selection.Shading.BackgroundPatternColor = -738132122
#>
}