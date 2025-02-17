<#
 .Synopsis
  Maakt een legenda aan

 .Description
  Aanmaken van een legenda
  
 .Parameter Object
  Is het Object, waarop de bewerking wordt toegepast


 .Example
   # Voegt de verboselogfile dat aangemaakt is op de D-schijf toe aan de Verbosefile.
   Maak-legenda
#>

Function Maak-Legenda1{
    Param ($Object)
    $FunctionName = "Maak-Legenda"
    Try{
        Maak_Kolommen_test -Object $Object -Rijen 1 -Kolommen 3 -Kolombreedtes @(1.05, 0.13, 4.92) -Borderlines True -TabelNr 1
    }catch{
        Write-log -message "[$FunctionName] : Unknown Error"
    
    }


    

}
Function Maak-Legenda{
    Param ($Object)
    $FunctionName = "Maak-Legenda"
    Try{
        $Selection = $Global:WORD.Selection
        #Maak_Kolommen -Object $Object -Rijen 1 -Kolommen 1 -Borderlines True
        #Move-Cursor -Object $Object -Direction Down -Unit $wdLine -Count 1
        typetekst -tekst "" -Object $WORD -Fontsize 11 -Fontname "Calibri"
        Maak_Kolommen -Object $Object -Rijen 1 -Kolommen 3 -Kolombreedtes @(1.05, 0.13, 4.92) -Borderlines False -TabelNr 1
        Move-Cursor -Object $Object -Direction Right -Unit $wdCell -Count 2
        typetekst -tekst "Historie" -Object $Object -Bold $True -Italic $True -Fontsize 16 -Fontname "Arial" -Enter $false
        Move-Cursor -Object $Word -Direction Down -Unit $wdLine -Count 1
        $Object.Selection.Endkey(6) | Out-Null
        $Object.Selection.TypeParagraph()
        #Move-Cursor -Object $Object -Direction Down -Unit $wdLine -Count 1

        Maak_Kolommen -Object $Object -Rijen 1 -Kolommen 3 -Kolombreedtes @(1.05, 0.13, 4.92) -Borderlines False -TabelNr 2
        Move-Cursor -Object $Object -Direction Up -Unit $wdLine -Count 1
        $Selection.Delete($wdCharacter, 1)
        $Range = $Selection.Range
        $Fontname = "Calibri"
        $Fontsize = 8
        typetekst -tekst "" -Object $WORD -Fontsize $Fontsize -Fontname $Fontname
        $Global:Doc.fields.add($Range, 21, "CREATEDATE  \@ ""d-MM-yyyy"" ", $true)
        typetekst -tekst "" -Object $WORD -Enter $False
        Move-Cursor -Object $Word -Direction Right -Unit $wdCharacter -Count 1
        #typetekst -tekst "" -Object $WORD -Fontsize $Fontsize -Fontname $Fontname
        typetekst -tekst "\" -Object $WORD -Enter $False -Fontsize $Fontsize -Fontname $Fontname
        Move-Cursor -Object $Word -Direction Right -Unit $wdCharacter -Count 1
        typetekst -tekst "" -Object $WORD -Fontsize $Fontsize -Fontname $Fontname
        $Range = $Selection.Range
        $Global:Doc.fields.add($Range, 61, "\* Upper", $true)| Out-Null
        #typetekst -tekst "" -Object $WORD -Enter $true

        Move-Cursor -Object $Object -Direction Down -Unit $wdLine -Count 1
        typetekst -tekst "Inhoud" -Object $Object -Bold $True -Italic $True -Fontsize 16 -Fontname "Calibri" -Enter $True
        Move-Cursor -Object $Object -Direction Down -Unit $wdline -Count 1
        Style -Object $Word -Style "Geen Afstand"

        Aanmaken_Inhoudsopgave -Object $Object 
        $Object.Selection.insertbreak($wdPageBreak)
    }catch{
        Write-log -Category Error -message "[$FunctionName] : Unknown error. "
        Write-log -Category Error -message "[$FunctionName] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$FunctionName] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$FunctionName] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$FunctionName] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$FunctionName] : Errormessage : $($_.Exception.message)"   
    
    }

<#  Call Maak_Kolommen(1)
    Call Regellager
    Call Enter
    Call Maak_Kolommen(2)

    Call ZoekBookmark("Pagina2")
    Rem Vul de Legenda
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Application.Run MacroName:="Project.NewMacros.Bold"
    Application.Run MacroName:="Project.NewMacros.Italic"
    Selection.Font.Size = 16
    Selection.Font.Name = "Arial"
    Selection.TypeText Text:="Historie"
    Selection.MoveDown Unit:=wdLine, Count:=2
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.Font.Size = 8
    Selection.Font.Name = "Arial"
    Application.Run MacroName:="Project.NewMacros.Italic"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "CREATEDATE  \@ ""d-MM-yyyy"" ", PreserveFormatting:=True
    Selection.TypeText Text:="/"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "USERINITIALS  \* Upper ", PreserveFormatting:=True
    Rem Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.Font.Size = 8
    Selection.Font.Name = "Arial"
    Selection.TypeText Text:="Initieel document"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.Font.Size = 16
    Selection.Font.Name = "Arial"
    Application.Run MacroName:="Project.NewMacros.Bold"
    Application.Run MacroName:="Project.NewMacros.Italic"
    Selection.TypeText Text:="Inhoud"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "TOC \o ""1-3"" \h \z \u ", PreserveFormatting:=True

End Sub
#>

    

}