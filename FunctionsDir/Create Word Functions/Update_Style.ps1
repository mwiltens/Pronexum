<#
 .Synopsis
  Korte omschrijving

 .Description
  Omschrijving.
  
 .Parameter 1


 .Parameter 1


 .Parameter 1


 .Example
   # Voegt de verboselogfile dat aangemaakt is op de D-schijf toe aan de Verbosefile.
   Read_RemoteLogfile -RemoteLogfile "\\$([Servernaam]\[Schijf]\$([scriptname])_Verbose.txt" -Logfile $Global:VerboseFile

#>
Function Update_Style{
    Param ($Object,$Style,$Fontname,$Fontsize,[ValidateSet($True,$False)]$Bold = $False,[ValidateSet($True,$False)]$Italic = $False,[ValidateSet($True,$False)]$Underline = $False)
    $FunctionName = "Update_Style"
    Try{
    $Object.ActiveDocument.Styles($Style).Font.Name = $FOntname
    $Object.ActiveDocument.Styles($Style).Font.Size = $Fontsize
    If ($Bold){
        $Object.ActiveDocument.Styles($Style).Font.Bold = $True
    }
    if ($Italic){
        $Object.ActiveDocument.Styles($Style).Font.Italic = $True
    }
    if ($Underline){
        $Object.ActiveDocument.Styles($Style).Font.Underline = wdUnderlineSingle
        $Object.ActiveDocument.Styles($Style).Font.UnderlineColor = wdColorAutomatic
    }

<#Sub Macro9()
'
' Macro9 Macro
'
'
    With ActiveDocument.Styles("Geen afstand").Font
        .Name = "Calibri"
        .Size = 11
        .Bold = True
        .Italic = True
        .Underline = wdUnderlineSingle
        .UnderlineColor = wdColorAutomatic
        .StrikeThrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Scaling = 100
        .Kerning = 1
        .Animation = wdAnimationNone
        .Ligatures = wdLigaturesStandardContextual
        .NumberSpacing = wdNumberSpacingDefault
        .NumberForm = wdNumberFormDefault
        .StylisticSet = wdStylisticSetDefault
        .ContextualAlternates = 0
    End With
    With ActiveDocument.Styles("Geen afstand")
        .AutomaticallyUpdate = False
        .BaseStyle = ""
        .NextParagraphStyle = "Geen afstand"
    End With
End Sub
>#>
    }Catch{
        Write-log -Category Error -message "[$FunctionName] : Unknown error. "
        Write-log -Category Error -message "[$FunctionName] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$FunctionName] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$FunctionName] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$FunctionName] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$FunctionName] : Errormessage : $($_.Exception.message)"   
    }
}