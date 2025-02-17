<#
 .Synopsis
  Zoeken naar een bookmark

 .Description
  Wordt gebruikt om naar een bookmark in een word document te gaan.
  
 .Parameter Object
  Is het Object, waarop de bewerking wordt toegepast


 .Parameter BookmarkName
  De naam van het bookmark waarop gezocht moet worden

 .Parameter 1


 .Example
   # Voegt de verboselogfile dat aangemaakt is op de D-schijf toe aan de Verbosefile.
   Read_RemoteLogfile -RemoteLogfile "\\$([Servernaam]\[Schijf]\$([scriptname])_Verbose.txt" -Logfile $Global:VerboseFile

#>
Function Zoek-Bookmark{
    Param ($Object, $BookmarkName)
    $FunctionName = "Zoek-Bookmark"
    Try{
        $object.Selection.GoTo($wdGoToBookmark, $BookmarkName)
    }catch{
        Write-log -Category Error -message "[$FunctionName] : Unknown error. "
        Write-log -Category Error -message "[$FunctionName] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$FunctionName] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$FunctionName] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$FunctionName] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$FunctionName] : Errormessage : $($_.Exception.message)"   
    }
    <#
    Selection.GoTo What:=wdGoToBookmark, Name:=Bookmark
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
End Sub
#>

}
