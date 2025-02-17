<#
 .Synopsis
  Aanmaken van een tabel

 .Description
  Aanmaken van een tabel in Word
  
 .Parameter Object


 .Parameter Rijen


 .Parameter Kolommen

 .Parameter Kolombreedtes

 .Parameter Borderlines

 .Parameter TabelNr


 .Example
   # Voegt de verboselogfile dat aangemaakt is op de D-schijf toe aan de Verbosefile.
   Read_RemoteLogfile -RemoteLogfile "\\$([Servernaam]\[Schijf]\$([scriptname])_Verbose.txt" -Logfile $Global:VerboseFile

#>
Function Maak_Kolommen{
	Param ($Object, 
    $Rijen, 
    $Kolommen,
    $Kolombreedtes,
    [Validateset($true,$False)]$Borderlines,
    $TabelNr=1,
    [Validateset($true,$False)]$BorderBottom = $False, 
    [Validateset($true,$False)]$BorderTop = $False, 
    [Validateset($true,$False)]$BorderLeft = $False, 
    [Validateset($true,$False)]$BorderRight = $False, 
    [Validateset($true,$False)]$BorderVertical = $False, 
    [Validateset($true,$False)]$BorderDiagonalDown = $False, 
    [Validateset($true,$False)]$BorderDiagonalUp = $False
    )
    $FunctionName = "Maak_Kolommen"
    Try{
        $Range = $Global:WordSelection.range
        
        #$Result = $Object.ActiveDocument.Tables.Add($Range, $Rijen, $Kolommen, $wdWord9TableBehavior, $wdAutoFitFixed)#| Out-Null
        #$Table = $Object.ActiveDocument.Tables.Add($Object.Selection.Range, $Rijen, $Kolommen, $wdWord9TableBehavior, $wdAutoFitFixed)#| Out-Null
        $Table = $Object.ActiveDocument.Tables.Add($Object.Selection.Range, $Rijen, $Kolommen)#| Out-Null
	    #$Table = $Object.ActiveDocument.Tables.item($TabelNr)	
	    #$Table = $GLOBAL:DOC.Tables.item(1)
        $Table.borders.enable = $true	
        If ($Borderlines -eq $False){
   	        $Table.Borders.InsideLineStyle = 0
   	        $Table.Borders.OutsideLineStyle = 0	
        }
        if ($BorderBottom -eq $true){
            $Table.Borders($wdBorderBottom).LineStyle = 1
        }
        if ($BorderTop -eq $true){
            $Table.Borders($wdBorderTop).LineStyle = 1
        }
        if ($BorderLeft -eq $true){
            $Table.Borders($wdBorderLeft).LineStyle = 1
        }
        if ($BorderRight -eq $true){
            $Table.Borders($wdBorderRight).LineStyle = 1
        }
        if ($BorderVertical -eq $true){
            $Table.Borders($wdBorderVertical).LineStyle = 1
        }
        if ($BorderDiagonalDown -eq $true){
            $Table.Borders($wdBorderDiagonalDown).LineStyle = 1
        }
        if ($BorderDiagonalUp -eq $true){
            $Table.Borders($wdBorderDiagonalUp).LineStyle = 1
        }
        $I = 1
	    Foreach ($Kolombreedte in $Kolombreedtes){
            $Kolom = $Table.columns.item($I)
	        $Kolom.PreferredWidth = InchesToPoints $Kolombreedte
            $I++
        }
    }catch{
        Write-log -Category Error -message "[$FunctionName] : Fout bij het toevoegen van een tabel. "
        Write-log -Category Error -message "[$Functionname] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $($_.Exception.message)"   
    }
}
