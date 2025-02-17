<#
 .Synopsis
  Leest de logfile uit, die op een remote server is aangemaakt.

 .Description
  Leest de logfile uit, die op een remote server is aangemaakt en voegt deze toe aan de door het script gebruikte logfile.
  Deze logfile wordt aangemaakt m.b.t. Invoke-command
  
 .Parameter RemoteLogFile
  De naam van de logfile op de remote server.

 .Parameter LogFile
  De naam van de logfile, dat door het script qordt gebruikt


 .Example
   # Voegt de verboselogfile dat aangemaakt is op de D-schijf toe aan de Verbosefile.
   Read_RemoteLogfile -RemoteLogfile "\\$([Servernaam]\[Schijf]\$([scriptname])_Verbose.txt" -Logfile $Global:VerboseFile

 .Example
   # Voegt de logfile dat aangemaakt is op de D-schijf toe aan de Logfile.
   Read_RemoteLogfile -RemoteLogfile "\\$([Servernaam]\[Schijf]\$([scriptname])_Log.txt" -Logfile $Global:LogFile

 .Example
   # Voegt de Errorfile dat aangemaakt is op de D-schijf toe aan de errorfile.
   Read_RemoteLogfile -RemoteLogfile "\\$([Servernaam]\[Schijf]\$([scriptname])_ERR.txt" -Logfile $Global:ErrFile

#>
Function Read_RemoteLogfile{
    Param ($RemoteLogfile,$Logfile)
    $Functionname = "Read_RemoteLogfile"
    Try{
        Write-log -Category Verbose -message "[$Functionname] : Execute :  `$RemoteLogfileRegels = gc $Remotelogfile "
        IF (test-path $Remotelogfile){
            $RemoteLogfileRegels = gc $Remotelogfile 
            Foreach ($Line in $RemoteLogfileRegels){
                write-Output  $Line | Out-File -Append $Logfile        
            }
        }
    }catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "   
        Write-log -Category Error -message "[$Functionname] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $($_.Exception.message)"   
    }
}
