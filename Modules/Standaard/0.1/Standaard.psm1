<#
 .Synopsis
  Plaats een omschrijving met de opgegeven waarde op het scherm

 .Description
  Plaats een omschrijving met de opgegeven waarde op het scherm
  
 .Parameter Omschrijving
  De omschrijving die op het scherm moet worden getoond

 .Parameter Waarde
  De waarde die achter de omschrijving moet worden getoond

 .Parameter Color
  De kleur die de de getoond tekst moet gaan krijgen


 .Example
   # De omschrijving "Time start" wordt op het scherm getoond met de bijbehoorde waarde in het groen
   Display_info -Omschrijving "Time start" -Waarde $time -Color Green

#>
Function Display_Info{
    Param ($Omschrijving, $Waarde, [ValidateSet("Black","Blue" ,"Cyan" ,"DarkBlue","DarkCyan","DarkGray","DarkGreen","DarkMagenta","DarkRed","DarkYellow","Gray","Green","Magenta","White","Red","Yellow")]$Color)  
    $Functionname = "Display_Info"
    Try{
        $Message = "$Omschrijving".padright($tab," ") + ": $Waarde"
        Write-log -Category host -message $Message -color $Color
    }Catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "   
        Write-log -Category Error -message "[$Functionname] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $($_.Exception.message)"   
    }
}

<#
 .Synopsis
  Versturen van een email

 .Description
  Deze functie verstuurd een email-bericht
  
 .Parameter Subject
  Welk onderwerp heeft de email

 .Parameter Body
  De teskt die in de body van de mail wordt verstuurd

 .Parameter Priority
  Met welke prioriteit wordt de mail verstuurd

 .Parameter Attachment
  In deze parameter wordt het bijgesloten bestand opgegeven

 .Parameter To
  Naar wie wordt de mail verstuurd

 .Parameter bodyashtml
  Wordt de body met HTML opmaak verstuurd

 .Parameter From
  Door wie wordt de mail gestuurd

 .Example
   # Er wordt een mail verstuurd naar Testmail@google.com met een hoge prioriteit .
   Execute_Sendmail -Subject "Test onderwerp" -Body "Dit is een testmail" -Priority High -Attachment -To Testmail@google.com -bodyashtml $False -From Mailtest@yahoo.com

 .Example
   # Er wordt een mail verstuurd naar Testmail@google.com met een lage prioriteit. De mail heeft als bijlage het bestand "D:\voorgang.txt"
   Execute_Sendmail -Subject "Test onderwerp" -Body "Dit is een testmail" -Priority Low -Attachment "D:\voorgang.txt" -To Testmail@google.com -bodyashtml $True -From Mailtest@yahoo.com
#>
Function Execute_Sendmail{
    Param ($Subject,$Body,[Validateset("High","Low","Normal")]$Priority = "Normal",$Attachment="",$To,[Validateset($True,$False)]$bodyashtml,$From)
    $Functionname = "Execute_Sendmail"
    Write-log -Category Debug -message "[$Functionname] : `$To = `'$To`' "
    Write-log -Category Debug -message "[$Functionname] : `$From = `'$From`' "
    Write-log -Category Debug -message "[$Functionname] : `$Subject = `'$Subject`' "
    Write-log -Category Debug -message "[$Functionname] : `$Body = `'`$Body`' "
    Write-log -Category Debug -message "[$Functionname] : `$Priority = `'$Priority`' "
    Write-log -Category Debug -message "[$Functionname] : `$Attachment = `'$Attachment`' "
    Write-log -Category Debug -message "[$Functionname] : `$bodyashtml = `'$bodyashtml`' "
    Try{
        $KopVoettekst = "Mail versturen"
        KopVoettekst -Tekst $KopVoettekst -Functie Kop
        Write-log -Category Host -message "[$Functionname] : Execute : $KopVoettekst" -color Magenta
        Write-log -Category Log -message "[$(Get-date -Format "dd-MM-yyyy HH:mm")] $KopVoettekst" -logfile $Global:VoortgangFile
        If ($Global:testrun){
            $SendTo = $global:TestTO
        }else{
            $SendTo = $global:To
        }
        Write-log -Category Verbose -message "[$Functionname] : Execute :  Sendmail -To $SendTo -Subject $Subject -Priority $Priority -Body $Body -bodyashtml $bodyashtml -Attachment $Attachment -From $From"
        Sendmail -To $SendTo -Subject $Subject -Priority $Priority -Body $Body -bodyashtml $bodyashtml -Attachment $Attachment -From $From 
        Write-log -Category Verbose -message "[$Functionname] : Execute :  Write-log -Category Log -message `"[$(Get-date -Format "dd-MM-yyyy HH:mm")] $KopVoettekst is uitgevoerd`""
        Write-log -Category Log -message "[$(Get-date -Format "dd-MM-yyyy HH:mm")] $KopVoettekst is uitgevoerd" 
        KopVoettekst -Tekst $KopVoettekst -Functie Voet
    }catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "   
        Write-log -Category Error -message "[$Functionname] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $($_.Exception.message)"   
    }
}

<#
 .Synopsis
  Korte omschrijving

 .Description
  Omschrijving.
  
 .Parameter Messagetext
  Het bericht voor de message-box

 .Parameter MessageTitle
  De titel voor de message-box
 .Parameter Button

 .Parameter Icon
  In deze variabele wordt een icon opgegeven die door de message-box wordt getoont

 .Example
   # Wordt later toegevoegd
   

#>
Function MessageBox{
    Param($Messagetext, $MessageTitle,$Button, $Icon)
    $top = new-Object System.Windows.Forms.Form -property @{Topmost=$true}
    $msgBox1=[system.windows.forms.messageboxbuttons]::$Button;
    return [system.windows.forms.messagebox]::Show($top,$Messagetext,$MessageTitle,$msgBox1,$icon)
}

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

<#
 .Synopsis
  Controleert of een bestand ook aanwezig is

 .Description
  Controleert of het bestand wat noodzakelijk is voor de werking van het script aanwezig is
  
 .Parameter TestFile
  Het bestand wat gecontroleerd moet gaan worden

 .Example
   # Controleert of het bestand "D:\Input\Test.csv" aanwezig is.
   Test-InputfileExists D:\Input\Test.csv

#>
Function Test-InputfileExists{
    Param ($Testfile)
    $Functionname = "Test-InputfileExists"
    Try{
        If (!(Test-Path $Testfile)){
            Write-log -Category Host -message "[$Functionname] : $Testfile is niet aanwezig" -color Red
            $PreReqResult = $false
            Return $PreReqResult
        }else{
            Write-log -Category verbose -message "[$Functionname] : Invoerbestand `'$Testfile`' is aanwezig" 
            $PreReqResult = $true
        }
        Return $PreReqResult
    }catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "   
        Write-log -Category Error -message "[$Functionname] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$Functionname] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$Functionname] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$Functionname] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$Functionname] : Errormessage : $($_.Exception.message)"   
    }

}

<#
 .Synopsis
  Aan de hand van de opgegeven status, wordt er bepaald of er aan de voorwaarden wordt voldaan, om het script te starten

 .Description
  Aan de hand van de opgegeven status, wordt er bepaald of er aan de voorwaarden wordt voldaan, om het script te starten
  
 .Parameter Status
 In deze variabele staat de opgegeven status

 .Example
   # De globale variabele wordt op True gezet. Er wordt voldaan aan de voorwaarden om het script verder uit te voeren
   Update_Status $true

 .Example
   # De globale variabele wordt op False gezet. Er wordt niet voldaan aan de voorwaarden om het script verder uit te voeren
   Update_Status $False

#>
Function Update_Status{
    Param ($Status)
    If ($Global:PreReqResult -ne $False ){
        $Global:PreReqResult = $Status
    }
}

Function Write-log{
    Param ($message, [Validateset('Host', 'LOG', 'Error', 'Verbose', 'Warning', 'Debug')]$Category = "Verbose", $logfile = $logfile, [ValidateSet("Black","Blue" ,"Cyan" ,"DarkBlue","DarkCyan","DarkGray","DarkGreen","DarkMagenta","DarkRed","DarkYellow","Gray","Green","Magenta","White","Red","Yellow")]$color = "White")
    Switch ($Category) {
        'Host'    { Write-Host $Message -ForegroundColor $Color}
        'LOG'     { }
        'Error'   { Write-host $Message -ForegroundColor Red -ErrorAction Continue}
        'Verbose' { Write-Verbose $Message}
        'Warning' { Write-Warning $Message}
        'Debug' { Write-Debug $Message}
        Default {throw 'Unknown Category'}
    }
    $time = Get-Date -Format "dd-MM-yy HH:mm"
    If ($LogFile) {
        Switch -Regex ($Category) {
            '^LOG$|^Warning$' {                
                write-Output  "$($Message)" | Out-File -Append $LogFile
            }
            '^Error$|^Warning$|^Verbose$|^Debug$|^Host$' {                
                write-Output  "[$time] - $("[$($Category)]".Padright(9)) $($Message)" | Out-File -Append $VerboseFile
            }
            '^Error$|^Warning$'   {
                write-Output  "[$time] - $("[$($Category)]".Padright(9)) $($Message)" | Out-File -Append $ERRFILE
            }
            Default {throw 'Unknown Category'}
        }                
    }    
}

