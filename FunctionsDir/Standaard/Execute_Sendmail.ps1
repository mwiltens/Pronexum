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
