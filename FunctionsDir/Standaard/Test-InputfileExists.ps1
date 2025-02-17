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
