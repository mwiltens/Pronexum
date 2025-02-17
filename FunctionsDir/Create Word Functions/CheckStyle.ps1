<#
 .Synopsis
  Leest aan de hand van de taal (Engels of Nederlands) de Style-name uit

 .Description
  Leest aan de hand van de taal (Engels of Nederlands) de Style-name uit
  
 .Parameter Object
  Is het Object, waarop de bewerking wordt toegepast

 .Parameter Style
  De Style die opgezocht moet gaan worden

.Example
   # Zoekt in $Word aan de hand van de Taal de Style "Kop 1" op
   $Style = CheckStyle -Object $Word -Stijl "Kop 1"
#>
Function CheckStyle{
	Param ($Object,$Stijl)
    $Functionname = "CheckStyle"
    Try{
	    $CheckStyle = $Stijl 
        Switch ($Object.Application.Language){
		    $msoLanguageIDDutch {
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
                    "Geen Afstand"{
                        $CheckStyle = "Geen Afstand"
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
                    "Geen Afstand"{
                        $CheckStyle = "Geen Afstand"
				    }
                }
		    }
	    }
	    Return $CheckStyle
    }catch{
        Write-log -Category Error -message "[$FunctionName] : Unknown error. "
        Write-log -Category Error -message "[$FunctionName] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$FunctionName] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$FunctionName] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$FunctionName] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$FunctionName] : Errormessage : $($_.Exception.message)"   
    }
}