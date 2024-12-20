Function CheckStyle{
	Param ($Stijl,$Object)
	$CheckStyle = $Stijl 
    Switch ($Object.Application.Language){
		$msoLanguageIDDutch {
            #MsgBox "The user interface language is Dutch."
            #Rem Dim StrHeading1, StrHeading2, StrHeading3, StrHeading4, StrHeading5, StrHeading6
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
            }
		}
	}
	Return $CheckStyle
}