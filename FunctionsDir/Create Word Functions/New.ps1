Function CheckStyle{
	Param ($Stijl)
	#Deze functie bepaald de stijl voor de tekst opmaak aan de hand van de taal die geinstalleerd is op de computer
	<# Start variabelen#>
	#$Word In deze variabele staat de DDE-Link naar Word
	#$Stijl In deze variabele staat de stijl die gecontroleerd moet worden
	#CheckStyle In deze variabele staat welke Style er wordt gebruikt voor de opmaak van het document
	#$msLanguageIDDutch In deze variabele staat de code voor de Nederlandse taal
	#$msoLanguageIDEnglishUS In deze variabele staat de code voor de Engelse taal
	<# End variabelen#>
	$CheckStyle = $Stijl 
    Switch ($word.Application.Language){
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