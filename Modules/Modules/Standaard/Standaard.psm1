#region Assemblies
	[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
	[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
	[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
#endregion Assemblies
<#
.Synopsis
   Deze functie heet een korte omschrijving
.DESCRIPTION
   Deze functie heeft een lange omschrijving
.EXAMPLE
   Example of how to use this cmdlet
   Display_info -omschrijving "Dit is een test", $Waarde, Magenta
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>Function Display_Info{
    Param ($Omschrijving, $Waarde, [ValidateSet("Black","Blue" ,"Cyan" ,"DarkBlue","DarkCyan","DarkGray","DarkGreen","DarkMagenta","DarkRed","DarkYellow","Gray","Green","Magenta","White","Red","Yellow")]$Color)  
    $Functionname = "Display_Info"
    Try{
        $Message = "$Omschrijving".padright($tab," ") + ": $Waarde"
        Write-log -Category host -message $Message -color $Color
    }Catch{
        Write-log -Category Error -message "[$Functionname] : Unknown Error. `r`n Position: $($_.invocationinfo.positionmessage). `r`n Errormessage : $($_.Exception.message)"   
    }
}

    <#
    .Synopsis
       Deze functie plaatst een messagebox op het scherm
    .DESCRIPTION
       Deze functie plaatst een messagebox op het scherm
    .EXAMPLE
       MessageBox "Dit is een test", "Test",32, 32
    .NOTES
       $Button kan de volgende waarden hebben : AbortRetryIgnore, OK, OKCancel, RetryCancel, YesNo, YesNoCancel
    .COMPONENT
       The component this cmdlet belongs to
    .ROLE
       The role this cmdlet belongs to
    .FUNCTIONALITY
       The functionality that best describes this cmdlet
    #>
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.visualbasic') | Out-Null
Function MessageBox{
    Param($Messagetext, $MessageTitle,$Button, $Icon)
    $top = new-Object System.Windows.Forms.Form -property @{Topmost=$true}
    $msgBox1=[system.windows.forms.messageboxbuttons]::$Button;
    return [system.windows.forms.messagebox]::Show($top,$Messagetext,$MessageTitle,$msgBox1,$icon)
}

<#
.Synopsis
   Deze functie schrijft een melding in een bestand
.DESCRIPTION
   Deze melding schrijft een melding in een standaard logile. 
   Indien de melding ook op het scherm moet worden getoond, moet de optie 'Host' als category worden opgegeven

   Wordt er als category 'LOG' gekozen dan wordt de melding alleen in een opgegeven LOG-bestand weggeschreven

.EXAMPLE
   Write-log -Category Debug -message "Message"
.EXAMPLE
   Write-log -Category Error -message "[$Functionname] : Execute : "
.EXAMPLE
   Write-log -Category Host -message "Message" -Color Yellow
.EXAMPLE
   Write-log -Category Log -message "Message" -$LOGFILE
.EXAMPLE
   Write-log -Category Verbose -message "Message"
.EXAMPLE
   Write-log -Category Warning -message "Message"
.INPUTS
   
.OUTPUTS
   Log-file 
#>
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

