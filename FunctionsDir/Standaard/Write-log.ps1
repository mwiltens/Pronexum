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
