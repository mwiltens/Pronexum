Cls
Try{
    #Import-Module "$env:USERPROFILE\documents\modules\Standaard\Standaard.psm1"
    #Import-Module "E:\Data\Scripts\Modules\Create_Word\0.3\Create_Word.psm1"
    #Import-Module "E:\Data\Scripts\Pronexum\Pronexum\Modules\Create_Word\0.3\Create_Word.psm1"
    #Import-Module "E:\Data\Scripts\Modules\Word\0.4\Word.psm1"
    #Import-Module "E:\Data\Scripts\Pronexum\Pronexum\Modules\Word\0.4\Word.psm1"
    Import-Module "E:\Data\Scripts\Pronexum\Pronexum\Modules\Standaard\0.3\Standaard.psm1"

    #Import-Module Word # -RequiredVersion 0.2
}catch{
    Write-Error "Import-module not found"
}
$SharedPath = "E:\Data\Scripts\FunctionsDir\"
$SharedPath = "E:\Data\Scripts\Pronexum\Pronexum\FunctionsDir\"
$FunctionsPath = join-path -Path $SharedPath -ChildPath "Create Word Functions"
$FunctionFiles = Get-ChildItem -Path $FunctionsPath -Filter '*.ps1' # -Recurse
$VerbosePreference = "Continue"
$VerbosePreference = "SilentlyContinue"
foreach  ($Function in $FunctionFiles){
    Write-verbose -Message "Including '$($Function.fullname)'"
    . $($Function.fullname)
}
$VerbosePreference = "SilentlyContinue"
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
Function Item{
    Param ($Object, $Tekst,$Style)
    typetekst -tekst $Tekst -Object $Object -Enter $true
    Move-Cursor -Object $Object -Direction Up -Unit $wdline -Count 1
    Selecteer_Regel -Object $Object
    Style -Object $Object -Style $Style | Out-Null
}
If ([bool](Get-Process -Name WINWORD -ErrorAction SilentlyContinue)){
    #Get-Process -Name WINWORD | Stop-Process
}

$Document = "F:\Data\Test.docx"
#$Document = "C:\Data\Github\Pronexum\Doc1.docx"
#https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdfieldtype?view=word-pia
#Test_Foutmelding
#return


OpenWord -document $Document
$Global:WordSelection = $Global:WORD.Selection    

    Style -Object $Word -Style "Geen Afstand"
    Update_Style -Object $Word -Style "Geen Afstand" -Fontname "Calibri" -Fontsize 11 
    


    #$Global:WordSelection.Selection.Style = $Global:WordSelection.ActiveDocument.Styles("Geen afstand")

    #Call Regellager
    #Call Enter
    #Call Maak_Kolommen(2)
    #Call ZoekBookmark("Pagina2")
    #Maak-Legenda -Object $Word
    #"E:\Data\Sjablonen\Edulas.jpg"
    #E:\Data\Sjablonen\Bloemendaal.png
    #E:\Data\Sjablonen\Pronexum.png
    #Aanmaken-KopVoetTekst -Object $Word -Type Koptekst -Picture "E:\Data\Sjablonen\Bloemendaal.png"
    #Aanmaken-KopVoetTekst -Object $Word -Type Koptekst -Picture "E:\Data\Sjablonen\Pronexum.png"
    Aanmaken-KopVoetTekst -Object $Word -Type Koptekst -Picture "E:\Data\Sjablonen\Edulas.jpg"
    Aanmaken-KopVoetTekst -Object $Word -Type Voettekst 
    Maak-Legenda -Object $Word
    <#
    $Selection = $Global:WORD.Selection
    $Range = $Selection.Range
    #$Global:Doc.fields.add($Range, 21, "CREATEDATE  \@ ""d-MM-yyyy"" ", $true)
    typetekst -tekst "" -Object $WORD -Enter $False
    Move-Cursor -Object $Word -Direction Right -Unit $wdCharacter -Count 1
    typetekst -tekst "\" -Object $WORD -Enter $False
    Move-Cursor -Object $Word -Direction Right -Unit $wdCharacter -Count 1
    $Range = $Selection.Range
    $Global:Doc.fields.add($Range, 61, "\* Upper", $true)| Out-Null
    #typetekst -tekst "" -Object $WORD -Enter $true
    Move-Cursor -Object $Word -Direction Down -Unit $wdLine -Count 1
    #>
Item -Object $Word -Tekst "Installatie gegevens" -Style "Kop 1"
Item -Object $Word -Tekst "Installatie gegevens" -Style "Kop 2"
Item -Object $Word -Tekst "Kop 3" -Style "Kop 3"
Item -Object $Word -Tekst "Kop 4" -Style "Kop 4"
Item -Object $Word -Tekst "Kop 5" -Style "Kop 5"
Item -Object $Word -Tekst "Kop 6" -Style "Kop 6"
Item -Object $Word -Tekst "Kop 7" -Style "Kop 7"

$Global:toc.Update()

$word.Documents.close($False)
$word.Quit()

return
Item -Object $Word -Tekst "Installatie gegevens" -Style "Kop 1"
Item -Object $Word -Tekst "Installatie gegevens" -Style "Kop 2"
Item -Object $Word -Tekst "Kop 3" -Style "Kop 3"
Item -Object $Word -Tekst "Kop 4" -Style "Kop 4"
Item -Object $Word -Tekst "Kop 5" -Style "Kop 5"
Item -Object $Word -Tekst "Kop 6" -Style "Kop 6"
Item -Object $Word -Tekst "Kop 7" -Style "Kop 7"

#$Global:toc.Update()
<#
typetekst -tekst "Marten" -Object $WORD -Enter $true
typetekst -tekst "Marten" -Object $WORD -bold $true -Enter $True
typetekst -tekst "Marten" -Object $WORD -Italic $True  -Enter $True
typetekst -tekst "Marten" -Object $WORD -Underline $True -Enter $True
typetekst -tekst "Marten" -Object $WORD -bold $true -Italic $true -Underline $true -Enter $true
Add-Table -Object $WORD -row 2 -Column 3 -BackgroundColor 3
#typetekst -tekst "Dennis" -Object $WORD -bold $true -Italic $true -Underline $true -Enter $False
typetekst -tekst "Dennis" -Object $WORD -Enter $False
Move-Cursor -Object $Word -Direction Right -Unit $wdCharacter
typetekst -tekst "Jacq" -Object $WORD -Enter $False
Move-Cursor -Object $Word -Direction Right -Unit $wdCharacter
typetekst -tekst "Yvonne" -Object $WORD -Enter $False
Move-Cursor -Object $Word -Direction Down -Unit $wdLine
Move-Cursor -Object $Word -Direction Left -Unit $wdCharacter -Count 2
typetekst -tekst "Maurice" -Object $WORD -Enter $False
Move-Cursor -Object $Word -Direction EndKey -Unit $wdStory
#$Word.Selection.Endkey($wdStory)
<#
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=2

#>
#$Word.Selection.Endkey($wdLine)
return
#$WORD.Selection.EndKey($wdLine, $wdExtend) | Out-Null
#Add-Table -Object $WORD -row 2 -Column 3 -BackgroundColor 0
#typetekst -tekst "Dennis" -Object $WORD -Enter $False
#$WORD.Selection.EndKey($wdLine, $wdExtend) | Out-Null
#Add-Table -Object $WORD -row 2 -Column 3 -BackgroundColor 1
#typetekst -tekst "Dennis" -Object $WORD -Enter $False
#$WORD.Selection.EndKey($wdLine, $wdExtend) | Out-Null
#Add-Table -Object $WORD -row 2 -Column 3 -BackgroundColor 2
#typetekst -tekst "Dennis" -Object $WORD -Enter $False



