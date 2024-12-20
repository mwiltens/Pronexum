Try{
    #Import-Module "$env:USERPROFILE\documents\modules\Standaard\Standaard.psm1"
    Import-Module "E:\Data\Scripts\Modules\Create_Word\0.3\Create_Word.psm1"
}catch{
    Write-Error "Import-module not found"
}
$Document = "E:\Data\Test.docx"
#$Document = "C:\Data\Github\Pronexum\Doc1.docx"
OpenWord -document $Document
typetekst -tekst "Marten" -Object $WORD -Enter $true
typetekst -tekst "Marten" -Object $WORD -bold $true -Enter $True
typetekst -tekst "Marten" -Object $WORD -Italic $True  -Enter $True
typetekst -tekst "Marten" -Object $WORD -Underline $True -Enter $True
typetekst -tekst "Marten" -Object $WORD -bold $true -Italic $true -Underline $true -Enter $true