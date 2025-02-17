<#
 .Synopsis
  Start MS-Word 

 .Description
  Start de applicatie MS-Word
  
 .Parameter $Document
  Indien deze variabele gevuld is, wordt het opgegeven document in $Document geopend

 .Example
  # Start MS-word en opent het bestand "E:\Data\Test.docx"
  OpenWord -document "E:\Data\Test.docx"
  
  # Start MS-word zonder een document te openen
  OpenWord

#>

Function OpenWord{
    param ($Document)
    $Functionname = "OpenWord"
    Try{
        Write-host "[$Functionname] : Execute : `$GLOBAL:WORD = New-Object -comobject Word.Application"
        $GLOBAL:WORD = New-Object -comobject Word.Application
        Write-log -Category Verbose -Message "[$Functionname] : Execute : `$WORD.Visible = $True"
        $WORD.Visible = $True
        If ($Document -eq ""){
            Write-log -Message "[$Functionname] : Execute : `$GLOBAL:DOC = $WORD.Documents.add()"
            $GLOBAL:DOC = $WORD.Documents.add()
        }else {
            Write-log -Message "[$Functionname] : Execute : `($Word.Documents.Open($Document)) | out-null"
            $GLOBAL:DOC = ($Word.Documents.Open($Document))# | out-null
        }

    }catch{
        Write-log -Category Error -message "[$FunctionName] : Unknown error. "
        Write-log -Category Error -message "[$FunctionName] : Targetname   : $($_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[$FunctionName] : Fullname     : $($_.exception.gettype().Fullname)"
        Write-log -Category Error -message "[$FunctionName] : Type fout    : $($_.CategoryInfo.category)"
        Write-log -Category Error -message "[$FunctionName] : Position     : $($_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[$FunctionName] : Errormessage : $($_.Exception.message)"   
    }
}