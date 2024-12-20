Function OpenWord{
    param ($Document)
    $Functionname = "OpenWord"
    Try{
        Write-host "[$Functionname] : Execute : `$GLOBAL:WORD = New-Object -comobject Word.Application"
        $GLOBAL:WORD = New-Object -comobject Word.Application
        Write_log -Category Verbose -Message "[$Functionname] : Execute : `$WORD.Visible = $True"
        $WORD.Visible = $True
        If ($Document -eq ""){
            Write_log -Message "[$Functionname] : Execute : `$GLOBAL:DOC = $WORD.Documents.add()"
            $GLOBAL:DOC = $WORD.Documents.add()
        }else {
            Write_log -Message "[$Functionname] : Execute : `($Word.Documents.Open($Document)) | out-null"
            ($Word.Documents.Open($Document)) | out-null
        }

    }catch{
        Write-log -Category Error -message "[Functionname] : Unknown error. "
        #Write-log -Category Error -message "[Functionname] : Targetname   : $(_.CategoryInfo.targetname)"
        #Write-log -Category Error -message "[Functionname] : Fullname     : $(_.exception.gettype.Fullname)"
        #Write-log -Category Error -message "[Functionname] : Type fout    : $(_.CategoryInfo.category)"
        #Write-log -Category Error -message "[Functionname] : Position     : $(_.invocationinfo.positionmessage)"
        #Write-log -Category Error -message "[Functionname] : Errormessage : $(_.Exception.message)"
    }
}