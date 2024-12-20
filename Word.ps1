$msoLanguageIDDutch = 1043 # Word variabele
$msoLanguageIDEnglishUS = 1033 # Word variabele
$wdAdjustNone = 0 # Word variabele
$wdAutoFitFixed = 0 # Word variabele
$wdAlignParagraphCenter = 1 # Word variabele
$wdBorderLeft = -2 # Word variabele
$wdBorderRight = -4 # Word variabele
$wdBorderTop = -1 # Word variabele
$wdBorderBottom = -3 # Word variabele
$wdBorderHorizontal = -5 # Word variabele
$wdBorderVertical = -6 # Word variabele
$wdBorderDiagonalDown = -7 # Word variabele
$wdBorderDiagonalUp = -8 # Word variabele
$wdBulletGallery = 1 # Word variabele
$wdCell = 12 # Word variabele
$wdCharacter = 1 # Word variabele
$wdColumn = 9 # Word variabele
$wdExtend = 1 # Word variabele
$wdFieldEmpty = -1 # Word variabele
$wdFormatXMLDocument = 12 # Word variabele
$wdGoToBookmark = -1 # Word variabele
$wdLine=5 # Word variabele
$wdLineStyleNone = 0 # Word variabele
$wdNumberParagraph = 1 # Word variabele
$wdPageBreak = 7 # Word variabele
$wdRow = 10 # Word variabele
$wdStory = 6 # Word variabele
$wdTable = 15 # Word variabele
$wdToggle=9999998  # Word variabele
$wdUnderlineNone =0 # Word variabele
$wdUnderlineSingle =1 # Word variabele
$Wdword = 2 # Word variabele
$wdWord9TableBehavior = 1 # Word variabele
$wdPropertyTitle = 1    # Word variabele
$wdSeekMainDocument = 0 # Word variabele
$wdSeekCurrentPageHeader = 9 # Word variabele
$wdSeekCurrentPageFooter = 10 # Word variabele
$Quote = [Char]34 #Variabele om een Quote weer te geven
$Tab = [Char]9 #Variabele om een tab weer te geven

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
        Write-log -Category Error -message "[Functionname] : Targetname   : $(_.CategoryInfo.targetname)"
        Write-log -Category Error -message "[Functionname] : Fullname     : $(_.exception.gettype.Fullname)"
        Write-log -Category Error -message "[Functionname] : Type fout    : $(_.CategoryInfo.category)"
        Write-log -Category Error -message "[Functionname] : Position     : $(_.invocationinfo.positionmessage)"
        Write-log -Category Error -message "[Functionname] : Errormessage : $(_.Exception.message)"
    }
}
Function Write-log{
    Param ([Validateset("Verbose","Debug","Error")]$Category, $Message)
    Write-host $Message
}
Function TypeTekst{
    param($Object,$bold=$False, $Italic=$False, $Underline=$False, $Fontsize, $Fontname,  $Tekst,  $Enter=$False)
    $FunctionName = "TypeTekst"
    Write-log Verbose -Message "`$Tekst = `'$Tekst`'"
    Write-log Verbose -Message "`$bold = `'$bold`'"
    Write-log Verbose -Message "`$Italic = `'$Italic`'"
    Write-log Verbose -Message "`$Underline = `'$Underline`'"
    Try{
        If ($Bold){
            $Object.Selection.Font.Bold = $wdToggle
        }
        If ($Italic){
            $Object.Selection.Font.Italic = $WDToggle
        }
        If ($Underline){
            If($Object.Selection.Font.Underline -eq $WdUnderlineNone){
                $Object.Selection.Font.Underline = $WdUnderlineSingle
            }else{
                $Object.Selection.Font.Underline = $WdUnderlineNone
            }
        }
        If ($FontName){
            $Object.Selection.Font.Name = $FontName
        }
        If ($FontSize){
            $Object.Selection.Font.Size = $FontSize
        }
            $Object.Selection.TypeText("$Tekst")
            $Object.Selection.Font.Italic = $False
            $Object.Selection.Font.Bold = $False
                $Object.Selection.Font.Underline = $WdUnderlineNone
        if ($Enter){
            $Object.Selection.TypeParagraph()
        }
    }Catch{
        Write-log -Category Error -message "[$Functionname] : Unknown error. "
        #Write-log -Category Error -message "[$Functionname] : Targetname   : $(_.CategoryInfo.targetname)"
        #Write-log -Category Error -message "[$Functionname] : Fullname     : $(_.exception.gettype.Fullname)"
        #Write-log -Category Error -message "[$Functionname] : Type fout    : $(_.CategoryInfo.category)"
        #Write-log -Category Error -message "[$Functionname] : Position     : $(_.invocationinfo.positionmessage)"
        #Write-log -Category Error -message "[$Functionname] : Errormessage : $(_.Exception.message)"
    }
}
$Document = "C:\Data\Github\Pronexum\Doc1.docx"
OpenWord -document $Document
typetekst -tekst "Marten" -Object $WORD -Enter $true
typetekst -tekst "Marten" -Object $WORD -bold $true -Enter $True
typetekst -tekst "Marten" -Object $WORD -Italic $True  -Enter $True
typetekst -tekst "Marten" -Object $WORD -Underline $True -Enter $True
typetekst -tekst "Marten" -Object $WORD -bold $true -Italic $true -Underline $true -Enter $true
Function VulDocumentProperties{
    Param ($Property)
    #$FunctionName = "VulDocumentProperties"

    ActiveDocument.BuiltinDocumentProperties($Property) = waarde
    ActiveDocument.CustomDocumentProperties.add name:=$Property,LinkToContent:False,Type:=PropertyType, Value := waarde
}

$binding = "System.Reflection.BindingFlags" -as [type]
$builtinProperties = $WORD.ActiveDocument.BuiltInDocumentProperties
[Array]$getArgs = "Title"
$builtinPropertiesType = $builtinProperties.GetType()
$builtinProperty = $builtinPropertiesType.InvokeMember("Item", $binding::GetProperty, $null, $builtinProperties, $getArgs)
$builtinPropertyType = $builtinProperty.GetType()
[Array]$setArgs = $ScriptbaseName
$builtinPropertyType.InvokeMember( "Value", $binding::SetProperty,  $null, $builtinProperty, $setArgs)
