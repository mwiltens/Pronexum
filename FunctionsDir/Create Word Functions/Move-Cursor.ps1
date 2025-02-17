<#
 .Synopsis
  Verplaatst de cursor

 .Description
  Verplaatst de cursor door het Word-document
  
 .Parameter $Object
  Het object, waarop de bewerking plaatsvindt.

 .Parameter $Direction
  Verplaatst de cursor een bepaalde kant op.
  Mogelijke waarden zijn : Right, Down, Left, Endkey en Homekey


 .Parameter Unit
  Mogelijke waarden zijn: $wdCharacter, $wdLine, $wdStory


 .Parameter Count
  Het aantal posities wat de cursor moet opschuiven. 
  Indien de waarde niet is opgegeven dat heeft deze de waarde 1


 .Example
   # Voegt de verboselogfile dat aangemaakt is op de D-schijf toe aan de Verbosefile.  
   # Verplaatst de cursor 1 karakter naar Rechts
   Move-Cursor -Object $Word -Direction Right -Unit $wdCharacter

   # Verplaatst de cursor 1 regel naar beneden
   Move-Cursor -Object $Word -Direction Down -Unit $wdLine
   
   # Verplaatst de cursor 2 karakters naar Links
   Move-Cursor -Object $Word -Direction Left -Unit $wdCharacter -Count 2

   # Verplaatst de cursor Naar het einde van de document
   Move-Cursor -Object $Word -Direction EndKey -Unit $wdStory

   # Verplaatst de cursor Naar het einde van de regel
   Move-Cursor -Object $Word -Direction EndKey -Unit $wdLine

#>

Function Move-Cursor{
    Param ($Object,[Validateset("Right","Left","EndKey","Up","Down", "HomeKey")]$Direction, $Unit,$Count=1)
    $Functionname = "Move-Cursor"
    Try{

        Switch ($Direction){ 
            "Right"
                {
                    $Object.Selection.MoveRight($Unit,$Count) | Out-Null
                }
            "Left"
                {
                    $Object.Selection.MoveLeft($Unit,$Count) | Out-Null
                }
            "EndKey"
                {
                    $Object.Selection.Endkey($Unit,$Count) | Out-Null
                }
            "Up"
                {
                    $Object.Selection.MoveUp($Unit,$Count) | Out-Null
                }
            "Down"
                {
                    $Object.Selection.MoveDown($Unit,$Count) | Out-Null
                }
            "HomeKey"
                {
                    $Object.Selection.HomeKey($Unit) | Out-Null
                }
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
