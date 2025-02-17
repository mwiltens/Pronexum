<#
 .Synopsis
  Korte omschrijving

 .Description
  Omschrijving.
  
 .Parameter Messagetext
  Het bericht voor de message-box

 .Parameter MessageTitle
  De titel voor de message-box
 .Parameter Button

 .Parameter Icon
  In deze variabele wordt een icon opgegeven die door de message-box wordt getoont

 .Example
   # Wordt later toegevoegd
   

#>
Function MessageBox{
    Param($Messagetext, $MessageTitle,$Button, $Icon)
    $top = new-Object System.Windows.Forms.Form -property @{Topmost=$true}
    $msgBox1=[system.windows.forms.messageboxbuttons]::$Button;
    return [system.windows.forms.messagebox]::Show($top,$Messagetext,$MessageTitle,$msgBox1,$icon)
}
