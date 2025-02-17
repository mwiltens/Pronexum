$Basefile = "E:\data\scripts\Base.txt"
$FieldsFile = "E:\data\scripts\Fields.txt"
$FunctionName = "AzADUser"
$Parameter = "-location westeurope"
$Parameter = ""
$Command = "Get-$FunctionName"
$base = (Get-Content $Basefile) 
$ExecuteCommand = "($command $Parameter | Select * -First 1)"
Invoke-Expression -Command $ExecuteCommand
$ExecuteCommand = "($command $Parameter | Get-member |where {`$_.MemberType -eq `"Property`"}).Name |set-clipboard"
$result = Invoke-Expression -Command $ExecuteCommand
Get-Clipboard | Set-Content $FieldsFile

#Write-host $ExecuteCommand

Start-Process "Notepad.exe" -ArgumentList $FieldsFile -Wait
Wait-Process -Name Notepad
$Fields = Get-Content $FieldsFile
$Newbase = $null
$Base = $Base.Replace("[Naam]",$FunctionName)
$RecFields = @()
Foreach ($Field in $Fields){
    $RecFields += "`t$Field = `"`"`n"
}
$DataFields = @()
Foreach ($Field in $Fields){
    $DataFields += "`t`t`$Data.$Field = `$$FunctionName.$Field`n"
}
$Base = $Base.Replace("[RecFields]",$RecFields)
$Base = $Base.Replace("[DataFields]",$DataFields)
$Base = $Base.Replace("[Command]",$Command)
$Base = $Base.Replace("[Results]","`$$FunctionName$("s")")
$Base = $Base.Replace("[Result]","`$$FunctionName")
cls
$Base | Set-Clipboard
