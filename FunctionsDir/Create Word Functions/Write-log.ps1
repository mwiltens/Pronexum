Function Write-log{
    Param ([Validateset("Verbose","Debug","Error")]$Category, $Message)
    Write-host $Message
}