$RecordAzNetworkUsage = [Ordered]@{
	CurrentValue = ""
 	Limit = ""
 	Name = ""
 	ResourceType = ""
 	Unit = ""

}
$RecordAzNetworkWatcher = [Ordered]@{
	Etag = ""
 	Id = ""
 	Location = ""
 	Name = ""
 	ProvisioningState = ""
 	ResourceGroupName = ""
 	ResourceGuid = ""
 	Tag = ""
 	Type = ""

}
Function Inventariseer_AzNetworkWatcher{
    $Global:AzNetworkWatcherRecords = @()
    $AzNetworkWatchers = Get-AzNetworkWatcher | Select *
    Foreach ($AzNetworkWatcher in $AzNetworkWatchers){
        $Data = New-Object -TypeName PSObject -Property $RecordAzNetworkWatcher
		$Data.Etag = $AzNetworkWatcher.Etag
 		$Data.Id = $AzNetworkWatcher.Id
 		$Data.Location = $AzNetworkWatcher.Location
 		$Data.Name = $AzNetworkWatcher.Name
 		$Data.ProvisioningState = $AzNetworkWatcher.ProvisioningState
 		$Data.ResourceGroupName = $AzNetworkWatcher.ResourceGroupName
 		$Data.ResourceGuid = $AzNetworkWatcher.ResourceGuid
 		$Data.Tag = $AzNetworkWatcher.Tag
 		$Data.Type = $AzNetworkWatcher.Type

        $Global:AzNetworkWatcherRecords += $Data  
    }
}
Inventariseer_AzNetworkWatcher
$RecordAzNetworkWatcher
