cls
$StandardWindowsServiceFile = "U:\Standard_Service.txt"
$InfoDataRecord = [Ordered]@{
    
}
$InfoServiceRecord = [Ordered]@{
    Name = ""    
    Status = ""
}
$InfoProfileRecord = [Ordered]@{
    Name = "" 
    InterfaceAlias = "" 
    InterfaceIndex  = "" 
    NetworkCategory = "" 
    IPv4Connectivity = "" 
    IPv6Connectivity = "" 
}
$InfoNetAdapterRecord = [Ordered]@{
    IPAddress = ""
    InterfaceIndex = ""
    InterfaceAlias = ""
    AddressFamily = ""
    PrefixLength = ""
}
$InfoVolumeRecord = [Ordered]@{
    DriveType = ""
    FileSystemType = ""
    AllocationUnitSize = ""
    DriveLetter = ""
    FileSystem = ""
    FileSystemLabel = ""
    Size = ""
    SizeRemaining = ""
}
<#$InfoProfileRecord = [Ordered]@{

}
#>
$InfoOSRecord = [Ordered]@{
    WindowsEditionId = ""
    WindowsInstallationType = ""
    WindowsInstallDateFromRegistry = ""
    WindowsProductName = ""
    WindowsRegisteredOrganization = ""
    WindowsRegisteredOwner = ""
    WindowsSystemRoot = ""
}
$InfoProccesorRecord = [Ordered]@{

}
$InfoCountryRecord = [Ordered]@{
    OsCountryCode = ""
    OsCurrentTimeZone = ""
    OsLocaleID = ""
    OsLocale = ""
    OsLocalDateTime = ""
    OsLastBootUpTime = ""
    Timezone = ""
    OsOrganization = ""
    OsArchitecture = ""
    OsLanguage = ""
}
<#
$InfoBIOSRecord = [Ordered]@{
BiosCharacteristics                                     : {4, 7, 8, 9...}
BiosBIOSVersion                                         : {INTEL  - 6040000, PhoenixBIOS 4.0 Release 6.0     }
BiosBuildNumber                                         : 
BiosCaption                                             : PhoenixBIOS 4.0 Release 6.0     
BiosCodeSet                                             : 
BiosCurrentLanguage                                     : 
BiosDescription                                         : PhoenixBIOS 4.0 Release 6.0     
BiosEmbeddedControllerMajorVersion                      : 0
BiosEmbeddedControllerMinorVersion                      : 0
BiosFirmwareType                                        : Bios
BiosIdentificationCode                                  : 
BiosInstallableLanguages                                : 
BiosInstallDate                                         : 
BiosLanguageEdition                                     : 
BiosListOfLanguages                                     : 
BiosManufacturer                                        : Phoenix Technologies LTD
BiosName                                                : PhoenixBIOS 4.0 Release 6.0     
BiosOtherTargetOS                                       : 
BiosPrimaryBIOS                                         : True
BiosReleaseDate                                         : 12-12-2018 01:00:00
BiosSeralNumber                                         : VMware-42 0f ee d6 51 36 02 2d-1a 97 e9 96 7c 9c db 2c
BiosSMBIOSBIOSVersion                                   : 6.00
BiosSMBIOSMajorVersion                                  : 2
BiosSMBIOSMinorVersion                                  : 7
BiosSMBIOSPresent                                       : True
BiosSoftwareElementState                                : Running
BiosStatus                                              : OK
BiosSystemBiosMajorVersion                              : 4
BiosSystemBiosMinorVersion                              : 6
BiosTargetOperatingSystem                               : 0
BiosVersion                                             : INTEL  - 6040000

}
#>
$InfoIPConfigRecord = [Ordered]@{
    InterfaceAlias = ""
    InterfaceIndex = ""
    InterfaceDescription = ""
    NetProfileName = ""
    IPv4Address = ""
    IPv6DefaultGateway = ""
    IPv4DefaultGateway = ""
    DNSServer = ""
}
Function Read_Netadapter{
    $Global:InfoNetAdapterData = @()
    Foreach ($Interface in Get-NetIpAddress -InterfaceIndex $($NetAdapter.ifIndex)){
        $Data = New-Object -TypeName PSObject -Property $InfoNetAdapterRecord
        $Data.IPAddress = $Interface.IPAddress
        $Data.InterfaceIndex = $Interface.InterfaceIndex
        $Data.InterfaceAlias = $Interface.InterfaceAlias
        $Data.AddressFamily = $Interface.AddressFamily
        $Data.PrefixLength = $Interface.PrefixLength
        $Global:InfoNetAdapterData += $Data
    }
}
Function Read_InfoProfile{
    $Global:InfoProfileData = @()
    Foreach ($NetProfile in Get-NetConnectionProfile){
        $Data = New-Object -TypeName PSObject -Property $InfoProfileRecord
        $Data.Name = $NetProfile.Name
        $Data.InterfaceAlias = $NetProfile.InterfaceAlias
        $Data.InterfaceIndex  = $NetProfile.InterfaceIndex 
        $Data.NetworkCategory = $NetProfile.NetworkCategory
        $Data.IPv4Connectivity = $NetProfile.IPv4Connectivity
        $Data.IPv6Connectivity = $NetProfile.IPv6Connectivity
        $Global:InfoProfileData += $Data
    }
}
Function Read_InfoProfile1{
    $Global:InfoNetAdapterData = @()
    Foreach ($Volumes in Get-Volume | Where {$_.Driveletter -ne $Null -and $_.Drivetype -eq "Fixed"}){
        $Data = New-Object -TypeName PSObject -Property $InfoNetAdapterRecord
    }
}

$InfoIPConfigRecord1 = [Ordered]@{
}
Function Read_InfoIPConfig{
    $Global:InfoIPConfigData = @()
    Foreach ($IPConfig in Get-NetIPConfiguration){
        $Data = New-Object -TypeName PSObject -Property $InfoIPConfigRecord
        $Data.InterfaceAlias = $IPConfig.InterfaceAlias
        $Data.InterfaceIndex = $IPConfig.InterfaceIndex
        $Data.InterfaceDescription = $IPConfig.InterfaceDescription
        $Data.NetProfileName = $IPConfig.NetProfile.Name
        $Data.IPv4Address = $IPConfig.IPv4Address
        $Data.IPv6DefaultGateway = $IPConfig.IPv6DefaultGateway.nexthop
        $Data.IPv4DefaultGateway = $IPConfig.IPv4DefaultGateway.nexthop
        $Data.DNSServer = $IPConfig.DNSServer.Serveraddresses
        $Global:InfoIPConfigData += $Data
    }
}

Function Read_InfoVolumes{
    $Global:InfoVolumesData = @()
    Foreach ($Volume in Get-Volume | Where {$_.Driveletter -ne $Null -and $_.Drivetype -eq "Fixed"}){
        $Data = New-Object -TypeName PSObject -Property $InfoVolumeRecord
        $Data.DriveType = $Volume.DriveType
        $Data.FileSystemType = $Volume.FileSystemType
        $Data.AllocationUnitSize = $Volume.AllocationUnitSize
        $Data.DriveLetter = $Volume.DriveLetter
        $Data.FileSystem = $Volume.FileSystem
        $Data.FileSystemLabel = $Volume.FileSystemLabel
        $Data.Size = [Math]::Ceiling($Volume.Size/1gb)
        $Data.SizeRemaining = [Math]::Ceiling($Volume.SizeRemaining/1gb)
        $Global:InfoVolumesData += $Data
    }
}

Read_InfoVolumes
Read_Netadapter
Read_InfoProfile
Read_InfoIPConfig
#$Global:InfoNetAdapterData
#$Global:InfoProfileData
#$Global:InfoVolumesData
$Global:InfoIPConfigData
return
$Computerinfo = Get-computerinfo
$DiskInfo = Get-disk
$VolumeInfo = Get-Volume
$NetAdapter = Get-NetAdapter

Get-NetConnectionProfile
#Get-NetFirewallProfile
#Get-NetFirewallRule
Get-NetIpAddress -AddressFamily IPv4 -InterfaceIndex $($NetAdapter.ifIndex)
Get-NetIpAddress -InterfaceIndex $($NetAdapter.ifIndex) | Select IPaddress, InterfaceIndex, InterfaceAlias, AddressFamily

Get-NetIPConfiguration |Select  Netprofile.name, Ipv4address, ipv4Defaultgateway, DNSserver
#Get-NetTCPConnection
Get-hotfix
<#
Get-service | Select name,DisplayName | Export-csv $StandardWindowsServiceFile -Delimiter "," -NoTypeInformation
notepad $StandardWindowsServiceFile
#>
$StandardWindowsService = (get-content $StandardWindowsServiceFile)#.Name
$StandardWindowsService = (import-csv $StandardWindowsServiceFile -Delimiter ",")#.Name
$LocalServices = Get-service | where {$_.status -eq "running"}
Get-service | where {$_.status -eq "running" -and $_.name -notin $StandardWindowsService.name}
Write-host "WindowsEditionID : $($Computerinfo.WindowsEditionID)"
Write-host "WindowsInstallationType: $($Computerinfo.WindowsInstallationType)"
Write-host "WindowsInstalldateFromRegistry : $($Computerinfo.WindowsInstalldateFromRegistry)"
Write-host "WindowProductname : $($Computerinfo.WindowsProductname)"
Write-host "BiosBiosVersion : $($Computerinfo.BiosBiosVersion)"
Write-host "BIOSeralNumber : $($Computerinfo.BIOSseralNumber)"
Write-host "CsCaption : $($Computerinfo.CsCaption)"
Write-host "CsDomain : $($Computerinfo.CsDomain)"
Write-host "CsDomainRole : $($Computerinfo.CsDomainRole)"
Write-host "CsManufacturer : $($Computerinfo.CsManufacturer)"
Write-host "CsModel : $($Computerinfo.CsModel)"
Write-host "CsNetworkAdapters : $($Computerinfo.CsNetworkAdapters)"
Write-host "CsNetworkServerModeEnabled : $($Computerinfo.CsNetworkServerModeEnabled)"
Write-host "CsPartOfDomain : $($Computerinfo.CsPartOfDomain)"
Write-host "CsPrimaryOwnerName : $($Computerinfo.CsPrimaryOwnerName)"
Write-host "OsName : $($Computerinfo.OsName)"
Write-host "OsType : $($Computerinfo.OsType)"
Write-host "OsVersion : $($Computerinfo.OsVersion)"
Write-host "OsCountryCode : $($Computerinfo.OsCountryCode)"
Write-host "OsCurrentTimezone : $($Computerinfo.OsCurrentTimezone)"
Write-host "OsLocaleID : $($Computerinfo.OsLocaleID)"
Write-host "OsLocale : $($Computerinfo.OsLocale)"
Write-host "OsLastBootupTime : $($Computerinfo.OsLastBootupTime)"
Write-host "OsUptime : $($Computerinfo.OsUptime)"
Write-host "OsTotalVisibleMemorySize : $($Computerinfo.OsTotalVisibleMemorySize)"
Write-host "OsFreePhysicalMemory : $($Computerinfo.OsFreePhysicalMemory)"
Write-host "OsTotalVirtualMemorySize : $($Computerinfo.OsTotalVirtualMemorySize)"
Write-host "OsFreeVirtualMemory : $($Computerinfo.OsFreeVirtualMemory)"
Write-host "OsInUseVirtualMemory : $($Computerinfo.OsInUseVirtualMemory)"
Write-host "OsInstallDate : $($Computerinfo.OsInstallDate)"
Write-host "OsMuiLanguages : $($Computerinfo.OsMuiLanguages)"
Write-host "OsNumberOfUsers : $($Computerinfo.OsNumberOfUsers)"
Write-host "OsOrganization : $($Computerinfo.OsOrganization)"
Write-host "OsArchitecture : $($Computerinfo.OsArchitecture)"
Write-host "OsLanguage : $($Computerinfo.OsLanguage)"
Write-host "OsProductType : $($Computerinfo.OsProductType)"
Write-host "OsSerialNumber : $($Computerinfo.OsSerialNumber)"
Write-host "OsServerLevel : $($Computerinfo.OsServerLevel)"
Write-host "KeyboardLayout : $($Computerinfo.KeyboardLayout)"
Write-host "TimeZone : $($Computerinfo.TimeZone)"