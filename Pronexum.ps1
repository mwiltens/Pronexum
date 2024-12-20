cls
$Computerinfo = Get-computerinfo
$DiskInfo = Get-disk
$VolumeInfo = Get-Volume
$NetAdapter = Get-NetAdapter
Get-NetConnectionProfile
Get-NetFirewallProfile
#Get-NetFirewallRule
Get-NetIpAddress -AddressFamily IPv4 -InterfaceIndex $($NetAdapter.ifIndex)
Get-NetIpAddress -InterfaceIndex $($NetAdapter.ifIndex) | Select IPaddress, InterfaceIndex, InterfaceAlias, AddressFamily

Get-NetIPConfiguration |Select * # Netprofile.name, Ipv4address, ipv4Defaultgateway, DNSserver
#Get-NetTCPConnection
Get-hotfix
#Get-services
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