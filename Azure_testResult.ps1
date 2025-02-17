cls
If ([string]::IsNullOrEmpty($Global:Locations)){
    $Global:Locations = Get-AzLocation
}
#$VM = Get-azvm -Name NAD-WEB-1 | Select *
$GLobal:Tabsetting = 40
Function Get_vm_InFo{
    Param ($VMname)
    $VM = Get-azvm -Name $VMname | Select *
    $Networkinterface  = Get-AzNetworkInterface -ResourceId $($vm.NetworkProfile.NetworkInterfaces.id)
    Write-host Computername : $vm.Name
    Write-host Resourcegroup : $vm.ResourceGroupName
    Write-host Interfacename : $Networkinterface.Name
    Write-host privateIpaddress : $Networkinterface.IpConfigurations.privateIpaddress
    Write-host Public Ipaddress : (Get-AzPublicIpAddress -ResourceGroupName $vm.ResourceGroupName).IpAddress


    #Write-host Besturingssysteem "---" -ForegroundColor Magenta
    Write-host Besturingssysteem : $vm.StorageProfile.ImageReference.Offer
    Write-host OS : $vm.StorageProfile.OsDisk.OsType
    Write-host Automatic Updates: $VM.OSProfile.WindowsConfiguration.EnableAutomaticUpdates
    Write-host VM generatie "---" -ForegroundColor Magenta
    If ($vm.HardwareProfile.VmSize -like "*ps_v5"){
        $VMArchitecture = "Arm64"
    }else{
        $VMArchitecture = "X64"
    }
    Write-host VM-architectuur $VMArchitecture
    $Nic = Get-AzNetworkInterface -ResourceId $vm.NetworkProfile.NetworkInterfaces[0].Id
    <#$nic.IpConfigurations
    $nic.EnableIPForwarding

    ProvisioningState            : Succeeded
    Sku                          : Microsoft.Azure.Commands.Network.Models.PSLoadBalancerSku
    FrontendIpConfigurations     : {pvwebContractContract}
    BackendAddressPools          : {HTTPS-443, WEB-80}
    LoadBalancingRules           : {HTTPS-443, WEB-80}
    Probes                       : {HTTPS-443, WEB-80}
    InboundNatRules              : {PowerShell-NAD-WEB-1, RemoteDesktop-NAD-WEB-1, PowerShell-NAD-WEB-2, RemoteDesktop-NAD-WEB-2...}
    InboundNatPools              : {}
    OutboundRules                : {}
    Ex
    #>

    Write-host Openbaar IP-adres (Get-AzPublicIpAddress -ResourceGroupName $vm.ResourceGroupName).IpAddress
    Write-host Load balancer : $(Get-AzLoadBalancer -ResourceGroupName $vm.ResourceGroupName).name
    Write-host Virtueel netwerk/subnet "---" -ForegroundColor Magenta
    Write-host DNS-naam : (Get-AzPublicIpAddress -ResourceGroupName $vm.ResourceGroupName).DnsSettings.Fqdn
    Write-host Grootte : $($vm.HardwareProfile.VmSize)
    $HardwareProfile = get-azvmsize -Location $vm.Location | Where {$_.name -eq $vm.HardwareProfile.VmSize }
    Write-host vCPUs : $HardwareProfile.NumberOfCores
    Write-host RAM-geheugen: $HardwareProfile.MemoryInMB

    Write-host Besturingssysteemschijf $HardwareProfile.OSDiskSizeInMB
    $HardwareProfile.ResourceDiskSizeInMB
    Write-host Beschikbaarheidsset : (Get-AzAvailabilitySet -ResourceGroupName $($vm.ResourceGroupName)).Name
    Write-host Extensies -ForegroundColor Yellow
    Foreach ($extension in Get-AzVMExtension -VMName $vm.name -ResourceGroupName $($vm.ResourceGroupName)){
        $extension.name
    }
}
Function Get_diskInfo{
    Get-AzDisk | Select * -First 1
}
Function Rules_info{
    Param ($Rules, $Fields,$RuleType,$Width,[Validateset ($True, $False)]$ShowRules = "True", $ExtraInfo)
    Write-host Info over : $RuleType
    $Koptekst = @()
    $I = 0 
    Foreach ($Field in $Fields){                
        $Koptekst += $(($Field).padright($($Width[$I])))
        $I++
    }
    
    Write-host $Koptekst
    
    If ($ShowRules.tolower() -eq "true" ){
        Foreach ($Rule in $Rules){
            $InfoRegel = @()
            $I=0
            Foreach ($Field in $Fields){
                <#
                If (!([string]::IsNullOrEmpty($extrainfo))){
                    $InfoRegel += "$($([String]$($Rule.$($Field))).PadRight($($Width[$I])))$($($extrainfo.Bronnaam).padright(15)) $($($extrainfo.BackendIpadres).padright(15))$($extrainfo.NetworkInterface) "
                }else{
                    $InfoRegel += "$($([String]$($Rule.$($Field))).PadRight($($Width[$I])))"
                }
                #>
                $InfoRegel += "$($([String]$($Rule.$($Field))).PadRight($($Width[$I])))"
                $I++
            }
            If (!([string]::IsNullOrEmpty($extrainfo))){
                $InfoRegel += "$($($extrainfo.Bronnaam).padright(15)) $($($extrainfo.BackendIpadres).padright(15))$($extrainfo.NetworkInterface) "
            }
            Write-Host $InfoRegel           
        }
    }
}
Function Get_Loadbalancer_Info{
    Param ($loadBalanceName)
    If (!($loadBalanceName)){
        $Global:Loadbalanceinfo = Get-AzLoadBalancer | Select * -first 1
    }Else{
        $Global:Loadbalanceinfo = Get-AzLoadBalancer -name $loadBalanceName | Select *
    }
    $Nic = Get-AzNetworkInterface -ResourceGroupName $Loadbalanceinfo.ResourceGroupName
    $VM = Get-AzNetworkInterface -ResourceGroupName $Loadbalanceinfo.ResourceGroupName 
    
    Write-host Name : $Loadbalanceinfo.name
    Write-host ResourceGroupName : $Loadbalanceinfo.ResourceGroupName
    Write-host Location : $Loadbalanceinfo.Location
    Write-host SKU : $Loadbalanceinfo.sku.Name
    Write-host Back-endpool : ($Loadbalanceinfo.BackendAddressPools | Measure-Object).count
    Write-host Taakverdelingsregel : ($Loadbalanceinfo.LoadBalancingRules | Measure-Object).count
    Write-host Nat-regels : ($Loadbalanceinfo.InboundNatRules | Measure-Object).count
    Write-host FrontendIpConfigurations  : ($Loadbalanceinfo.FrontendIpConfigurations.name)
    Write-host FrontendIpadres : $(Get-AzPublicIpAddress -ResourceGroupName $Loadbalanceinfo.ResourceGroupName).IpAddress 

    Write-host Network Interface : $($Nic.name)
    Write-host BackendIpadres : $Nic.IpConfigurations.privateipaddress
    Write-host Bronnaam : (Get-AzVM -ResourceId $($VM.VirtualMachine.id)).name
            $BackendInfo = [PSCustomObject]@{
                NetworkInterface = $($Nic.name)
                BackendIpadres  = $Nic.IpConfigurations.privateipaddress
                Bronnaam     = (Get-AzVM -ResourceId $($VM.VirtualMachine.id)).name
            }

    Write-host Back-end-Pools :  
    #Rules_info -Rules $Loadbalanceinfo.BackendAddressPools -Fields @("Name","ProvisioningState") -RuleType BackEndAddressPools -Width @(20) -ShowRules true
    Rules_info -Rules $Loadbalanceinfo.BackendAddressPools -Fields @("Name") -RuleType BackEndAddressPools -Width @(20) -ShowRules true -extrainfo $BackendInfo
    Write-host Taakverdelingsregels:  
    Rules_info -Rules $Loadbalanceinfo.LoadBalancingRules -RuleType LoadBalancingRules -Fields ("Name","Protocol","FrontendPort","BackendPort") -Width @(20,10,15,15) -ShowRules True
    Write-host Inkomende NAT-regels:  
    Rules_info -Rules $Loadbalanceinfo.inboundNatrules -RuleType inboundNatrules -Fields ("Protocol","Name","FrontendPort","BackendPort") -Width @(10,20,15,15) -ShowRules True  -extrainfo $BackendInfo
    #$result = Rules_info -Rules $Loadbalanceinfo.FrontendIpConfigurations  -RuleType FrontendIpConfigurations -fields @("Name") -Width @(20) -ShowRules False
    #$result.Count
    #$result = Rules_info -Rules $Loadbalanceinfo.inboundNatrules -RuleType inboundNatrules -Fields ("Protocol","Name","FrontendPort","BackendPort") -Width @(10,20,15,15) -ShowRules True
    #$result.Count
<#ProvisioningState            : Succeeded
Sku                          : Microsoft.Azure.Commands.Network.Models.PSLoadBalancerSku
FrontendIpConfigurations     : {pvdcdb1ContractContract}
BackendAddressPools          : {FTP20-20, FTP21-21, FTP7000-7000, FTP7001-7001...}
LoadBalancingRules           : {FTP20-20, FTP21-21, FTP7000-7000, FTP7001-7001...}
Probes                       : {}
InboundNatRules              : {FTP-7003-PV-DC-DB-1, FTP-7004-PV-DC-DB-1, FTP-7005-PV-DC-DB-1, FTP-7006-PV-DC-DB-1...}
InboundNatPools              : {}
OutboundRules                : {}
ExtendedLocation             : 
SkuText                      : {
#>

}
Function Get_Loadbalancer_Info1{
    Param ($loadBalanceName)
    If (!($loadBalanceName)){
        $Global:Loadbalanceinfo = Get-AzLoadBalancer | Select * -first 1
    }Else{
        $Global:Loadbalanceinfo = Get-AzLoadBalancer -name $loadBalanceName | Select *
    }
    Write-host Name : $Loadbalanceinfo.name
    Write-host ResourceGroupName : $Loadbalanceinfo.ResourceGroupName
    Write-host Location : $Loadbalanceinfo.Location
    $FrontendIpConfigurations = $Loadbalanceinfo | Select -ExpandProperty FrontendIpConfigurations
    
    $result = Rules_info -Rules $Loadbalanceinfo.FrontendIpConfigurations  -RuleType FrontendIpConfigurations -fields @("Name") -Width @(20) -ShowRules False
    $result.Count
    $result = Rules_info -Rules $Loadbalanceinfo.inboundNatrules -RuleType inboundNatrules -Fields ("Protocol","Name","FrontendPort","BackendPort") -Width @(10,20,15,15) -ShowRules True
    $result.Count
    Foreach ($FrontendIpConfiguration in $FrontendIpConfigurations){
        #($FrontendIpConfiguration | Get-member | where {$_.membertype -eq "Property" -and $_.name -notlike "*Text*"}).name | Set-Clipboard
        #$FrontendIpConfiguration | Select name, Frontendport,Backendport
        #Write-host FrontendIpConfiguration name : $FrontendIpConfiguration.name
        #InboundNatPools
        #InboundNatRules
        Write-host FrontendIpConfiguration name : $FrontendIpConfiguration.name
        #Write-host FrontendIpConfiguration InboundNatPools : $FrontendIpConfiguration | Select -ExpandProperty InboundNatPools
        #Write-host FrontendIpConfiguration InboundNatPools : $FrontendIpConfiguration.InboundNatPools
        #Write-host FrontendIpConfiguration InboundNatRules : $FrontendIpConfiguration.InboundNatRules
        #Write-host FrontendIpConfiguration InboundNatRules : $FrontendIpConfiguration | Select -ExpandProperty InboundNatRules
        #
        Write-host FrontendIpConfiguration LoadBalancingRules : 
        Rules_info -Rules $Loadbalanceinfo.LoadBalancingRules -RuleType LoadBalancingRules -Fields ("Protocol","Name","FrontendPort","BackendPort") -Width @(10,20,15,15) -ShowRules False
        Rules_info -Rules $Loadbalanceinfo.BackendAddressPools -Fields @("Name") -RuleType BackEndAddressPools -Width @(20) -ShowRules False
        Write-host FrontendIpConfiguration Name : $FrontendIpConfiguration.Name
        Write-host FrontendIpConfiguration OutboundRules : $FrontendIpConfiguration.OutboundRules
        Write-host FrontendIpConfiguration PrivateIpAddress : $FrontendIpConfiguration.PrivateIpAddress
        #Get-AzPublicIpAddress -ResourceId $frontend.PublicIpAddress.Id
        Write-host FrontendIpConfiguration PrivateIpAddressVersion : $FrontendIpConfiguration.PrivateIpAddressVersion
        Write-host FrontendIpConfiguration PrivateIpAllocationMethod : $FrontendIpConfiguration.PrivateIpAllocationMethod
        Write-host FrontendIpConfiguration ProvisioningState : $FrontendIpConfiguration.ProvisioningState
        Write-host FrontendIpConfiguration PublicIpAddress : $FrontendIpConfiguration.PublicIpAddress
        If (!([string]::IsNullOrEmpty($FrontendIpConfiguration.PublicIpAddress))){
            Write-host FrontendIpConfiguration PublicIpAddress : $(Get-AzPublicIpAddress -ResourceId $FrontendIpConfiguration.PublicIpAddress.Id) -ForegroundColor Yellow
        }
        Write-host FrontendIpConfiguration PublicIPPrefix : $FrontendIpConfiguration.PublicIPPrefix
        Write-host FrontendIpConfiguration Subnet : $FrontendIpConfiguration.Subnet
        Write-host FrontendIpConfiguration Zones : $FrontendIpConfiguration.Zones
        ##>
    }
}
Function Get_User_Info{
    get-azaduser | Select  DisplayName, Givenname,Surname, UserPrincipalname,  Usertype, Resourcegroupname, Prefferedlanguage, BusinessPhone  | ft
}
Function Get_Database_Info{
    Param ($Servername,[Validateset("SQL","Progress", "MySQL")]$Databasetype = "SQL")

    Switch ($Databasetype){
        "SQL" 
            {
                $CommandDBServer = "Get-AzSqlServer"
                $CommandFireWallrules = "Get-AzSqlServerFirewallRule"
            }
        "Progress"
            {
                $CommandDBServer = "Get-AzPostgreSqlServer"
                $CommandFireWallrules = "Get-AzPostgreSqlFirewallRule"
            }
        "MySQL"
            {
                $CommandDBServer = "Get-AzMySqlServer"
                $CommandFireWallrules = "Get-AzMySqlFirewallRule"                
            }
    
    }
    If (!($Servername)){
        $DatabaseServers = Invoke-Expression $CommandDBServer | Select * #-first 1
    }Else{
        #$DatabaseServers = Invoke-Expression $command -name $Servername | Select *
        $DatabaseServers = Invoke-Expression $CommandDBServer | Where {$_.Servername -eq $Servername} | Select *
    }
    #$DatabaseServers
    $Tab = 30
    Foreach ($DatabaseServer in $DatabaseServers){
        $Global:Titlecolor = "Yellow"
        #Write-host $($("ServerName").PadRight($Tab)) " : " $DatabaseServer.ServerName -ForegroundColor Yellow
        #Write-host $($("ResourceGroupName").PadRight($Tab)) " : " $DatabaseServer.ResourceGroupName
        #Write-host $($("Location").PadRight($Tab)) " : " $DatabaseServer.Location
        #Write-host $($("SqlAdministratorLogin").PadRight($Tab)) " : " $DatabaseServer.SqlAdministratorLogin
        #Write-host $($("ServerVersion").PadRight($Tab)) " : " $DatabaseServer.ServerVersion
        #Write-host $($("FullyQualifiedDomainName").PadRight($Tab)) " : " $DatabaseServer.FullyQualifiedDomainName
        #Write-host $($("ResourceId").PadRight($Tab)) " : " $DatabaseServer.ResourceId
        #Write-host $($("MinimalTlsVersion").PadRight($Tab)) " : " $DatabaseServer.MinimalTlsVersion
        #Write-host $($("PublicNetworkAccess").PadRight($Tab)) " : " $DatabaseServer.PublicNetworkAccess
        #Write-host $($("RestrictOutboundNetworkAccess").PadRight($Tab)) " : " $DatabaseServer.RestrictOutboundNetworkAccess
        #Write-host $($("Administrators").PadRight($Tab)) " : " $DatabaseServer.Administrators
        Display_Info -Title ServerName -Titlecolor $Global:Titlecolor -Field $DatabaseServer.ServerName -FieldColor $Global:FieldColor
        Display_Info -Title ResourceGroupName -Titlecolor $Global:Titlecolor -Field $DatabaseServer.ResourceGroupName -FieldColor $Global:FieldColor
        Display_Info -Title Locations -Titlecolor $Global:Titlecolor -Field $(Read_Location -Location $DatabaseServer.Location) -FieldColor $Global:FieldColor
        Display_Info -Title SqlAdministratorLogin -Titlecolor $Global:Titlecolor -Field $DatabaseServer.SqlAdministratorLogin -FieldColor $Global:FieldColor
        Display_Info -Title FullyQualifiedDomainName -Titlecolor $Global:Titlecolor -Field $DatabaseServer.FullyQualifiedDomainName -FieldColor $Global:FieldColor        
        Display_Info -Title ResourceId -Titlecolor $Global:Titlecolor -Field $DatabaseServer.ResourceId -FieldColor $Global:FieldColor
        Display_Info -Title MinimalTlsVersion -Titlecolor $Global:Titlecolor -Field $DatabaseServer.MinimalTlsVersion -FieldColor $Global:FieldColor
        Display_Info -Title PublicNetworkAccess -Titlecolor $Global:Titlecolor -Field $DatabaseServer.PublicNetworkAccess -FieldColor $Global:FieldColor
        Display_Info -Title RestrictOutboundNetworkAccess -Titlecolor $Global:Titlecolor -Field $DatabaseServer.RestrictOutboundNetworkAccess -FieldColor $Global:FieldColor
        Display_Info -Title Administrators -Titlecolor $Global:Titlecolor -Field $DatabaseServer.Administrators -FieldColor $Global:FieldColor
        
        Read_DatabaseServerRules -PSCommand $CommandFireWallrules -Server $DatabaseServer.ServerName -Resourcegroup $DatabaseServer.ResourceGroupName

        
        $Databases = Get-AzSqlDatabase -ServerName $DatabaseServer.ServerName -ResourceGroupName $DatabaseServer.ResourceGroupName

        #$Databases = $null
        $Global:Titlecolor = "Magenta"
        $Global:FieldColor = "Yellow"
        Foreach ($Database in $Databases){
            Display_Info -Title DatabaseName -Titlecolor $Global:Titlecolor -Field $Database.DatabaseName -FieldColor Green
            Display_Info -Title ResourceGroupName -Titlecolor $Global:Titlecolor -Field $Database.ResourceGroupName -FieldColor $Global:FieldColor
            Display_Info -Title ServerName -Titlecolor $Global:Titlecolor -Field $Database.ServerName -FieldColor $Global:FieldColor            
            Display_Info -Title Status -Titlecolor $Global:Titlecolor -Field $Database.Status -FieldColor $Global:FieldColor
            Display_Info -Title Location -Titlecolor $Global:Titlecolor -Field $(Read_Location -Location $Database.Location) -FieldColor $Global:FieldColor

            Display_Info -Title Edition -Titlecolor $Global:Titlecolor -Field $Database.Edition -FieldColor $Global:FieldColor
            Display_Info -Title CollationName -Titlecolor $Global:Titlecolor -Field $Database.CollationName -FieldColor $Global:FieldColor
            Display_Info -Title MaxSizeBytes -Titlecolor $Global:Titlecolor -Field "$($Database.MaxSizeBytes/1gb) GB" -FieldColor $Global:FieldColor
            
            Display_Info -Title CreationDate -Titlecolor $Global:Titlecolor -Field $Database.CreationDate -FieldColor $Global:FieldColor
            Display_Info -Title ResourceId -Titlecolor $Global:Titlecolor -Field $Database.ResourceId -FieldColor $Global:FieldColor
            Display_Info -Title CreateMode -Titlecolor $Global:Titlecolor -Field $Database.CreateMode -FieldColor $Global:FieldColor
            Display_Info -Title ReadScale -Titlecolor $Global:Titlecolor -Field $Database.ReadScale -FieldColor $Global:FieldColor
            Display_Info -Title ZoneRedundant -Titlecolor $Global:Titlecolor -Field $Database.ZoneRedundant -FieldColor $Global:FieldColor
            Display_Info -Title Capacity -Titlecolor $Global:Titlecolor -Field $Database.Capacity -FieldColor $Global:FieldColor
            Display_Info -Title Prijscategorie -Titlecolor $Global:Titlecolor -Field $Database.SkuName -FieldColor $Global:FieldColor
            Display_Info -Title CurrentBackupStorageRedundancy -Titlecolor $Global:Titlecolor -Field $Database.CurrentBackupStorageRedundancy -FieldColor $Global:FieldColor
            Display_Info -Title RequestedBackupStorageRedundancy -Titlecolor $Global:Titlecolor -Field $Database.RequestedBackupStorageRedundancy -FieldColor $Global:FieldColor
            Display_Info -Title MaintenanceConfigurationId -Titlecolor $Global:Titlecolor -Field $Database.MaintenanceConfigurationId -FieldColor $Global:FieldColor

        }
        #Get-AzPostgreSqlFirewallRule
        #Get-AzMySqlFirewallRule
        #Get-AzSqlDatabaseLongTermRetentionPolicy
        #Get-AzSqlInstanceTDEProtector


    }
    return
    Get-AzSqlDatabase -ServerName nadecontest -ResourceGroupName Database-Test
    Get-AzResource  | Where {$_.name -like "*nadecontest*"} | Select *
    Get-AzResource  | Where {$_.type -eq "Microsoft.Sql/servers/databases"}
    Get-AzMySqlServer
Get-AzOracleDbServer

Get-AzPostgreSqlServer
Get-AzPostgreSqlFirewallRule
Get-AzSqlServer | Select *
Get-AzSqlServerFirewallRule -ServerName nadecontest -ResourceGroupName Database-Test
}
Function Display_Info{
    param ($Tab = $GLobal:Tabsetting, $Title, [ValidateSet("Black","Blue" ,"Cyan" ,"DarkBlue","DarkCyan","DarkGray","DarkGreen","DarkMagenta","DarkRed","DarkYellow","Gray","Green","Magenta","White","Red","Yellow")]$Titlecolor = "White", $Field, [ValidateSet("Black","Blue" ,"Cyan" ,"DarkBlue","DarkCyan","DarkGray","DarkGreen","DarkMagenta","DarkRed","DarkYellow","Gray","Green","Magenta","White","Red","Yellow")]$FieldColor = "White")
    #Write-host $($("Administrators").PadRight($Tab)) " : " $DatabaseServer.Administrators
    write-host "$($($Title).PadRight($Tab)) : " -ForegroundColor $Titlecolor -NoNewline
    If (!([string]::IsNullOrEmpty($Field))){
        write-host $Field -ForegroundColor $FieldColor
    }else {Write-host}
}
Function Read_Location{
    Param ($Location)
    Return $($Global:locations | where {$_.location -eq $Location}).displayname
}
Function Read_DatabaseServerRules{
    Param ($PSCommand, $Server,$Resourcegroup)
    $Executecommand = "$Pscommand -servername $Server -ResourceGroupName $Resourcegroup"
    Return Invoke-Expression $Executecommand | Select Firewallrulename, StartIpaddress, EndIpaddress | Ft
    Return Invoke-Expression $Pscommand -ServerName $Server -ResourceGroupName $Resourcegroup
    $rules = Get-AzSqlServerFirewallRule -ServerName nadecontest -ResourceGroupName Database-Test
}
Function Get_Netwerk_info{
    param ($Networkname)
    If (!($Networkname)){
        $Networks = Get-AzVirtualNetwork | Select * #-first 1
    }Else{
        #$DatabaseServers = Invoke-Expression $command -name $Servername | Select *
        $Networks = Get-AzVirtualNetwork | Where {$_.name -eq $Networkname} | Select *
    }
    #$DatabaseServers
    $Tab = 30
    Foreach ($Network in $Networks){
        $Global:Titlecolor = "Yellow"
        Display_Info -Title Name -Titlecolor $Global:Titlecolor -Field $Network.Name -FieldColor $Global:FieldColor
        Display_Info -Title ResourceGroupName -Titlecolor $Global:Titlecolor -Field $Network.ResourceGroupName -FieldColor $Global:FieldColor
        Display_Info -Title Locations -Titlecolor $Global:Titlecolor -Field $(Read_Location -Location $Network.Location) -FieldColor $Global:FieldColor
        Display_Info -Title Adresruimte -Titlecolor $Global:Titlecolor -Field $($Network.AddressSpace.AddressPrefixes ) -FieldColor $Global:FieldColor
        Display_Info -Title DNSServers -Titlecolor $Global:Titlecolor  -Field $($Network.DhcpOptions.DnsServers) -FieldColor $Global:FieldColor
        Display_Info -Title Subnetten -Titlecolor $Global:Titlecolor
        Display_Info -Title DDOS-netwerkbeveiliging -Titlecolor $Global:Titlecolor
        Display_Info -Title Firewall
    }
    
}
<#
Benodige informatie:
Virtual Machines : Get_VM_Info
Netwerk interfaces

Loadbalancers : Get_Loadbalancer_Info
Abonnementen (2)
Users : Get_User_Info
Disk info : Get_diskInfo
Databases (1) : Get_Database_Info
Express Routes
gekoppelde abonnementen (1)
Resource groepen (62)
Automatische SChaalaanpassing
Netwerken    - DNS-zones (1)
             - Virtuele netwerken (3) : Get_Netwerk_info
             - Openbare Ip-adressen (11)
             - Virtuele Netwerkgateways (1 - DevVPN) 
             - NetwerkBeveilingsgroepen (4)
             - Load Balancers (6)

Storage      - opslagaccounts (26)
             - Schijven (10)
Web & Mobile - App Service-domeinen (1)
             - App Service-plannen (3)
             - Application Insights (5)
             - App Services (3)

#>
#Get_vm_InFo -VMname pv-mail
#Get_diskInfo
#Get_Loadbalancer_Info 
#Get_Database_Info -Servername nadecontest -Databasetype SQL
Get_Netwerk_info -Networkname pvintern
#Get_User_Info
<#
AddressSpaceText            : {
                                "AddressPrefixes": [
                                  "10.4.0.0/16"
                                ],
                                "IpamPoolPrefixAllocations": []
                              }
DhcpOptionsText             : {
                                "DnsServers": [
                                  "10.4.3.4",
                                  "10.4.3.20"
                                ]
                              }
#>