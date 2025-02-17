<#
    $psISE.CurrentFile.Editor.ToggleOutliningExpansion()
#>

Install-Module -Name Az -AllowClobber -Scope CurrentUser
$Credentials = Get-Credential -Message "Geef het wachtwoord op" -UserName wiltensa@proveiling.com
Connect-AzAccount -Credential $Credentials
#region Record defenitions
$RecordSubscriptions = [Ordered]@{
    Name = ""
    Id = ""
    State = ""
}
$RecordVM = [Ordered]@{
    Name = ""
    Type = ""
    Location = ""
    HardwareProfile = ""
    Cores = ""
    Memory = ""
    OSDiskSize  = ""
    ResourceDiskSize = ""
    NetworkProfile = ""
    OSname = ""
    OsVersion = ""
    OSProfile = ""
    TimeCreated = ""
    StatusCode = ""
    Resourcegroep = ""
    Abonnement = ""
    Grootte = ""
    OutsideIPadres = ""
    InsideIPadres = ""
    VirtualNetwork = ""
    DNSname = ""
    LoadBalancer = ""
    Schijf = ""

}
$RecordsVirtualNetwork = [Ordered]@{
    Name = ""
    Subnets = ""
    ProvisioningState = ""
    EnableDdosProtection = ""
    Location = $VirtualNetwork.Location
    PrivateEndpointVNetPolicies = ""
    AddressSpaceText = ""
    ResourceGroupName = ""
    Type = ""
    ResourceGuid = ""
}
$RecordAZDomain = [Ordered]@{
    ID = ""
    TenantID = ""
    TenantCategory = ""
    CountryCode = ""
    Name = ""
    Domains = ""
    DefaultDomain = ""
}

$RecordApplicationInsights = [Ordered]@{
    Name = ""
    ApplicationId = ""
    ApplicationType = ""
    Etag = ""
    FlowType = ""
    Id = ""
    Kind = ""
    Location = ""
    PublicNetworkAccessForIngestion = ""
    PublicNetworkAccessForQuery = ""
    RetentionInDay = ""
    TenantId = ""
    Type = ""
}
$RecordAzAutoscaleSetting = [Ordered]@{
    Name = ""
    Location = "" 
    ID = ""
    Propertiesname = ""
    Profile = ""
}
$RecordAvailabilitySet = [Ordered]@{
    Name = ""
    ResourceGroupName = ""
    ID = ""
    Type = ""
    Location = ""
    VirtualMachinesReferences = ""
    VirtualMachines = @()
}
$RecordAzDisk = [Ordered]@{
    Name = ""
    ID = ""
    ResourceGroupName = ""
    OsType = ""
    HyperVGeneration = ""
    DiskSizeGB = ""
    DiskState = ""
    Type = ""
    Location = ""
    NetworkAccessPolicy = ""
    PublicNetworkAccess = ""
}
$RecordAzDnsZone = [Ordered]@{
    Name = ""
    ResourceGroupName = ""
    Etag = ""
    Tags = ""
    NameServers = ""
    ZoneType = ""
    RegistrationVirtualNetworkIds = ""
    ResolutionVirtualNetworkIds = ""
}
$RecordAzImage = [Ordered]@{
    ResourceGroupName = ""
    SourceVirtualMachine = ""
    StorageProfile = ""
    ProvisioningState = ""
    HyperVGeneration = ""
    Id = ""
    Name = ""
    Type = ""
    Location = ""
    Tags = ""
}
$RecordAzLoadBalancer = [Ordered]@{
    ResourceId = ""
    VaultName = ""
    ResourceGroupName = ""
    Location = ""
    Tags = ""
}
$RecordAzNetworkInterface = [Ordered]@{
    Name = ""
    VirtualMachine = ""
    IpConfigurations = ""
    DnsSettings = ""
    Id = ""
    Location = ""
    MacAddress = ""
    Primary = ""
    NetworkSecurityGroup = ""
    EnableIPForwarding = ""
    ProvisioningState = ""
    ResourceGroupName = ""
    Type = ""
    Tag = ""
    Etag = ""
    PrivateIpAddressVersion = ""
    GatewayLoadBalancer = ""
    PrivateIpAddress = ""
    PrivateIpAllocationMethod = ""
    SubnetName = ""
    PublicIpAddressName = ""
}
$RecordAzNetworkSecurityGroup = [Ordered]@{
    FlushConnection = ""
    SecurityRules = ""
    DefaultSecurityRules = ""
    NetworkInterfaces = ""
    Subnets = ""
    ProvisioningState = ""
    ResourceGroupName = ""
    ResourceGuid = ""
    Location = ""
    Type = ""
    Name = ""
    Id = ""
    Tag = ""
    Etag = ""
}
$RecordAzNetworkUsage = [Ordered]@{
	CurrentValue = ""
 	Limit = ""
 	Name = ""
 	ResourceType = ""
 	Unit = ""
 	Location = ""

}
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
$RecordAzPublicIpAddress = [Ordered]@{
	DdosSettings = ""
 	DdosSettingsText = ""
 	DnsSettings = ""
 	DnsSettingsText = ""
 	Etag = ""
 	ExtendedLocation = ""
 	Id = ""
 	IdleTimeoutInMinutes = ""
 	IpAddress = ""
 	IpConfiguration = ""
 	IpTagsText = ""
 	Location = ""
 	Name = ""
 	ProvisioningState = ""
 	PublicIpAddressVersion = ""
 	PublicIpAllocationMethod = ""
 	PublicIpPrefix = ""
 	ResourceGroupName = ""
 	ResourceGuid = ""
 	Sku = ""
 	Tag = ""
 	Type = ""
 	Zones = ""

}
$RecordAzResource = [Ordered]@{
	ChangedTime = ""
 	CreatedTime = ""
 	ETag = ""
 	ExtensionResourceName = ""
 	ExtensionResourceType = ""
 	Id = ""
 	Identity = ""
 	Kind = ""
 	Location = ""
 	ManagedBy = ""
 	Name = ""
 	ParentResource = ""
 	Plan = ""
 	Properties = ""
 	ResourceGroupName = ""
 	ResourceId = ""
 	ResourceName = ""
 	ResourceType = ""
 	Sku = ""
 	SubscriptionId = ""
 	Tags = ""
 	TagsTable = ""
 	Type = ""

}
$RecordAzResourceGroup = [Ordered]@{
	Location = ""
 	ManagedBy = ""
 	ProvisioningState = ""
 	ResourceGroupName = ""
 	ResourceId = ""
 	Tags = ""

}
$RecordAzRoleAssignment = [Ordered]@{
	CanDelegate = ""
 	Condition = ""
 	ConditionVersion = ""
 	Description = ""
 	DisplayName = ""
 	ObjectId = ""
 	ObjectType = ""
 	RoleAssignmentId = ""
 	RoleAssignmentName = ""
 	RoleDefinitionId = ""
 	RoleDefinitionName = ""
 	Scope = ""
 	SignInName = ""

}
$RecordAzRoleDefinition = [Ordered]@{
	Actions = ""
 	AssignableScopes = ""
 	Condition = ""
 	ConditionVersion = ""
 	DataActions = ""
 	Description = ""
 	Id = ""
 	IsCustom = ""
 	Name = ""
 	NotActions = ""
 	NotDataActions = ""

}
$RecordAzSecuritySecureScore = [Ordered]@{
	CurrentScore = ""
 	DisplayName = ""
 	Id = ""
 	MaxScore = ""
 	Name = ""
 	Percentage = ""
 	Type = ""
 	Weight = ""

}
$RecordAzSecuritySecureScoreControl = [Ordered]@{
	CurrentScore = ""
 	DisplayName = ""
 	HealthyResourceCount = ""
 	Id = ""
 	MaxScore = ""
 	Name = ""
 	NotApplicableResourceCount = ""
 	Percentage = ""
 	Type = ""
 	UnhealthyResourceCount = ""
 	Weight = ""

}
$RecordAzSecuritySecureScoreControlDefinition = [Ordered]@{
	AssessmentDefinitions = ""
 	Description = ""
 	DisplayName = ""
 	Id = ""
 	MaxScore = ""
 	Name = ""
 	Source = ""
 	Type = ""

}
$RecordAzStorageAccount = [Ordered]@{
	AccessTier = ""
 	AllowBlobPublicAccess = ""
 	AllowCrossTenantReplication = ""
 	AllowedCopyScope = ""
 	AllowSharedKeyAccess = ""
 	AzureFilesIdentityBasedAuth = ""
 	BlobRestoreStatus = ""
 	Context = ""
 	CreationTime = ""
 	CustomDomain = ""
 	DnsEndpointType = ""
 	EnableHierarchicalNamespace = ""
 	EnableHttpsTrafficOnly = ""
 	EnableLocalUser = ""
 	EnableNfsV3 = ""
 	EnableSftp = ""
 	Encryption = ""
 	ExtendedLocation = ""
 	ExtendedProperties = ""
 	FailoverInProgress = ""
 	GeoReplicationStats = ""
 	Id = ""
 	Identity = ""
 	ImmutableStorageWithVersioning = ""
 	KeyCreationTime = ""
 	KeyPolicy = ""
 	Kind = ""
 	LargeFileSharesState = ""
 	LastGeoFailoverTime = ""
 	Location = ""
 	MinimumTlsVersion = ""
 	NetworkRuleSet = ""
 	PrimaryEndpoints = ""
 	PrimaryLocation = ""
 	ProvisioningState = ""
 	PublicNetworkAccess = ""
 	ResourceGroupName = ""
 	RoutingPreference = ""
 	SasPolicy = ""
 	SecondaryEndpoints = ""
 	SecondaryLocation = ""
 	Sku = ""
 	StatusOfPrimary = ""
 	StatusOfSecondary = ""
 	StorageAccountName = ""
 	StorageAccountSkuConversionStatus = ""
 	Tags = ""

}
$RecordAzSubscription = [Ordered]@{
	AuthorizationSource = ""
 	CurrentStorageAccount = ""
 	CurrentStorageAccountName = ""
 	ExtendedProperties = ""
 	HomeTenantId = ""
 	Id = ""
 	ManagedByTenantIds = ""
 	Name = ""
 	State = ""
 	SubscriptionId = ""
 	Tags = ""
 	TenantId = ""

}
$RecordAzTenant = [Ordered]@{
	Country = ""
 	CountryCode = ""
 	DefaultDomain = ""
 	Domains = ""
 	ExtendedProperties = ""
 	Id = ""
 	Name = ""
 	TenantBrandingLogoUrl = ""
 	TenantCategory = ""
 	TenantId = ""
 	TenantType = ""

}
$RecordAzVirtualNetwork = [Ordered]@{
	AddressSpace = ""
 	BgpCommunities = ""
 	DdosProtectionPlan = ""
 	DhcpOptions = ""
 	EnableDdosProtection = ""
 	Encryption = ""
 	Etag = ""
 	ExtendedLocation = ""
 	FlowTimeoutInMinutes = ""
 	Id = ""
 	IpAllocations = ""
 	Location = ""
 	Name = ""
 	PrivateEndpointVNetPolicies = ""
 	ProvisioningState = ""
 	ResourceGroupName = ""
 	ResourceGuid = ""
 	Subnets = ""
 	Tag = ""
 	TagsTable = ""
 	Type = ""
 	VirtualNetworkPeerings = ""

}
$RecordAzVMUsage = [Ordered]@{
	CurrentValue = ""
 	Limit = ""
 	Name = ""
 	RequestId = ""
 	StatusCode = ""
 	Unit = ""

}
$RecordAzWebApp = [Ordered]@{
	AvailabilityState = ""
 	AzureStorageAccounts = ""
 	AzureStoragePath = ""
 	ClientAffinityEnabled = ""
 	ClientCertEnabled = ""
 	ClientCertExclusionPaths = ""
 	ClientCertMode = ""
 	CloningInfo = ""
 	ContainerSize = ""
 	CustomDomainVerificationId = ""
 	DailyMemoryTimeQuota = ""
 	DefaultHostName = ""
 	Enabled = ""
 	EnabledHostNames = ""
 	ExtendedLocation = ""
 	GitRemoteName = ""
 	GitRemotePassword = ""
 	GitRemoteUri = ""
 	GitRemoteUsername = ""
 	HostingEnvironmentProfile = ""
 	HostNames = ""
 	HostNamesDisabled = ""
 	HostNameSslStates = ""
 	HttpsOnly = ""
 	HyperV = ""
 	Id = ""
 	Identity = ""
 	InProgressOperationId = ""
 	IsDefaultContainer = ""
 	IsXenon = ""
 	KeyVaultReferenceIdentity = ""
 	Kind = ""
 	LastModifiedTimeUtc = ""
 	Location = ""
 	MaxNumberOfWorkers = ""
 	Name = ""
 	OutboundIpAddresses = ""
 	PossibleOutboundIpAddresses = ""
 	RedundancyMode = ""
 	RepositorySiteName = ""
 	Reserved = ""
 	ResourceGroup = ""
 	ScmSiteAlsoStopped = ""
 	ServerFarmId = ""
 	SiteConfig = ""
 	SlotSwapStatus = ""
 	State = ""
 	StorageAccountRequired = ""
 	SuspendedTill = ""
 	Tags = ""
 	TargetSwapSlot = ""
 	TrafficManagerHostNames = ""
 	Type = ""
 	UsageState = ""
 	VirtualNetworkSubnetId = ""
 	VnetInfo = ""

}
$RecordAzADUser = [Ordered]@{
	AccountEnabled = ""
 	AdditionalProperties = ""
 	AgeGroup = ""
 	ApproximateLastSignInDateTime = ""
 	BusinessPhone = ""
 	City = ""
 	CompanyName = ""
 	ComplianceExpirationDateTime = ""
 	ConsentProvidedForMinor = ""
 	Country = ""
 	CreatedDateTime = ""
 	CreationType = ""
 	DeletedDateTime = ""
 	Department = ""
 	DeviceVersion = ""
 	DisplayName = ""
 	EmployeeHireDate = ""
 	EmployeeId = ""
 	EmployeeOrgData = ""
 	EmployeeType = ""
 	ExternalUserState = ""
 	ExternalUserStateChangeDateTime = ""
 	FaxNumber = ""
 	GivenName = ""
 	Id = ""
 	Identity = ""
 	ImAddress = ""
 	IsResourceAccount = ""
 	JobTitle = ""
 	LastPasswordChangeDateTime = ""
 	LegalAgeGroupClassification = ""
 	Mail = ""
 	MailNickname = ""
 	Manager = ""
 	MobilePhone = ""
 	OdataId = ""
 	OdataType = ""
 	OfficeLocation = ""
 	OnPremisesImmutableId = ""
 	OnPremisesLastSyncDateTime = ""
 	OnPremisesSyncEnabled = ""
 	OperatingSystem = ""
 	OperatingSystemVersion = ""
 	OtherMail = ""
 	PasswordPolicy = ""
 	PasswordProfile = ""
 	PhysicalId = ""
 	PostalCode = ""
 	PreferredLanguage = ""
 	ProxyAddress = ""
 	ResourceGroupName = ""
 	ShowInAddressList = ""
 	SignInSessionsValidFromDateTime = ""
 	State = ""
 	StreetAddress = ""
 	Surname = ""
 	TrustType = ""
 	UsageLocation = ""
 	UserPrincipalName = ""
 	UserType = ""

}

#endregion Record defenitions

#region Functions
Function Inventariseer_SubScriptions{
    $SubScriptions = Get-AzSubscription
    $Global:SubScriptionsRecords = @()
    Foreach ($SubScription in $SubScriptions){
        $Data = New-Object -TypeName PSObject -Property $RecordSubscriptions
        $Data.Name = $Subscription.Name
        $Data.Id = $Subscription.Id
        $Data.State = $Subscription.State
        $Global:SubScriptionsRecords += $Data
    }

}
Function Inventariseer_VirtualNetwork{
    $FunctionName = ""
    $Global:VirtualNetworkRecords = @()
    $VirtualNetworks = Get-AzVirtualNetwork
    Foreach ($VirtualNetwork in $VirtualNetworks){
        $Data = New-Object -TypeName PSObject -Property $RecordsVirtualNetwork
        $Data.Name = $VirtualNetwork.Name
        $Data.Subnets = $VirtualNetwork.Subnets
        $Data.ProvisioningState = $VirtualNetwork.ProvisioningState 
        $Data.EnableDdosProtection = $VirtualNetwork.EnableDdosProtection
        $Data.Location = $VirtualNetwork.Location
        $Data.PrivateEndpointVNetPolicies = $VirtualNetwork.PrivateEndpointVNetPolicies
        $Data.AddressSpaceText = $VirtualNetwork.AddressSpaceText
        $Data.ResourceGroupName = $VirtualNetwork.ResourceGroupName 
        $Data.Type = $VirtualNetwork.Type
        $Data.ResourceGuid = $VirtualNetwork.ResourceGuid
        $Global:VirtualNetworkRecords += $Data
    }
}
Function Inventariseer_VirtualMachines{
    $Global:VirtualMachineRecords = @()
    $VirtualMachines = get-azvm -Name NAD-WEB-1 | Select *
    $VirtualMachine = get-azvm -Name NAD-WEB-1 | Select *
    Foreach ($VirtualMachine in $VirtualMachines){
        $Data = New-Object -TypeName PSObject -Property $RecordSubscriptions
        $Data.Name = $VirtualMachine.Name
        $Data.Type = $VirtualMachine.Type
        #$Data.Location = $VirtualMachine.Location
        $Data.Location = (Get-AzLocation | Where {$_.Location -eq $($VirtualMachine.Location)}).DisplayName
        $Data.HardwareProfile = $VirtualMachine.HardwareProfile.VmSize
        $HardwareProfile = get-azvmsize -Location $VirtualMachine.Location | Where {$_.name -eq $VirtualMachine.HardwareProfile.VmSize }
        $Data.Cores = $HardwareProfile.NumberOfCores
        $Data.Memory = $HardwareProfile.MemoryInMB
        $Data.OSDiskSize  = $HardwareProfile.OSDiskSizeInMB
        $Data.ResourceDiskSize = $HardwareProfile.ResourceDiskSizeInMB
        Cores
        Memory
        OSDiskSize 
        ResourceDiskSize

        
        $Data.NetworkProfile = $VirtualMachine.NetworkProfile
        $Data.NetworkProfile = (Get-AzNetworkInterface | where {$_.name -eq $VirtualMachine.name}).MacAddress

        $NI = Get-AzNetworkInterface | where {$_.id -eq $vm.NetworkProfile.NetworkInterfaces.id}| Select name, DnsSettings , IpConfigurations,MacAddress

        FrontendIpConfigurations     : {pvwebContractContract}
BackendAddressPools          : {HTTPS-443, WEB-80}
LoadBalancingRules           : {HTTPS-443, WEB-80}
Probes                       : {HTTPS-443, WEB-80}
InboundNatRules              : {PowerShell-NAD-WEB-1, RemoteDesktop-NAD-WEB-1, PowerShell-NAD-WEB-2, RemoteDesktop-NAD-WEB-2...}
InboundNatPools              : {}



        $Data.OSname = $VirtualMachine.OSname
        $Data.OsVersion = $VirtualMachine.OsVersion
        $Data.OSProfile = $VirtualMachine.OSProfile
        $Data.TimeCreated = $VirtualMachine.TimeCreated
        $Data.StatusCode = $VirtualMachine.StatusCode
        $Data.Resourcegroep = $VirtualMachine.Resourcegroep
        $Data.Abonnement = $VirtualMachine.Abonnement
        $Data.Grootte = $VirtualMachine.Grootte
        $Data.OutsideIPadres = $VirtualMachine.OutsideIPadres
        $Data.InsideIPadres = $VirtualMachine.InsideIPadres
        $Data.VirtualNetwork = $VirtualMachine.VirtualNetwork
        $Data.DNSname = $VirtualMachine.DNSname
        $Data.LoadBalancer = $VirtualMachine.LoadBalancer
        $Data.Schijf = $VirtualMachine.Schijf
        $Global:VirtualMachineRecords += $Data
    }
}
Function Inventariseer_AzADUser{
    $Global:AzADUserRecords = @()
    $AzADUsers = Get-AzADUser
    Foreach ($AzADUser in $AzADUsers){
        $Data = New-Object -TypeName PSObject -Property $RecordAzADUser
		$Data.AccountEnabled = $AzADUser.AccountEnabled
 		$Data.AdditionalProperties = $AzADUser.AdditionalProperties
 		$Data.AgeGroup = $AzADUser.AgeGroup
 		$Data.ApproximateLastSignInDateTime = $AzADUser.ApproximateLastSignInDateTime
 		$Data.BusinessPhone = $AzADUser.BusinessPhone
 		$Data.City = $AzADUser.City
 		$Data.CompanyName = $AzADUser.CompanyName
 		$Data.ComplianceExpirationDateTime = $AzADUser.ComplianceExpirationDateTime
 		$Data.ConsentProvidedForMinor = $AzADUser.ConsentProvidedForMinor
 		$Data.Country = $AzADUser.Country
 		$Data.CreatedDateTime = $AzADUser.CreatedDateTime
 		$Data.CreationType = $AzADUser.CreationType
 		$Data.DeletedDateTime = $AzADUser.DeletedDateTime
 		$Data.Department = $AzADUser.Department
 		$Data.DeviceVersion = $AzADUser.DeviceVersion
 		$Data.DisplayName = $AzADUser.DisplayName
 		$Data.EmployeeHireDate = $AzADUser.EmployeeHireDate
 		$Data.EmployeeId = $AzADUser.EmployeeId
 		$Data.EmployeeOrgData = $AzADUser.EmployeeOrgData
 		$Data.EmployeeType = $AzADUser.EmployeeType
 		$Data.ExternalUserState = $AzADUser.ExternalUserState
 		$Data.ExternalUserStateChangeDateTime = $AzADUser.ExternalUserStateChangeDateTime
 		$Data.FaxNumber = $AzADUser.FaxNumber
 		$Data.GivenName = $AzADUser.GivenName
 		$Data.Id = $AzADUser.Id
 		$Data.Identity = $AzADUser.Identity
 		$Data.ImAddress = $AzADUser.ImAddress
 		$Data.IsResourceAccount = $AzADUser.IsResourceAccount
 		$Data.JobTitle = $AzADUser.JobTitle
 		$Data.LastPasswordChangeDateTime = $AzADUser.LastPasswordChangeDateTime
 		$Data.LegalAgeGroupClassification = $AzADUser.LegalAgeGroupClassification
 		$Data.Mail = $AzADUser.Mail
 		$Data.MailNickname = $AzADUser.MailNickname
 		$Data.Manager = $AzADUser.Manager
 		$Data.MobilePhone = $AzADUser.MobilePhone
 		$Data.OdataId = $AzADUser.OdataId
 		$Data.OdataType = $AzADUser.OdataType
 		$Data.OfficeLocation = $AzADUser.OfficeLocation
 		$Data.OnPremisesImmutableId = $AzADUser.OnPremisesImmutableId
 		$Data.OnPremisesLastSyncDateTime = $AzADUser.OnPremisesLastSyncDateTime
 		$Data.OnPremisesSyncEnabled = $AzADUser.OnPremisesSyncEnabled
 		$Data.OperatingSystem = $AzADUser.OperatingSystem
 		$Data.OperatingSystemVersion = $AzADUser.OperatingSystemVersion
 		$Data.OtherMail = $AzADUser.OtherMail
 		$Data.PasswordPolicy = $AzADUser.PasswordPolicy
 		$Data.PasswordProfile = $AzADUser.PasswordProfile
 		$Data.PhysicalId = $AzADUser.PhysicalId
 		$Data.PostalCode = $AzADUser.PostalCode
 		$Data.PreferredLanguage = $AzADUser.PreferredLanguage
 		$Data.ProxyAddress = $AzADUser.ProxyAddress
 		$Data.ResourceGroupName = $AzADUser.ResourceGroupName
 		$Data.ShowInAddressList = $AzADUser.ShowInAddressList
 		$Data.SignInSessionsValidFromDateTime = $AzADUser.SignInSessionsValidFromDateTime
 		$Data.State = $AzADUser.State
 		$Data.StreetAddress = $AzADUser.StreetAddress
 		$Data.Surname = $AzADUser.Surname
 		$Data.TrustType = $AzADUser.TrustType
 		$Data.UsageLocation = $AzADUser.UsageLocation
 		$Data.UserPrincipalName = $AzADUser.UserPrincipalName
 		$Data.UserType = $AzADUser.UserType

        $Global:AzADUserRecords += $Data  
    }
}

Function Inventariseer_AZDomain{
    $Global:AZDomainRecords = @()
    $AZDomain = Get-AzDomain | Select *
    Foreach ($Domain in $AZDomain){
        $Data = New-Object -TypeName PSObject -Property $RecordAZDomain
        $Data.ID= $Domain.ID
        $Data.TenantID= $Domain.TenantID
        $Data.TenantCategory= $Domain.TenantCategory
        $Data.CountryCode= $Domain.CountryCode
        $Data.Name= $Domain.Name
        $Data.Domains= $Domain.Domains
        $Data.DefaultDomain= $Domain.DefaultDomain
        $Global:AZDomainRecords += $Data  
    }
}
Function Inventariseer_ApplicationInsights{
    $Global:ApplicationInsightsRecords = @()
    $AzApplicationInsights = Get-AzApplicationInsights |Select *
    Foreach ($AzApplicationInsight in $AzApplicationInsights){
        $Data = New-Object -TypeName PSObject -Property $RecordApplicationInsights
        $Data.Name = $AzApplicationInsight.Name
        $Data.ApplicationId = $AzApplicationInsight.ApplicationId
        $Data.ApplicationType = $AzApplicationInsight.ApplicationType
        $Data.Etag = $AzApplicationInsight.Etag
        $Data.FlowType = $AzApplicationInsight.FlowType
        $Data.Id = $AzApplicationInsight.Id
        $Data.Kind = $AzApplicationInsight.Kind
        $Data.Location = $AzApplicationInsight.Location
        $Data.PublicNetworkAccessForIngestion = $AzApplicationInsight.PublicNetworkAccessForIngestion
        $Data.PublicNetworkAccessForQuery = $AzApplicationInsight.PublicNetworkAccessForQuery
        $Data.RetentionInDay = $AzApplicationInsight.RetentionInDay
        $Data.TenantId = $AzApplicationInsight.TenantId
        $Data.Type = $AzApplicationInsight.Type
        $Global:ApplicationInsightsRecords += $Data  


    }
}
Function Inventariseer_AzAutoscaleSetting{
    $Global:AzAutoscaleSettingRecords = @()
    $AzAutoscaleSettings = Get-AzAutoscaleSetting | Select Name, Location, ID, Propertiesname, Profile
    Foreach ($AzAutoscaleSetting in $AzAutoscaleSettings){
        $Data = New-Object -TypeName PSObject -Property $RecordAzAutoscaleSetting
        $Data.Name = $AzAutoscaleSetting.Name
        $Data.Location = $AzAutoscaleSetting.Location
        $Data.ID = $AzAutoscaleSetting.ID
        $Data.Propertiesname = $AzAutoscaleSetting.Propertiesname
        #$Data.Profile = $AzAutoscaleSetting.Profile
        $Global:AzAutoscaleSettingRecords += $Data  
    }
}
Function Inventariseer_AvailabilitySet{
    $Global:AvailabilitySetRecords = @()
    $AvailabilitySets = Get-AzAvailabilitySet | Select Name,ResourceGroupName,ID,Type,Location,VirtualMachinesReferences 
    Foreach ($AvailabilitySet in $AvailabilitySets){
        $Data = New-Object -TypeName PSObject -Property $RecordAvailabilitySet
        $Data.Name = $AvailabilitySet.Name
        $Data.ResourceGroupName = $AvailabilitySet.ResourceGroupName
        $Data.ID = $AvailabilitySet.ID
        $Data.Type = $AvailabilitySet.Type
        $Data.Location = $AvailabilitySet.Location
        $Data.VirtualMachinesReferences = $AvailabilitySet.VirtualMachinesReferences
        $Data.VirtualMachines = (($(Get-AzAvailabilitySet  | Select VirtualMachinesReferences -first 1) | Select -ExpandProperty VirtualMachinesReferences).id | Get-AzVM).Name
        $Global:AvailabilitySetRecords += $Data  
        
    }
}
Function Inventariseer_AzDisk{
    $Global:AzDiskRecords = @()
    $AzDisks = Get-AzDisk
    Foreach ($AzDisk in $AzDisks){
        $Data = New-Object -TypeName PSObject -Property $RecordAzDisk
        $Data.Name = $AzDisk.Name
        $Data.ID = $AzDisk.ID
        $Data.ResourceGroupName = $AzDisk.ResourceGroupName
        $Data.OsType = $AzDisk.OsType
        $Data.HyperVGeneration = $AzDisk.HyperVGeneration
        $Data.DiskSizeGB = $AzDisk.DiskSizeGB
        $Data.DiskState = $AzDisk.DiskState
        $Data.Type = $AzDisk.Type
        $Data.Location = $AzDisk.Location
        $Data.NetworkAccessPolicy = $AzDisk.NetworkAccessPolicy
        $Data.PublicNetworkAccess = $AzDisk.PublicNetworkAccess
        $Global:AzDiskRecords += $Data  
    }
}
Function Inventariseer_AzDnsZoneRecords{
    $Global:AzDnsZoneRecords = @()
    $AzDnsZones = Get-AzDnsZone
    Foreach ($AzDnsZone in $AzDnsZones){
        $Data = New-Object -TypeName PSObject -Property $RecordAzDnsZone
        $Data.Name = $AzDnsZone.Name
        $Data.ResourceGroupName = $AzDnsZone.ResourceGroupName
        $Data.Etag = $AzDnsZone.Etag
        $Data.Tags = $AzDnsZone.Tags
        $Data.NameServers = $AzDnsZone.NameServers
        $Data.ZoneType = $AzDnsZone.ZoneType
        $Data.RegistrationVirtualNetworkIds = $AzDnsZone.RegistrationVirtualNetworkIds
        $Data.ResolutionVirtualNetworkIds = $AzDnsZone.ResolutionVirtualNetworkIds
        $Global:AzDnsZoneRecords += $Data  
    }
}
Function Inventariseer_AzImageRecords{
    $Global:AzImageRecords = @()
    $AzImages = Get-AzImage
    Foreach ($AzImage in $AzImages){
        $Data = New-Object -TypeName PSObject -Property $RecordAzImage
        $Data.ResourceGroupName = $AzImage.ResourceGroupName
        $Data.SourceVirtualMachine = $AzImage.SourceVirtualMachine
        $Data.StorageProfile = $AzImage.StorageProfile
        $Data.ProvisioningState = $AzImage.ProvisioningState
        $Data.HyperVGeneration = $AzImage.HyperVGeneration
        $Data.Id = $AzImage.Id
        $Data.Name = $AzImage.Name
        $Data.Type = $AzImage.Type
        $Data.Location = $AzImage.Location
        $Data.Tags = $AzImage.Tags
        $Global:AzImageRecords += $Data  
    }
}
Function Inventariseer_AzKeyVault{
    $Global:AzKeyVaultRecords = @()
    $AzKeyVaults = get-AzKeyVault
    Foreach ($AzKeyVault in $AzKeyVaults){
        $Data = New-Object -TypeName PSObject -Property $RecordAzKeyVault
        $Global:AzKeyVaultRecords += $Data  
        $Data.ResourceId = $AzKeyVault.ResourceId
        $Data.VaultName = $AzKeyVault.VaultName
        $Data.ResourceGroupName = $AzKeyVault.ResourceGroupName
        $Data.Location = $AzKeyVault.Location
        $Data.Tags = $AzKeyVault.Tags
    }
}
Function Inventariseer_AzLoadBalancer{
    $Global:AzLoadBalancerRecords = @()
    $AzLoadBalancers = Get-AzLoadBalancer
    Foreach ($AzLoadBalancer in $AzLoadBalancers){
        $Data = New-Object -TypeName PSObject -Property $RecordAzLoadBalancer
        $Global:AzLoadBalancerRecords += $Data  
        $Data.Name = $AzLoadBalancer.Name
        $Data.ResourceGroupName = $AzLoadBalancer.ResourceGroupName
        $Data.Id = $AzLoadBalancer.Id
        $Data.Location = $AzLoadBalancer.Location
        $Data.ResourceGuid = $AzLoadBalancer.ResourceGuid
        $Data.Type = $AzLoadBalancer.Type
        $Data.LoadBalancingRulesText = $AzLoadBalancer.LoadBalancingRulesText
        $Data.FrontendIpConfigurations = $AzLoadBalancer.FrontendIpConfigurations
        $Data.LoadBalancingRules = $AzLoadBalancer.LoadBalancingRules
        $Data.InboundNatRules = $AzLoadBalancer.InboundNatRules
        $Data.OutboundRules = $AzLoadBalancer.OutboundRules
        $Data.Sku = $AzLoadBalancer.Sku
        $Data.BackendAddressPoolsText = $AzLoadBalancer.BackendAddressPoolsText
        $Data.BackendAddressPools = $AzLoadBalancer.BackendAddressPools
        $Data.Etag = $AzLoadBalancer.Etag
    }
}
Function Inventariseer_AzNetworkInterface{
    $Global:AzNetworkInterfaceRecords = @()
    $NetworkInterfaces = Get-AzNetworkInterface
    Foreach ($NetworkInterface in $NetworkInterfaces){
        $Data = New-Object -TypeName PSObject -Property $RecordAzNetworkInterface
        $InterfaceIpConfig = Get-AzNetworkInterfaceIpConfig -Name ipconfig1 -NetworkInterface $(Get-AzNetworkInterface -Name pv-mail911 -ResourceGroupName "pv-mail_group")

        $Global:AzNetworkInterfaceRecords += $Data  
        $Data.Name= $AzNetworkInterfaceRecords.Name
        $Data.VirtualMachine= $AzNetworkInterfaceRecords.VirtualMachine
        $Data.IpConfigurations= $AzNetworkInterfaceRecords.IpConfigurations
        $Data.DnsSettings= $AzNetworkInterfaceRecords.DnsSettings
        $Data.Id= $AzNetworkInterfaceRecords.Id
        $Data.Location= $AzNetworkInterfaceRecords.Location
        $Data.MacAddress= $AzNetworkInterfaceRecords.MacAddress
        $Data.Primary= $AzNetworkInterfaceRecords.Primary
        $Data.NetworkSecurityGroup= $AzNetworkInterfaceRecords.NetworkSecurityGroup
        $Data.EnableIPForwarding= $AzNetworkInterfaceRecords.EnableIPForwarding
        $Data.ProvisioningState= $AzNetworkInterfaceRecords.ProvisioningState
        $Data.ResourceGroupName= $AzNetworkInterfaceRecords.ResourceGroupName
        $Data.Type= $AzNetworkInterfaceRecords.Type
        $Data.Tag= $AzNetworkInterfaceRecords.Tag
        $Data.Etag= $AzNetworkInterfaceRecords.Etag
        $Data.PrivateIpAddressVersion = $InterfaceIpConfig.PrivateIpAddressVersion
        $Data.GatewayLoadBalancer = $InterfaceIpConfig.GatewayLoadBalancer
        $Data.Primary = $InterfaceIpConfig.Primary
        $Data.PrivateIpAddress = $InterfaceIpConfig.PrivateIpAddress
        $Data.PrivateIpAllocationMethod  = $InterfaceIpConfig.PrivateIpAllocationMethod
        $Data.SubnetName  = $InterfaceIpConfig.'Subnet Name'
        $Data.PublicIpAddressName = $InterfaceIpConfig.'PublicIpAddress Name'
        $Data.ProvisioningState = $InterfaceIpConfig.ProvisioningState
    }
}
Function Inventariseer_AzNetworkSecurityGroup{
    $Global:AzNetworkSecurityGroupRecords = @()
    $NetworkSecurityGroups = Get-AzNetworkSecurityGroup
    Foreach ($NetworkSecurityGroup in $NetworkSecurityGroups){
        $Data = New-Object -TypeName PSObject -Property $RecordAzNetworkSecurityGroup
        $Data.FlushConnection = $NetworkSecurityGroup.FlushConnection
        $Data.SecurityRules = $NetworkSecurityGroup.SecurityRules
        $Data.DefaultSecurityRules = $NetworkSecurityGroup.DefaultSecurityRules
        $Data.NetworkInterfaces = $NetworkSecurityGroup.NetworkInterfaces
        $Data.Subnets = $NetworkSecurityGroup.Subnets
        $Data.ProvisioningState = $NetworkSecurityGroup.ProvisioningState
        $Data.ResourceGroupName = $NetworkSecurityGroup.ResourceGroupName
        $Data.ResourceGuid = $NetworkSecurityGroup.ResourceGuid
        $Data.Location = $NetworkSecurityGroup.Location
        $Data.Type = $NetworkSecurityGroup.Type
        $Data.Name = $NetworkSecurityGroup.Name
        $Data.Id = $NetworkSecurityGroup.Id
        $Data.Tag = $NetworkSecurityGroup.Tag
        $Data.Etag = $NetworkSecurityGroup.Etag
        $Global:AzNetworkSecurityGroupRecords += $Data  
    }
}
Function Inventariseer_AzNetworkUsage{
    $Global:AzNetworkUsageRecords = @()
    $Locations = (Get-AzResourceGroup).location | Select -Unique
    Foreach ($Location in $Locations){
        $AzNetworkSecurityGroups = Get-AzNetworkUsage -Location $Location | Where {$_.Currentvalue -ne 0}
        Foreach ($AzNetworkSecurityGroup in $AzNetworkSecurityGroups){
            $Data = New-Object -TypeName PSObject -Property $RecordAzNetworkUsage
		    $Data.CurrentValue = $AzNetworkSecurityGroup.CurrentValue
 		    $Data.Limit = $AzNetworkSecurityGroup.Limit
 		    $Data.Name = $AzNetworkSecurityGroup.Name
 		    $Data.ResourceType = $AzNetworkSecurityGroup.ResourceType
 		    $Data.Unit = $AzNetworkSecurityGroup.Unit
 		    $Data.Location = $Location

            $Global:AzNetworkUsageRecords += $Data  
        }
    }
}
Function Inventariseer_AzNetworkWatcher{
    $Global:AzNetworkWatcherRecords = @()
    $AzNetworkWatchers = Get-AzNetworkWatcher
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
Function Inventariseer_AzPublicIpAddress{
    $Global:AzPublicIpAddressRecords = @()
    $AzPublicIpAddresss = Get-AzPublicIpAddress
    Foreach ($AzPublicIpAddress in $AzPublicIpAddresss){
        $Data = New-Object -TypeName PSObject -Property $RecordAzPublicIpAddress
		$Data.DdosSettings = $AzPublicIpAddress.DdosSettings
 		$Data.DdosSettingsText = $AzPublicIpAddress.DdosSettingsText
 		$Data.DnsSettings = $AzPublicIpAddress.DnsSettings
 		$Data.DnsSettingsText = $AzPublicIpAddress.DnsSettingsText
 		$Data.Etag = $AzPublicIpAddress.Etag
 		$Data.ExtendedLocation = $AzPublicIpAddress.ExtendedLocation
 		$Data.Id = $AzPublicIpAddress.Id
 		$Data.IdleTimeoutInMinutes = $AzPublicIpAddress.IdleTimeoutInMinutes
 		$Data.IpAddress = $AzPublicIpAddress.IpAddress
 		$Data.IpConfiguration = $AzPublicIpAddress.IpConfiguration
 		$Data.IpTagsText = $AzPublicIpAddress.IpTagsText
 		$Data.Location = $AzPublicIpAddress.Location
 		$Data.Name = $AzPublicIpAddress.Name
 		$Data.ProvisioningState = $AzPublicIpAddress.ProvisioningState
 		$Data.PublicIpAddressVersion = $AzPublicIpAddress.PublicIpAddressVersion
 		$Data.PublicIpAllocationMethod = $AzPublicIpAddress.PublicIpAllocationMethod
 		$Data.PublicIpPrefix = $AzPublicIpAddress.PublicIpPrefix
 		$Data.ResourceGroupName = $AzPublicIpAddress.ResourceGroupName
 		$Data.ResourceGuid = $AzPublicIpAddress.ResourceGuid
 		$Data.Sku = $AzPublicIpAddress.Sku
 		$Data.Tag = $AzPublicIpAddress.Tag
 		$Data.Type = $AzPublicIpAddress.Type
 		$Data.Zones = $AzPublicIpAddress.Zones

        $Global:AzPublicIpAddressRecords += $Data  
    }
}

Function Inventariseer_AzResource{
    $Global:AzResourceRecords = @()
    $AzResources = Get-AzResource
    Foreach ($AzResource in $AzResources){
        $Data = New-Object -TypeName PSObject -Property $RecordAzResource
		$Data.ChangedTime = $AzResource.ChangedTime
 		$Data.CreatedTime = $AzResource.CreatedTime
 		$Data.ETag = $AzResource.ETag
 		$Data.ExtensionResourceName = $AzResource.ExtensionResourceName
 		$Data.ExtensionResourceType = $AzResource.ExtensionResourceType
 		$Data.Id = $AzResource.Id
 		$Data.Identity = $AzResource.Identity
 		$Data.Kind = $AzResource.Kind
 		$Data.Location = $AzResource.Location
 		$Data.ManagedBy = $AzResource.ManagedBy
 		$Data.Name = $AzResource.Name
 		$Data.ParentResource = $AzResource.ParentResource
 		$Data.Plan = $AzResource.Plan
 		$Data.Properties = $AzResource.Properties
 		$Data.ResourceGroupName = $AzResource.ResourceGroupName
 		$Data.ResourceId = $AzResource.ResourceId
 		$Data.ResourceName = $AzResource.ResourceName
 		$Data.ResourceType = $AzResource.ResourceType
 		$Data.Sku = $AzResource.Sku
 		$Data.SubscriptionId = $AzResource.SubscriptionId
 		$Data.Tags = $AzResource.Tags
 		$Data.TagsTable = $AzResource.TagsTable
 		$Data.Type = $AzResource.Type

        $Global:AzResourceRecords += $Data  
    }
}
Function Inventariseer_AzResourceGroup{
    $Global:AzResourceGroupRecords = @()
    $AzResourceGroups = Get-AzResourceGroup
    Foreach ($AzResourceGroup in $AzResourceGroups){
        $Data = New-Object -TypeName PSObject -Property $RecordAzResourceGroup
		$Data.Location = $AzResourceGroup.Location
 		$Data.ManagedBy = $AzResourceGroup.ManagedBy
 		$Data.ProvisioningState = $AzResourceGroup.ProvisioningState
 		$Data.ResourceGroupName = $AzResourceGroup.ResourceGroupName
 		$Data.ResourceId = $AzResourceGroup.ResourceId
 		$Data.Tags = $AzResourceGroup.Tags

        $Global:AzResourceGroupRecords += $Data  
    }
}
Function Inventariseer_AzRoleAssignment{
    $Global:AzRoleAssignmentRecords = @()
    $AzRoleAssignments = Get-AzRoleAssignment
    Foreach ($AzRoleAssignment in $AzRoleAssignments){
        $Data = New-Object -TypeName PSObject -Property $RecordAzRoleAssignment
		$Data.CanDelegate = $AzRoleAssignment.CanDelegate
 		$Data.Condition = $AzRoleAssignment.Condition
 		$Data.ConditionVersion = $AzRoleAssignment.ConditionVersion
 		$Data.Description = $AzRoleAssignment.Description
 		$Data.DisplayName = $AzRoleAssignment.DisplayName
 		$Data.ObjectId = $AzRoleAssignment.ObjectId
 		$Data.ObjectType = $AzRoleAssignment.ObjectType
 		$Data.RoleAssignmentId = $AzRoleAssignment.RoleAssignmentId
 		$Data.RoleAssignmentName = $AzRoleAssignment.RoleAssignmentName
 		$Data.RoleDefinitionId = $AzRoleAssignment.RoleDefinitionId
 		$Data.RoleDefinitionName = $AzRoleAssignment.RoleDefinitionName
 		$Data.Scope = $AzRoleAssignment.Scope
 		$Data.SignInName = $AzRoleAssignment.SignInName

        $Global:AzRoleAssignmentRecords += $Data  
    }
}
Function Inventariseer_AzRoleDefinition{
    $Global:AzRoleDefinitionRecords = @()
    $AzRoleDefinitions = Get-AzRoleDefinition
    Foreach ($AzRoleDefinition in $AzRoleDefinitions){
        $Data = New-Object -TypeName PSObject -Property $RecordAzRoleDefinition
		$Data.Actions = $AzRoleDefinition.Actions
 		$Data.AssignableScopes = $AzRoleDefinition.AssignableScopes
 		$Data.Condition = $AzRoleDefinition.Condition
 		$Data.ConditionVersion = $AzRoleDefinition.ConditionVersion
 		$Data.DataActions = $AzRoleDefinition.DataActions
 		$Data.Description = $AzRoleDefinition.Description
 		$Data.Id = $AzRoleDefinition.Id
 		$Data.IsCustom = $AzRoleDefinition.IsCustom
 		$Data.Name = $AzRoleDefinition.Name
 		$Data.NotActions = $AzRoleDefinition.NotActions
 		$Data.NotDataActions = $AzRoleDefinition.NotDataActions

        $Global:AzRoleDefinitionRecords += $Data  
    }
}
Function Inventariseer_AzSecuritySecureScore{
    $Global:AzSecuritySecureScoreRecords = @()
    $AzSecuritySecureScores = Get-AzSecuritySecureScore
    Foreach ($AzSecuritySecureScore in $AzSecuritySecureScores){
        $Data = New-Object -TypeName PSObject -Property $RecordAzSecuritySecureScore
		$Data.CurrentScore = $AzSecuritySecureScore.CurrentScore
 		$Data.DisplayName = $AzSecuritySecureScore.DisplayName
 		$Data.Id = $AzSecuritySecureScore.Id
 		$Data.MaxScore = $AzSecuritySecureScore.MaxScore
 		$Data.Name = $AzSecuritySecureScore.Name
 		$Data.Percentage = $AzSecuritySecureScore.Percentage
 		$Data.Type = $AzSecuritySecureScore.Type
 		$Data.Weight = $AzSecuritySecureScore.Weight

        $Global:AzSecuritySecureScoreRecords += $Data  
    }
}
Function Inventariseer_AzSecuritySecureScore{
    $Global:AzSecuritySecureScoreRecords = @()
    $AzSecuritySecureScores = Get-AzSecuritySecureScore
    Foreach ($AzSecuritySecureScore in $AzSecuritySecureScores){
        $Data = New-Object -TypeName PSObject -Property $RecordAzSecuritySecureScore
		$Data.CurrentScore = $AzSecuritySecureScore.CurrentScore
 		$Data.DisplayName = $AzSecuritySecureScore.DisplayName
 		$Data.Id = $AzSecuritySecureScore.Id
 		$Data.MaxScore = $AzSecuritySecureScore.MaxScore
 		$Data.Name = $AzSecuritySecureScore.Name
 		$Data.Percentage = $AzSecuritySecureScore.Percentage
 		$Data.Type = $AzSecuritySecureScore.Type
 		$Data.Weight = $AzSecuritySecureScore.Weight

        $Global:AzSecuritySecureScoreRecords += $Data  
    }
}
Function Inventariseer_AzSecuritySecureScoreControl{
    $Global:AzSecuritySecureScoreControlRecords = @()
    $AzSecuritySecureScoreControls = Get-AzSecuritySecureScoreControl
    Foreach ($AzSecuritySecureScoreControl in $AzSecuritySecureScoreControls){
        $Data = New-Object -TypeName PSObject -Property $RecordAzSecuritySecureScoreControl
		$Data.CurrentScore = $AzSecuritySecureScoreControl.CurrentScore
 		$Data.DisplayName = $AzSecuritySecureScoreControl.DisplayName
 		$Data.HealthyResourceCount = $AzSecuritySecureScoreControl.HealthyResourceCount
 		$Data.Id = $AzSecuritySecureScoreControl.Id
 		$Data.MaxScore = $AzSecuritySecureScoreControl.MaxScore
 		$Data.Name = $AzSecuritySecureScoreControl.Name
 		$Data.NotApplicableResourceCount = $AzSecuritySecureScoreControl.NotApplicableResourceCount
 		$Data.Percentage = $AzSecuritySecureScoreControl.Percentage
 		$Data.Type = $AzSecuritySecureScoreControl.Type
 		$Data.UnhealthyResourceCount = $AzSecuritySecureScoreControl.UnhealthyResourceCount
 		$Data.Weight = $AzSecuritySecureScoreControl.Weight

        $Global:AzSecuritySecureScoreControlRecords += $Data  
    }
}
Function Inventariseer_AzSecuritySecureScoreControlDefinition{
    $Global:AzSecuritySecureScoreControlDefinitionRecords = @()
    $AzSecuritySecureScoreControlDefinitions = Get-AzSecuritySecureScoreControlDefinition
    Foreach ($AzSecuritySecureScoreControlDefinition in $AzSecuritySecureScoreControlDefinitions){
        $Data = New-Object -TypeName PSObject -Property $RecordAzSecuritySecureScoreControlDefinition
		$Data.AssessmentDefinitions = $AzSecuritySecureScoreControlDefinition.AssessmentDefinitions
 		$Data.Description = $AzSecuritySecureScoreControlDefinition.Description
 		$Data.DisplayName = $AzSecuritySecureScoreControlDefinition.DisplayName
 		$Data.Id = $AzSecuritySecureScoreControlDefinition.Id
 		$Data.MaxScore = $AzSecuritySecureScoreControlDefinition.MaxScore
 		$Data.Name = $AzSecuritySecureScoreControlDefinition.Name
 		$Data.Source = $AzSecuritySecureScoreControlDefinition.Source
 		$Data.Type = $AzSecuritySecureScoreControlDefinition.Type

        $Global:AzSecuritySecureScoreControlDefinitionRecords += $Data  
    }
}
Function Inventariseer_AzStorageAccount{
    $Global:AzStorageAccountRecords = @()
    $AzStorageAccounts = Get-AzStorageAccount
    Foreach ($AzStorageAccount in $AzStorageAccounts){
        $Data = New-Object -TypeName PSObject -Property $RecordAzStorageAccount
		$Data.AccessTier = $AzStorageAccount.AccessTier
 		$Data.AllowBlobPublicAccess = $AzStorageAccount.AllowBlobPublicAccess
 		$Data.AllowCrossTenantReplication = $AzStorageAccount.AllowCrossTenantReplication
 		$Data.AllowedCopyScope = $AzStorageAccount.AllowedCopyScope
 		$Data.AllowSharedKeyAccess = $AzStorageAccount.AllowSharedKeyAccess
 		$Data.AzureFilesIdentityBasedAuth = $AzStorageAccount.AzureFilesIdentityBasedAuth
 		$Data.BlobRestoreStatus = $AzStorageAccount.BlobRestoreStatus
 		$Data.Context = $AzStorageAccount.Context
 		$Data.CreationTime = $AzStorageAccount.CreationTime
 		$Data.CustomDomain = $AzStorageAccount.CustomDomain
 		$Data.DnsEndpointType = $AzStorageAccount.DnsEndpointType
 		$Data.EnableHierarchicalNamespace = $AzStorageAccount.EnableHierarchicalNamespace
 		$Data.EnableHttpsTrafficOnly = $AzStorageAccount.EnableHttpsTrafficOnly
 		$Data.EnableLocalUser = $AzStorageAccount.EnableLocalUser
 		$Data.EnableNfsV3 = $AzStorageAccount.EnableNfsV3
 		$Data.EnableSftp = $AzStorageAccount.EnableSftp
 		$Data.Encryption = $AzStorageAccount.Encryption
 		$Data.ExtendedLocation = $AzStorageAccount.ExtendedLocation
 		$Data.ExtendedProperties = $AzStorageAccount.ExtendedProperties
 		$Data.FailoverInProgress = $AzStorageAccount.FailoverInProgress
 		$Data.GeoReplicationStats = $AzStorageAccount.GeoReplicationStats
 		$Data.Id = $AzStorageAccount.Id
 		$Data.Identity = $AzStorageAccount.Identity
 		$Data.ImmutableStorageWithVersioning = $AzStorageAccount.ImmutableStorageWithVersioning
 		$Data.KeyCreationTime = $AzStorageAccount.KeyCreationTime
 		$Data.KeyPolicy = $AzStorageAccount.KeyPolicy
 		$Data.Kind = $AzStorageAccount.Kind
 		$Data.LargeFileSharesState = $AzStorageAccount.LargeFileSharesState
 		$Data.LastGeoFailoverTime = $AzStorageAccount.LastGeoFailoverTime
 		$Data.Location = $AzStorageAccount.Location
 		$Data.MinimumTlsVersion = $AzStorageAccount.MinimumTlsVersion
 		$Data.NetworkRuleSet = $AzStorageAccount.NetworkRuleSet
 		$Data.PrimaryEndpoints = $AzStorageAccount.PrimaryEndpoints
 		$Data.PrimaryLocation = $AzStorageAccount.PrimaryLocation
 		$Data.ProvisioningState = $AzStorageAccount.ProvisioningState
 		$Data.PublicNetworkAccess = $AzStorageAccount.PublicNetworkAccess
 		$Data.ResourceGroupName = $AzStorageAccount.ResourceGroupName
 		$Data.RoutingPreference = $AzStorageAccount.RoutingPreference
 		$Data.SasPolicy = $AzStorageAccount.SasPolicy
 		$Data.SecondaryEndpoints = $AzStorageAccount.SecondaryEndpoints
 		$Data.SecondaryLocation = $AzStorageAccount.SecondaryLocation
 		$Data.Sku = $AzStorageAccount.Sku
 		$Data.StatusOfPrimary = $AzStorageAccount.StatusOfPrimary
 		$Data.StatusOfSecondary = $AzStorageAccount.StatusOfSecondary
 		$Data.StorageAccountName = $AzStorageAccount.StorageAccountName
 		$Data.StorageAccountSkuConversionStatus = $AzStorageAccount.StorageAccountSkuConversionStatus
 		$Data.Tags = $AzStorageAccount.Tags

        $Global:AzStorageAccountRecords += $Data  
    }
}
Function Inventariseer_AzSubscription{
    $Global:AzSubscriptionRecords = @()
    $AzSubscriptions = Get-AzSubscription
    Foreach ($AzSubscription in $AzSubscriptions){
        $Data = New-Object -TypeName PSObject -Property $RecordAzSubscription
		$Data.AuthorizationSource = $AzSubscription.AuthorizationSource
 		$Data.CurrentStorageAccount = $AzSubscription.CurrentStorageAccount
 		$Data.CurrentStorageAccountName = $AzSubscription.CurrentStorageAccountName
 		$Data.ExtendedProperties = $AzSubscription.ExtendedProperties
 		$Data.HomeTenantId = $AzSubscription.HomeTenantId
 		$Data.Id = $AzSubscription.Id
 		$Data.ManagedByTenantIds = $AzSubscription.ManagedByTenantIds
 		$Data.Name = $AzSubscription.Name
 		$Data.State = $AzSubscription.State
 		$Data.SubscriptionId = $AzSubscription.SubscriptionId
 		$Data.Tags = $AzSubscription.Tags
 		$Data.TenantId = $AzSubscription.TenantId

        $Global:AzSubscriptionRecords += $Data  
    }
}
Function Inventariseer_AzTenant{
    $Global:AzTenantRecords = @()
    $AzTenants = Get-AzTenant
    Foreach ($AzTenant in $AzTenants){
        $Data = New-Object -TypeName PSObject -Property $RecordAzTenant
		$Data.Country = $AzTenant.Country
 		$Data.CountryCode = $AzTenant.CountryCode
 		$Data.DefaultDomain = $AzTenant.DefaultDomain
 		$Data.Domains = $AzTenant.Domains
 		$Data.ExtendedProperties = $AzTenant.ExtendedProperties
 		$Data.Id = $AzTenant.Id
 		$Data.Name = $AzTenant.Name
 		$Data.TenantBrandingLogoUrl = $AzTenant.TenantBrandingLogoUrl
 		$Data.TenantCategory = $AzTenant.TenantCategory
 		$Data.TenantId = $AzTenant.TenantId
 		$Data.TenantType = $AzTenant.TenantType

        $Global:AzTenantRecords += $Data  
    }
}
Function Inventariseer_AzVirtualNetwork{
    $Global:AzVirtualNetworkRecords = @()
    $AzVirtualNetworks = Get-AzVirtualNetwork
    Foreach ($AzVirtualNetwork in $AzVirtualNetworks){
        $Data = New-Object -TypeName PSObject -Property $RecordAzVirtualNetwork
		$Data.AddressSpace = $AzVirtualNetwork.AddressSpace
 		$Data.BgpCommunities = $AzVirtualNetwork.BgpCommunities
 		$Data.DdosProtectionPlan = $AzVirtualNetwork.DdosProtectionPlan
 		$Data.DhcpOptions = $AzVirtualNetwork.DhcpOptions
 		$Data.EnableDdosProtection = $AzVirtualNetwork.EnableDdosProtection
 		$Data.Encryption = $AzVirtualNetwork.Encryption
 		$Data.Etag = $AzVirtualNetwork.Etag
 		$Data.ExtendedLocation = $AzVirtualNetwork.ExtendedLocation
 		$Data.FlowTimeoutInMinutes = $AzVirtualNetwork.FlowTimeoutInMinutes
 		$Data.Id = $AzVirtualNetwork.Id
 		$Data.IpAllocations = $AzVirtualNetwork.IpAllocations
 		$Data.Location = $AzVirtualNetwork.Location
 		$Data.Name = $AzVirtualNetwork.Name
 		$Data.PrivateEndpointVNetPolicies = $AzVirtualNetwork.PrivateEndpointVNetPolicies
 		$Data.ProvisioningState = $AzVirtualNetwork.ProvisioningState
 		$Data.ResourceGroupName = $AzVirtualNetwork.ResourceGroupName
 		$Data.ResourceGuid = $AzVirtualNetwork.ResourceGuid
 		$Data.Subnets = $AzVirtualNetwork.Subnets
 		$Data.Tag = $AzVirtualNetwork.Tag
 		$Data.TagsTable = $AzVirtualNetwork.TagsTable
 		$Data.Type = $AzVirtualNetwork.Type
 		$Data.VirtualNetworkPeerings = $AzVirtualNetwork.VirtualNetworkPeerings

        $Global:AzVirtualNetworkRecords += $Data  
    }
}
Function Inventariseer_AzVMUsage{
    $Global:AzVMUsageRecords = @()
    $Locations = (Get-AzResourceGroup).location | Select -Unique
    Foreach ($Location in $Locations){
        $AzVMUsages = Get-AzVMUsage -Location $Location | Where {$_.Currentvalue -ne 0}
        Foreach ($AzVMUsage in $AzVMUsages){
            $Data = New-Object -TypeName PSObject -Property $RecordAzVMUsage
		    $Data.CurrentValue = $AzVMUsage.CurrentValue
 		    $Data.Limit = $AzVMUsage.Limit
 		    $Data.Name = $AzVMUsage.Name
 		    $Data.RequestId = $AzVMUsage.RequestId
 		    $Data.StatusCode = $AzVMUsage.StatusCode
 		    $Data.Unit = $AzVMUsage.Unit

            $Global:AzVMUsageRecords += $Data  
        }
    }
}
Function Inventariseer_AzWebApp{
    $Global:AzWebAppRecords = @()
    $AzWebApps = Get-AzWebApp
    Foreach ($AzWebApp in $AzWebApps){
        $Data = New-Object -TypeName PSObject -Property $RecordAzWebApp
		$Data.AvailabilityState = $AzWebApp.AvailabilityState
 		$Data.AzureStorageAccounts = $AzWebApp.AzureStorageAccounts
 		$Data.AzureStoragePath = $AzWebApp.AzureStoragePath
 		$Data.ClientAffinityEnabled = $AzWebApp.ClientAffinityEnabled
 		$Data.ClientCertEnabled = $AzWebApp.ClientCertEnabled
 		$Data.ClientCertExclusionPaths = $AzWebApp.ClientCertExclusionPaths
 		$Data.ClientCertMode = $AzWebApp.ClientCertMode
 		$Data.CloningInfo = $AzWebApp.CloningInfo
 		$Data.ContainerSize = $AzWebApp.ContainerSize
 		$Data.CustomDomainVerificationId = $AzWebApp.CustomDomainVerificationId
 		$Data.DailyMemoryTimeQuota = $AzWebApp.DailyMemoryTimeQuota
 		$Data.DefaultHostName = $AzWebApp.DefaultHostName
 		$Data.Enabled = $AzWebApp.Enabled
 		$Data.EnabledHostNames = $AzWebApp.EnabledHostNames
 		$Data.ExtendedLocation = $AzWebApp.ExtendedLocation
 		$Data.GitRemoteName = $AzWebApp.GitRemoteName
 		$Data.GitRemotePassword = $AzWebApp.GitRemotePassword
 		$Data.GitRemoteUri = $AzWebApp.GitRemoteUri
 		$Data.GitRemoteUsername = $AzWebApp.GitRemoteUsername
 		$Data.HostingEnvironmentProfile = $AzWebApp.HostingEnvironmentProfile
 		$Data.HostNames = $AzWebApp.HostNames
 		$Data.HostNamesDisabled = $AzWebApp.HostNamesDisabled
 		$Data.HostNameSslStates = $AzWebApp.HostNameSslStates
 		$Data.HttpsOnly = $AzWebApp.HttpsOnly
 		$Data.HyperV = $AzWebApp.HyperV
 		$Data.Id = $AzWebApp.Id
 		$Data.Identity = $AzWebApp.Identity
 		$Data.InProgressOperationId = $AzWebApp.InProgressOperationId
 		$Data.IsDefaultContainer = $AzWebApp.IsDefaultContainer
 		$Data.IsXenon = $AzWebApp.IsXenon
 		$Data.KeyVaultReferenceIdentity = $AzWebApp.KeyVaultReferenceIdentity
 		$Data.Kind = $AzWebApp.Kind
 		$Data.LastModifiedTimeUtc = $AzWebApp.LastModifiedTimeUtc
 		$Data.Location = $AzWebApp.Location
 		$Data.MaxNumberOfWorkers = $AzWebApp.MaxNumberOfWorkers
 		$Data.Name = $AzWebApp.Name
 		$Data.OutboundIpAddresses = $AzWebApp.OutboundIpAddresses
 		$Data.PossibleOutboundIpAddresses = $AzWebApp.PossibleOutboundIpAddresses
 		$Data.RedundancyMode = $AzWebApp.RedundancyMode
 		$Data.RepositorySiteName = $AzWebApp.RepositorySiteName
 		$Data.Reserved = $AzWebApp.Reserved
 		$Data.ResourceGroup = $AzWebApp.ResourceGroup
 		$Data.ScmSiteAlsoStopped = $AzWebApp.ScmSiteAlsoStopped
 		$Data.ServerFarmId = $AzWebApp.ServerFarmId
 		$Data.SiteConfig = $AzWebApp.SiteConfig
 		$Data.SlotSwapStatus = $AzWebApp.SlotSwapStatus
 		$Data.State = $AzWebApp.State
 		$Data.StorageAccountRequired = $AzWebApp.StorageAccountRequired
 		$Data.SuspendedTill = $AzWebApp.SuspendedTill
 		$Data.Tags = $AzWebApp.Tags
 		$Data.TargetSwapSlot = $AzWebApp.TargetSwapSlot
 		$Data.TrafficManagerHostNames = $AzWebApp.TrafficManagerHostNames
 		$Data.Type = $AzWebApp.Type
 		$Data.UsageState = $AzWebApp.UsageState
 		$Data.VirtualNetworkSubnetId = $AzWebApp.VirtualNetworkSubnetId
 		$Data.VnetInfo = $AzWebApp.VnetInfo

        $Global:AzWebAppRecords += $Data  
    }
}
#endregion Functions


Install-Module -Name az -AllowClobber -Scope CurrentUser
$Credentials = Get-Credential -UserName wiltensa@proveiling.onmicrosoft.com -Message "Geef het wachtwoord"
#Connect-AzAccount -Credential $Credentials -Tenantid 8b3c2a29-1816-4fc0-a3d2-aeab5a1fefd6 + Auctionhouse Infrastructure
#Connect-AzAccount -Credential $Credentials -Tenantid 8eb322ab7-eac7-4eb2-b615-dccece0aa212 + Metrics
Connect-AzAccount -Tenantid 8b3c2a29-1816-4fc0-a3d2-aeab5a1fefd6
Connect-AzAccount -Tenantid 8eb322ab7-eac7-4eb2-b615-dccece0aa212
#region Call Functions
Inventariseer_VirtualMachines
Inventariseer_AzKeyVault
Inventariseer_AzLoadBalancer
Inventariseer_SubScriptions
Inventariseer_VirtualNetwork
Inventariseer_AzResourceGroup
Inventariseer_AzPublicIpAddress
Inventariseer_ApplicationInsights
Inventariseer_AZDomain
Inventariseer_AzAutoscaleSetting
Inventariseer_AvailabilitySet
Inventariseer_AzDisk
Inventariseer_AzDnsZoneRecords
Inventariseer_AzImageRecords
Inventariseer_AzNetworkInterface
Inventariseer_AzNetworkSecurityGroup
Inventariseer_AzNetworkUsage
Inventariseer_AzNetworkWatcher
Inventariseer_AzResource
Inventariseer_AzRoleAssignment
Inventariseer_AzRoleDefinition
Inventariseer_AzSecuritySecureScore
Inventariseer_AzSecuritySecureScoreControl
Inventariseer_AzSecuritySecureScoreControlDefinition
Inventariseer_AzStorageAccount
Inventariseer_AzSubscription
Inventariseer_AzTenant
Inventariseer_AzVirtualNetwork
Inventariseer_AzVMUsage
Inventariseer_AzWebApp
Inventariseer_AzADUser
#endregion Call Functions
$Global:AzNetworkWatcherRecords | ft -AutoSize
$Global:AzNetworkUsageRecords | ft -AutoSize
$Global:AzPublicIpAddressRecords | ft -AutoSize
$Global:AzResourceRecords | ft -AutoSize
$Global:AzResourceGroupRecords | ft -AutoSize
$Global:AzRoleAssignmentRecords | ft -AutoSize
$Global:AzRoleDefinitionRecords | ft -AutoSize
$Global:AzSecuritySecureScoreRecords | ft -AutoSize
$Global:AzSecuritySecureScoreControlRecords | ft -AutoSize
$Global:AzSecuritySecureScoreControlDefinitionRecords | ft -AutoSize
$Global:AzStorageAccountRecords | ft -AutoSize
$Global:AzSubscriptionRecords | ft -AutoSize
$Global:AzTenantRecords | ft -AutoSize
$Global:AzVirtualNetworkRecords | Select ResourceGroupName, name, Location | ft -AutoSize
$Global:AzVMUsageRecords | ft -AutoSize
$Global:AzWebAppRecords | Select name, ResourceGroup, location, kind,rule | ft -AutoSize

$Global:SubScriptionsRecords
$Global:VirtualNetworkRecords
$Global:VirtualMachineRecords
$Global:AZDomainRecords
$Global:ApplicationInsightsRecords
$Global:AzAutoscaleSettingRecords
$Global:AvailabilitySetRecords
$Global:AzDiskRecords
$Global:AzDnsZoneRecords
$Global:AzImageRecords
$Global:AzKeyVaultRecords
$Global:AzLoadBalancerRecords
$Global:AzNetworkInterfaceRecords
$Global:AzNetworkSecurityGroupRecords





$Global:AzADUserRecords | Select DisplayName, Mail, UserPrincipalName






<#
    $psISE.CurrentFile.Editor.ToggleOutliningExpansion()
#>


$ResourceTypes = Get-AzResource | Select ResourceType -Unique

Get-AzSubscription | Out-GridView -PassThru 
Get-AzResource | Select-Object ResourceType, Name, Location #| Export-Csv -Path ./AllResources.csv -NoTypeInformation

# VM Inventory: List VMs and export their details
Get-AzVM | Select-Object Name, Location, HardwareProfile.VmSize #| Export-Csv -Path ./VMInventory.csv -NoTypeInformation

# Storage Accounts: List accounts and export their details
Get-AzStorageAccount | Select-Object StorageAccountName, Location, SkuName #| Export-Csv -Path ./StorageAccounts.csv -NoTypeInformation

# Network Resources: List VNets and export their details
Get-AzVirtualNetwork | Select-Object Name, Location, AddressSpace #| Export-Csv -Path ./VNetInventory.csv -NoTypeInformation
Get-AzADUser
Get-az
Get-azloadbalancer | Select * -first 1

Get-AzResource
Get-AzResourceGroup
$Vm = Get-AzVM -Name NAD-WEB-1 | Select *
$Vm.NetworkProfile.NetworkInterfaces
$Vm.HardwareProfile.VmSize
$Vm.AvailabilitySetReference
$Vm.DiagnosticsProfile
$Vm.OSProfile
$Vm.StorageProfile

