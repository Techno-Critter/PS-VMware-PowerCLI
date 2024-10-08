<#
Author: Stan Crider
Date: 7October2019
What this crap does:
Gather information from specified vCenter servers and outputs to Excel
### Must have at least read-access to vCenter
### Must have VMware.PowerCLI.VCenter module installed!!!
### Must have ImportExcel module installed!!!
### https://github.com/dfinke/ImportExcel
#>

#Requires -Modules ImportExcel, VMware.PowerCLI.VCenter

#region Configure variables
$Date = Get-Date -Format yyyyMMdd
# File name and folder location where spreadsheet will be created
$LogFile = "C:\Logs\vCenter\vCenter_Report_$Date.xlsx"
# vCenter servers
$VCServers = @("vcenter-1.acme.com","vcenter-2.acme.com")
# Logon credentials. NOTE: if login account changes a new credential file will be created!
$LoginAccount = "vcenterer@acme.com"
# Folder location where credential file will be created and stored; file name will be created automatically
# User must have write access to credential file directory!!!
$CredentialFileDirectory = "\\fileserver.acme.com\Credentials"
#endregion

#region Function: Change data sizes to legible values; converts number to string
Function Get-Size([double]$DataSize){
    Switch($DataSize){
        {$_ -lt 1KB}{
            $DataValue =  "$DataSize B"
        }
        {($_ -ge 1KB) -and ($_ -lt 1MB)}{
            $DataValue = "{0:N2}" -f ($DataSize/1KB) + " KiB"
        }
        {($_ -ge 1MB) -and ($_ -lt 1GB)}{
            $DataValue = "{0:N2}" -f ($DataSize/1MB) + " MiB"
        }
        {($_ -ge 1GB) -and ($_ -lt 1TB)}{
            $DataValue = "{0:N2}" -f ($DataSize/1GB) + " GiB"
        }
        {($_ -ge 1TB) -and ($_ -lt 1PB)}{
            $DataValue = "{0:N2}" -f ($DataSize/1TB) + " TiB"
        }
        Default{
            $DataValue = "{0:N2}" -f ($DataSize/1PB) + " PiB"
        }
    }
    $DataValue
}
#endregion

#region Function: Convert number of object items into Excel column headers
Function Get-ColumnName ([int]$ColumnCount){
    <#
    .SYNOPSIS
    Converts integer into Excel column headers

    .DESCRIPTION
    Takes a provided number of columns in a table and converts the number into Excel header format
    Input: 27 - Output: AA
    Input: 2 - Ouput: B

    .EXAMPLE
    Get-ColumnName 27

    .INPUTS
    Integer

    .OUTPUTS
    String

    .NOTES
    Author: Stan Crider and Dennis Magee
    #>

    If(($ColumnCount -le 702) -and ($ColumnCount -ge 1)){
        $ColumnCount = [Math]::Floor($ColumnCount)
        $CharStart = 64
        $FirstCharacter = $null

        # Convert number into double letter column name (AA-ZZ)
        If($ColumnCount -gt 26){
            $FirstNumber = [Math]::Floor(($ColumnCount)/26)
            $SecondNumber = ($ColumnCount) % 26

            # Reset increment for base-26
            If($SecondNumber -eq 0){
                $FirstNumber--
                $SecondNumber = 26
            }

            # Left-side column letter (first character from left to right)
            $FirstLetter = [int]($FirstNumber + $CharStart)
            $FirstCharacter = [char]$FirstLetter

            # Right-side column letter (second character from left to right)
            $SecondLetter = $SecondNumber + $CharStart
            $SecondCharacter = [char]$SecondLetter

            # Combine both letters into column name
            $CharacterOutput = $FirstCharacter + $SecondCharacter
        }

        # Convert number into single letter column name (A-Z)
        Else{
            $CharacterOutput = [char]($ColumnCount + $CharStart)
        }
    }
    Else{
        $CharacterOutput = "ZZ"
    }

    # Output column name
    $CharacterOutput
}
#endregion

#region Configure arrays and counters
$vCenterError = @()
$vCenterObject = @()
$DatacenterData = @()
$ClusterData = @()
$ClusterDRSData = @()
$VCLicenseServers = @()
$LicenseCustomObject = @()
$AssignedLicenseObject = @()
$HostData = @()
$HostNicData = @()
$HostVMKData = @()
$VMData = @()
$VMNicData = @()
$VMDriveData = @()
$VMHardDiskData = @()
$DatastoresData = @()
$SnapshotData = @()
$LicenseDataCounter = 0
$VCServerCounter = 0
#endregion

#region Credentials
# Set username and password for vCenter access. NOTE: account must have at least read access!
$Hostname = $ENV:COMPUTERNAME
$CurrentUser = $ENV:USERNAME #NOTE: user must have modify access to CredentialFileDirectory location!
$CredentialFile = "$CredentialFileDirectory\$Hostname\vCenter Creds $CurrentUser.xml"
If(Test-Path $CredentialFile){
    $Credentials = Import-Clixml $CredentialFile
}
Else{
    $Credentials = Get-Credential -UserName $LoginAccount -Message "Provide the password for $LoginAccount"
    If(-Not (Test-Path "$CredentialFileDirectory\$Hostname")){
        New-Item -Path $CredentialFileDirectory -Name $Hostname -ItemType Directory
    }
    $Credentials | Export-Clixml $CredentialFile
}
#endregion

# Connect to vCenters and retrieve data
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$False
ForEach($VCServer in $VCServers){
$VCServerCounter ++
    Write-Progress -Activity "vCenter server $VCServer" -Status ("Connecting to vCenter server " + $VCServerCounter + " of " + (($VCServers |Measure-Object).Count) + ".")
    Try{
        Connect-VIServer -Server $VCServer -Credential $Credentials
    }
    Catch{
        $vCenterError += [PSCustomObject]@{
            "Object" = "vCenter"
            "Name"   = $VCServer
            "Error"  = "The server $VCServer did not accept the connection request. This vCenter server will be skipped."
        }
        Continue
    }

    #region vCenter Servers
    $vCenterObject += [PSCustomObject]@{
        "Name"           = ($global:DefaultVIServer).Name
        "Port"           = ($global:DefaultVIServer).Port
        "Version"        = ($global:DefaultVIServer).Version
        "Build"          = ($global:DefaultVIServer).Build
        "Patch"          = ($global:DefaultVIServer).ExtensionData.Content.About.PatchLevel
        "OS Type"        = ($global:DefaultVIServer).ExtensionData.Content.About.OsType
        "Last Boot"      = ($global:DefaultVIServer).ExtensionData.ServerClock
        "Client Timeout" = ("" + ((($global:DefaultVIServer).ExtensionData.Client.ServiceTimeout)/60000) + " minute(s)")
    }
    #endregion

    #region Datacenters
    Try{
        $VDatacenters = Get-Datacenter -Server $VCServer -ErrorAction Stop
    }
    Catch{
        $VDatacenters = $null
    }

    Write-Progress -Activity "vCenter server $VCServer" -Status "Gathering Datacenter information..."
    If($VDatacenters){
        ForEach($VDatacenter in $VDatacenters){
            $DatacenterData += [PSCustomObject]@{
                "Datacenter" = $VDatacenter.Name
                "vCenter"    = $VCServer
                "Hosts"      = (Get-VMHost -Location $VDatacenter.Name | Measure-Object).Count
                "VMs"        = (Get-VM -Location $VDatacenter.Name | Measure-Object).Count
            }
        }
    }
    #endregion

    #region Clusters
    Try{
        $VClusters = Get-Cluster -Server $VCServer -ErrorAction Stop
    }
    Catch{
        $VClusters = $null
    }

    If($VClusters){
        Write-Progress -Activity "vCenter server $VCServer" -Status "Gathering Cluster information..."
        ForEach($VCluster in $VCLusters){
            $ClusterData += [PSCustomObject]@{
                "Cluster"    = $VCluster.Name
                "Datacenter" = (Get-Datacenter -Cluster $VCluster).Name
                "vCenter"    = $VCServer
                "HA"         = $VCluster.HAEnabled
                "DRS"        = $VCluster.DrsEnabled
                "EVC"        = $VCluster.ExtensionData.Summary.CurrentEVCModeKey
                "AutoLevel"  = $VCluster.DrsAutomationLevel
                "Hosts"      = (Get-VMHost -Location $VCluster.Name | Measure-Object).Count
                "VMs"        = (Get-VM -Location $VCluster.Name | Measure-Object).Count
            }
            #endregion

            #region ClusterRules
            # DRS Affinity Rules
            Try{
                $ClusterDRSRules = Get-DrsRule -Cluster $VCluster.Name -ErrorAction Stop
            }
            Catch{
                $ClusterDRSRules = $null
            }
            If($ClusterDRSRules){
                Write-Progress -Activity "vCenter server $VCServer" -Status "Gathering DRS rules information..."
                ForEach($DRSRule in $ClusterDRSRules){
                    $RuleMachineNames = @()
                    If($DRSRule.VMIDs){
                        $AppliedRuleMachines = $DRSRule.VMIDs
                        ForEach($RuleMachine in $AppliedRuleMachines){
                            $RuleMachineNames += Get-VM -Id $RuleMachine
                        }
                    }
                    $ClusterDRSData += [PSCustomObject]@{
                        "vCenter"    = $VCServer
                        "Datacenter" = (Get-Datacenter -Cluster $VCluster).Name
                        "Cluster"    = $VCluster.Name
                        "Rule"       = $DRSRule.Name
                        "Enabled"    = $DRSRule.Enabled
                        "Type"       = $DRSRule.Type
                        "VMs"        = $RuleMachineNames -join ", "
                        "Hosts"      = "N/A"
                    }
                }
            }
            # DRS VM/Host Rules
            Try{
                $ClusterDRSVMHostRules = Get-DrsVMHostRule -Cluster $VCluster.Name -ErrorAction Stop
            }
            Catch{
                $ClusterDRSVMHostRules = $null
            }
            If($ClusterDRSVMHostRules){
                Write-Progress -Activity "vCenter server $VCServer" -Status "Gather VM/Host rules information..."
                ForEach($VMHostRule in $ClusterDRSVMHostRules){
                    $ClusterDRSData += [PSCustomObject]@{
                        "vCenter"    = $VCServer
                        "Datacenter" = (Get-Datacenter -Cluster $VCluster).Name
                        "Cluster"    = $VCluster.Name
                        "Rule"       = $VMHostRule.Name
                        "Enabled"    = $VMHostRule.Enabled
                        "Type"       = $VMHostRule.Type
                        "VMs"        = $VMHostRule.VMGroup.Member.Name -join ", "
                        "Hosts"      = $VMHostRule.VMHostGroup.Member.Name -join ", "
                    }
                }
            }
        }
    }
    #endregion

    #region Licensing
    # Licensing only needs to be run once. Once complete, increment counter so license section doesn't run again.
    If($LicenseDataCounter -eq 0){
        Write-Progress -Activity "vCenter server $VCServer" -Status "Gathering licensing information..."
        $VCLicenseServers = Get-View LicenseManager
        # Enumerate license servers
        ForEach($VCLicenseServer in $VCLicenseServers){
            $LicenseCustomObjects = $VCLicenseServer.Licenses

            ForEach($LicenseObj in $LicenseCustomObjects){
                $LicenseProperties = $LicenseObj.Properties

                # License product and version handling
                $LicenseProduct = $LicenseProperties | Where-Object {$_.Key -eq 'ProductName'} | Select-Object -ExpandProperty Value
                $LicenseVersion = $LicenseProperties | Where-Object {$_.Key -eq 'ProductVersion'} | Select-Object -ExpandProperty Value

                # License expiration handling
                $LicenseExpiration = "Never"
                $LicenseExpiresValues = $LicenseProperties | Where-Object {$_.Key -eq 'ExpirationDate'} | Select-Object -ExpandProperty Value
                If($LicenseObj.Name -eq "Product Evaluation"){
                    $LicenseExpiration = "Evaluation"
                }
                ElseIf($LicenseExpiresValues){
                    $LicenseExpiration = $LicenseExpiresValues
                }

                # License count handling
                If($LicenseObj.Total -eq 0){
                    $LicenseCount = "Unlimited"
                }
                Else{
                    $LicenseCount = $LicenseObj.Total
                }

                $LicenseCustomObject += [PSCustomObject]@{
                    "vCenter Server" = $VCServer
                    "License Host"   = ([System.uri]$VCLicenseServer.Client.ServiceUrl).Host
                    "Name"           = $LicenseObj.Name
                    "Product"        = $LicenseProduct
                    "Version"        = $LicenseVersion
                    "Edition Key"    = $LicenseObj.EditionKey
                    "License Key"    = $LicenseObj.LicenseKey
                    "Total"          = $LicenseCount
                    "In Use"         = $LicenseObj.Used
                    "Units"          = $LicenseObj.CostUnit
                    "Expires"        = $LicenseExpiration
                    "Labels"         = $LicenseObj.Labels
                }
            }
            #endregion

            #region Assigned Licenses
            $AssignmentManager = Get-View $VCLicenseServer.LicenseAssignmentManager
            $AssignedLicenses = $null
            $AssignedLicenses = $AssignmentManager.QueryAssignedLicenses($VCLicenseServer.InstanceUUID)

            ForEach($AssignedLicense in $AssignedLicenses){
                $AssignedLicenseObject += [PSCustomObject]@{
                    "Entity"          = $AssignedLicense.EntityDisplayName
                    "License Name"    = $AssignedLicense.AssignedLicense.Name
                    "Product Name"    = $AssignedLicense.AssignedLicense.Properties | Where-Object {$_.Key -eq 'ProductName'} | Select-Object -ExpandProperty Value
                    "Product Version" = $AssignedLicense.AssignedLicense.Properties | Where-Object {$_.Key -eq 'ProductVersion'} | Select-Object -ExpandProperty Value
                    "License Key"     = $AssignedLicense.AssignedLicense.LicenseKey
                    "Edition Key"     = $AssignedLicense.AssignedLicense.EditionKey
                    "Scope"           = $AssignedLicense.Scope
                }
            }
            #endregion
        }
        $LicenseDataCounter ++
    }

    #region Hosts
    Try{
        $VMHosts = Get-VMHost -Server $VCServer -ErrorAction Stop
    }
    Catch{
        $VMHosts = $null
        $vCenterError += [PSCustomObject]@{
            "Object" = "vCenter"
            "Name"   = $VCServer
            "Error"  = "The Get-VMHost command failed on $VCServer"
        }
    }

    If($VMHosts){
        $HostCounter = 0
        ForEach($VMHost in $VMHosts){
            $ErrorCount = 0
            $HostCounter ++
            Write-Progress -Activity "vCenter server $VCServer" -Status "Gathering host information..." -CurrentOperation $VMHost.Name -PercentComplete ($HostCounter*100/($VMHosts | Measure-Object).Count)

            Try{
                $HostObjView = $VMHost | Get-View -ErrorAction Stop
            }
            Catch{
                $HostObjView = $null
                $ErrorCount ++
                $vCenterError += [PSCustomObject]@{
                    "Object" = "Host"
                    "Name"   = $VMHost.Name
                    "Error"  = "The Get-View command failed on host $($VMHost.Name)"
                }
            }

            Try{
                $NTPServers =  Get-VMHostNtpServer -VMHost $VMHost.Name -ErrorAction Stop
            }
            Catch{
                $NTPServers = "Error"
                $ErrorCount ++
                $vCenterError += [PSCustomObject]@{
                    "Object" = "Host"
                    "Name"   = $VMHost.Name
                    "Error"  = "The Get-VMHostNtpServer command failed on host $($VMHost.Name)"
                }
            }

            If($HostObjView.Hardware.SystemInfo.OtherIdentifyingInfo){
                    $HostSerialNumber = $HostObjView.Hardware.SystemInfo.OtherIdentifyingInfo[1].IdentifierValue
            }
            Else{
                $HostSerialNumber = "N/A"
            }

            Try{
                $ESXCli = Get-EsxCli -VMHost $VMHost.Name -V2 -ErrorAction Stop
            }
            Catch{
                $ESXCli = $null
            }

            Try{
                $PCINIC = $ESXCli.Network.NIC.List.Invoke()
            }
            Catch{
                $PCINIC = $null
                $vCenterError.Add([PSCustomObject]@{
                    'Object' = 'Host'
                    'Name'   = $VMHost.Name
                    'Error'  = "Failed to enumerate NIC list on $($VMHost.Name)"
                }) | Out-Null
            }

            $NetworkSystem = $HostObjView.ConfigManager.NetworkSystem
            $NetworkSystemView = Get-View $NetworkSystem
            Try{
                $HostDistributedVirtualSwitches = Get-VDSwitch -VMHost $VMHost.Name -ErrorAction Stop
            }
            Catch{
                $HostDistributedVirtualSwitches = $null
            }

            $HostCertificate = $HostObjView.ConfigManager.CertificateManager
            Try{
                $HostCertificateView = Get-View $HostCertificate -ErrorAction Stop
                $HostCertificateExpiration = $HostCertificateView.CertificateInfo.NotAfter
            }
            Catch{
                $HostCertificateExpiration = $null
            }

            If($HostCertificateExpiration){
                $HostCertExpireInDays = ($HostCertificateExpiration - (Get-Date)).Days
            }
            Else{
                $HostCertExpireInDays = $null
            }

            Try{
                $DataCenterName = ($VMHost | Get-Datacenter -ErrorAction Stop).Name
            }
            Catch{
                $DataCenterName = $null
            }

            If($ErrorCount -eq 0){
                $HostData += [PSCustomObject]@{
                    "Name"             = $VMHost.Name
                    "ESXi"             = $VMHost.Version
                    "Build"            = $VMHost.Build
                    "Maintenance Mode" = $HostObjView.Runtime.InMaintenanceMode
                    "Lockdown Mode"    = $HostObjView.Config.LockdownMode
                    "Vendor"           = $HostObjView.Hardware.SystemInfo.Vendor
                    "Model"            = $HostObjView.Hardware.SystemInfo.Model
                    "Serial"           = $HostSerialNumber
                    "Processor Type"   = $VMHost.ProcessorType
                    "CPU Count"        = $HostObjView.Hardware.CpuInfo.NumCpuPackages
                    "Cores"            = $HostObjView.Hardware.CpuInfo.NumCpuCores
                    "RAM"              = ("" + [math]::round($HostObjView.Hardware.MemorySize/1GB,0) + "GB")
                    "BIOS"             = $HostObjView.Hardware.BiosInfo.BiosVersion
                    "Days Up"          = New-TimeSpan -Start $VMHost.ExtensionData.Summary.Runtime.BootTime -End (Get-Date) | Select-Object -ExpandProperty Days
                    "NTP Servers"      = $NTPServers -join ", "
                    "DNS Servers"      = $VMHost.ExtensionData.Config.Network.DnsConfig.Address -join ", "
                    "Cert Expires"     = $HostCertificateExpiration
                    "Days to Expire"   = $HostCertExpireInDays
                    "VMs"              = ($VMHost | Get-VM | Measure-Object).Count
                    "Cluster"          = $VMHost.Parent.Name
                    "Datacenter"       = $DataCenterName
                    "vCenter Server"   = $VCServer
                }
            }
            Else{
                Write-Warning "Host error count equals $ErrorCount on host $($VMHost.Name)"
            }
            #endregion

            #region Host NICs
            Try{
                $HostNICs = Get-VMHostNetworkAdapter -VMHost $VMHost.Name -Physical -ErrorAction Stop
            }
            Catch{
                $HostNICs = $null
                $vCenterError += [PSCustomObject]@{
                    "Object" = "Host"
                    "Name"   = $VMHost.Name
                    "Error"  = "The Get-VMHostNetworkAdapter (Physical) command failed on $($VMHost.Name)"
                }
            }

            ForEach($HostNic in $HostNICs){
                $NetworkHint = $null
                $CDPExtended = $null
                $vSwitch = $null
                $vSwitchType = $null

                If($PCINIC){
                    $PCINICProps = $PCINIC | Where-Object{$HostNic.Name -eq $_.Name}
                }
                Else{
                    $PCINICProps = $null
                }

                $NetworkHint = $NetworkSystemView.QueryNetworkHint($HostNic.Name)
                $CDPExtended = $NetworkHint.ConnectedSwitchPort
                # Check if NIC is connected to distributed switch
                If($HostDistributedVirtualSwitches){
                    $vSwitchType = "Distributed"
                    ForEach($HostDVS in $HostDistributedVirtualSwitches){
                        $DVSMatch = $null
                        $DVSMatch = Get-VMHostNetworkAdapter -DistributedSwitch $HostDVS -VMHost $VMHost -Physical | Where-Object{$_.Name -eq $HostNic.Name}

                        If($DVSMatch){
                            $vSwitch = $HostDVS
                            Break
                        }
                    }
                }

                # If no distributed switch detected, check for standard switch
                If($null -eq $vSwitch){
                    $vSwitchType = "Standard"
                    $vSwitch = $VMHost | Get-VirtualSwitch -Standard | Where-Object{$_.NIC -eq $HostNic.DeviceName}
                }

                If($null -eq $vSwitch){
                    $vSwitchType = $null
                }

                $HostNicData += [PSCustomObject]@{
                    "Host"         = $VMHost.Name
                    "Name"         = $HostNic.Name
                    "MAC"          = $HostNic.MAC
                    "DHCP"         = $HostNic.DhcpEnabled
                    "Link"         = $PCINICProps.Link
                    "Speed"        = $PCINICProps.Speed
                    "Duplex"       = $PCINICProps.Duplex
                    "Vendor"       = $PCINICProps.Description
                    "Driver"       = $PCINICProps.Driver
                    "vSwitch"      = $vSwitch.Name
                    "vSwitch Type" = $vSwitchType
                    "DVS Ver"      = $vSwitch.Version
                    "DVS MTU"      = $vSwitch.Mtu
                    "Switch"       = $CDPExtended.DevID
                    "Switch IP"    = $CDPExtended.Address
                    "Switch Port"  = $CDPExtended.PortID
                    "Cluster"      = $VMHost.Parent.Name
                    "Datacenter"   = $DataCenterName
                }
            }
            #endregion

            #region VMKernel Adapters
            Try{
                $HostVMKs = Get-VMHostNetworkAdapter -VMHost $VMHost.Name -VMKernel -ErrorAction Stop | Select-Object Name,DeviceName,Mac,IP,DhcpEnabled,SubnetMask,MTU,PortGroupName,VMotionEnabled
            }
            Catch{
                $HostVMKs = $null
                $vCenterError += [PSCustomObject]@{
                    "Object" = "Host"
                    "Name"   = $VMHost.Name
                    "Error"  = "The Get-VMHostNetworkAdapter (VMKernel) command failed on $($VMHost.Name)"
                }
            }

            ForEach($HostVMK in $HostVMKs){
                $HostVMKData += [PSCustomObject]@{
                    "Host"        = $VMHost.Name
                    "Name"        = $HostVMK.Name
                    "Device"      = $HostVMK.DeviceName
                    "IP"          = $HostVMK.IP
                    "Subnet Mask" = $HostVMK.SubnetMask
                    "MAC"         = $HostVMK.MAC
                    "DHCP"        = $HostVMK.DhcpEnabled
                    "MTU"         = $HostVMK.Mtu
                    "Port Group"  = $HostVMK.PortGroupName
                    "vMotion"     = $HostVMK.VMotionEnabled
                    "Cluster"     = $VMHost.Parent.Name
                    "Datacenter"  = $DataCenterName
                }
            }
            #endregion
        }
    }

    #region VMs
    Try{
        $VMachines = Get-VM -Server $VCServer -ErrorAction Stop
    }
    Catch{
        $VMachines = $null
    }

    If($VMachines){
        $VMCounter = 0
        ForEach($VMachine in $VMachines){
            $VMCounter ++
            Write-Progress -Activity "vCenter server $VCServer" -Status "Gathering virtual machine information..." -CurrentOperation $VMachine.Name -PercentComplete ($VMCounter*100/($VMachines | Measure-Object).Count)
            $VMConnectionState = $VMachine.ExtensionData.Summary.Runtime.ConnectionState
            Try{
                $VMProps = Get-VM -Name $VMachine.Name | Get-View -ErrorAction Stop
            }
            Catch{
                $VMProps = $null
                $vCenterError += [PSCustomObject]@{
                    "Object" = "Virtual Machine"
                    "Name"   = $VMachine.Name
                    "Error"  = "The Get-VM command failed on $($VMachine.Name)"
                }
            }

            Try{
                $VMNotes = $VMachine | Select-Object -ExpandProperty Notes -ErrorAction Stop
            }
            Catch{
                $VMNotes = $null
            }

            Try{
                $VMNicProps = Get-NetworkAdapter -VM $VMachine.Name -ErrorAction Stop
            }
            Catch{
                $VMNicProps = $null
                $vCenterError += [PSCustomObject]@{
                    "Object" = "Virtual Machine"
                    "Name"   = $VMachine.Name
                    "Error"  = "The Get-NetworkAdapter command failed on $($VMachine.Name)"
                }
            }

            Try{
                $VMHardDiskProps = Get-HardDisk -VM $VMachine.Name -ErrorAction Stop
            }
            Catch{
                $VMHardDiskProps = $null
                $vCenterError += [PSCustomObject]@{
                    "Object" = "Virtual Machine"
                    "Name"   = $VMachine.Name
                    "Error"  = "The Get-HardDisk command failed on $($VMachine.Name)"
                }
            }

            # VM USB Controller attached
            Try{
                $VMUSB = $VMProps.Config.Hardware.Device | Where-Object {$_.Gettype().Name -match 'VirtualUSB'}
            }
            Catch{
                $VMUSB = $null
                $vCenterError.Add([PSCustomObject]@{
                    'Object' = 'Virtual Machine'
                    'Name'   = $VMachine.Name
                    'Error'  = "Could not retrieve Virtual USB information on $($VMachine.Name)"
                }) | Out-Null
            }
            If($VMUSB){
                $VMUSBAttached = $true
            }
            Else{
                $VMUSBAttached = $false
            }

            Try{
                $VSnapshots = $VMachine | Get-Snapshot -ErrorAction Stop
            }
            Catch{
                $VSnapshots = $null
            }

            $EncryptedProps = $VMProps.ExtensionData.Config.KeyId
            If($null -eq $EncryptedProps){
                $EncryptedVM = $False
            }
            Else{
                $EncryptedVM = $EncryptedProps.KeyId
            }

            $VMData += [PSCustomObject]@{
                'Name'                   = $VMachine.Name                                                  # Column A
                'OS'                     = $VMProps.Summary.Config.GuestFullName                           # Column B
                'OS Family'              = $VMProps.Guest.GuestFamily                                      # Column C
                'Tools Version'          = $VMProps.Guest.ToolsVersion                                     # Column D
                'Tools Status'           = $VMProps.Guest.ToolsVersionStatus                               # Column E
                'Tools Policy'           = $VMProps.Config.Tools.ToolsUpgradePolicy                        # Column F
                'HardwareVer'            = $VMachine.HardwareVersion                                       # Column G
                'Key ID'                 = $EncryptedVM                                                    # Column H
                'Virtual Based Security' = $VMachine.ExtensionData.Config.Flags.VbsEnabled                 # Column I
                'Secure Boot'            = $VMachine.ExtensionData.Config.BootOptions.EfiSecureBootEnabled # Column J
                'State'                  = $VMachine.PowerState                                            # Column K
                'IP'                     = $VMachine.Guest.IPAddress -join ', '                            # Column L
                'CPUs'                   = $VMachine.NumCpu                                                # Column M
                'RAM'                    = ('' + [math]::round($VMachine.MemoryGB) + 'GB')                 # Column N
                'NICs'                   = ($VMNicProps | Measure-Object).Count                            # Column O
                'USB Controller'         = $VMUSBAttached                                                  # Column P
                'Disks'                  = ($VMHardDiskProps | Measure-Object).Count                       # Column Q
                'Used Raw'               = $VMachine.UsedSpaceGB * 1GB                                     # Column R
                'Used Space'             = Get-Size ($VMachine.UsedSpaceGB * 1GB)                          # Column S
                'Snapshots'              = ($VSnapshots | Measure-Object).Count                            # Column T
                'Consolidate'            = $VMachine.ExtensionData.Runtime.ConsolidationNeeded             # Column U
                'Folder'                 = $VMProps.Folder.Name                                            # Column V
                'Host'                   = $VMachine.VMHost.Name                                           # Column W
                'Cluster'                = $VMachine.VMHost.Parent.Name                                    # Column X
                'Datacenter'             = ($VMachine | Get-Datacenter).Name                               # Column Y
                'Notes'                  = $VMNotes                                                        # Column Z
                'VM Path'                = $VMachine.ExtensionData.Config.Files.VmPathName                 # Column AA
                'Connection State'       = $VMConnectionState                                              # Column AB
            }
            #endregion

            #region VM NICs
            ForEach($VNic in $VMNicProps){
                $VMNicData += [PSCustomObject]@{
                    "VM"             = $VMachine.Name
                    "NIC"            = $VNic.Name
                    "Type"           = $VNic.Type
                    "Connected"      = $VNic.ConnectionState.Connected
                    "ConnectAtStart" = $VNic.ConnectionState.StartConnected
                    "Network"        = $VNic.NetworkName
                    "MAC"            = $VNic.MacAddress
                }
            }
            #endregion

            # region VM Snapshots
            If($VSnapshots){
                ForEach($VSnapshot in $VSnapshots){
                    $SnapshotData += [PSCustomObject]@{
                        "VM"             = $VMachine.Name
                        "Snapshot"       = $VSnapshot.Name
                        "Created"        = $VSnapshot.Created
                        "Raw Size"       = ([int]$VSnapshot.SizeMB)*1MB
                        "Size"           = Get-Size (([int]$VSnapshot.SizeMB)*1MB)
                        "VM State"       = $VMachine.PowerState
                        "Snapshot State" = $VSnapshot.PowerState
                        "Description"    = $VSnapshot.Description
                    }
                }
            }
            #endregion

            #region VM Hard Disk
            ForEach($VMHardDisk in $VMHardDiskProps){
                $VMDKUsedSize = $null
                Try{
                    $VMDKRawUsedSize = ($VMachine.ExtensionData.LayoutEx.file | Where-Object{$_.Name -contains $VMHardDisk.FileName.replace(".vmdk","-flat.vmdk")} -ErrorAction Stop).Size
                }
                Catch{
                    $VMDKRawUsedSize = $null
                }
                If($VMDKRawUsedSize){
                    $VMDKUsedSize = Get-Size $VMDKRawUsedSize
                }
                Else{
                    $VMDKUsedSize = "N/A"
                }
                If($VMHardDisk.CapacityGB){
                    $VMDKRawCapacity = $VMHardDisk.CapacityGB*1GB
                    $VMDKCapacity = Get-Size $VMDKRawCapacity
                }
                Else{
                    $VMDKRawCapacity = "N/A"
                    $VMDKCapacity = "N/A"
                }

                $VMHardDiskData += [PSCustomObject]@{
                    "VM"           = $VMachine.Name
                    "Disk"         = $VMHardDisk.Name
                    "Raw Capacity" = $VMDKRawCapacity
                    "Capacity"     = $VMDKCapacity
                    "Raw Used"     = $VMDKRawUsedSize
                    "Used"         = $VMDKUsedSize
                    "Persistence"  = $VMHardDisk.Persistence
                    "Format"       = $VMHardDisk.StorageFormat
                    "Type"         = $VMHardDisk.DiskType
                    "Datastore"    = ($VMHardDisk.Filename).Split("]")[0].Split("[")[1]
                    "File Name"    = $VMHardDisk.FileName
                }
            }

            ForEach($VMDrive in $VMProps.Guest.Disks){
                $VMDriveData += [PSCustomObject]@{
                    "VM"           = $VMachine.Name
                    "Path"         = $VMDrive.Path
                    "Raw Capacity" = $VMDrive.Capacity
                    "Capacity"     = Get-Size $VMDrive.Capacity
                    "Raw Free"     = $VMDrive.FreeSpace
                    "Free"         = Get-Size $VMDrive.FreeSpace
                    "Raw Used"     = $VMDrive.Capacity - $VMDrive.FreeSpace
                    "Used"         = Get-Size ($VMDrive.Capacity - $VMDrive.FreeSpace)
                }
            }
        }
    }
    #endregion

    #region Datastores
    Try{
        $VDatastores = Get-Datastore -Server $VCServer -ErrorAction Stop
        Write-Progress -Activity "vCenter server $VCServer" -Status "Gathering Datastore information..."
    }
    Catch{
        $VDatastores = $null
        $vCenterError += [PSCustomObject]@{
            "Object" = "Datastore"
            "Name"   = $VCServer
            "Error"  = "The Get-Datastore command failed on $VCServer"
        }
    }

    ForEach($VDatastore in $VDatastores){
        If($VDatastore.CapacityGB -eq 0){
            $DatastorePctFree = 0
        }
        Else{
            $DatastorePctFree = [math]::Round((($VDatastore.FreeSpaceGB/$VDatastore.CapacityGB)*100),2)
        }
        $DatastoresData += [PSCustomObject]@{
            "Store"      = $VDatastore.Name
            "CapacityGB" = [math]::Round($VDatastore.CapacityGB,2)
            "FreeGB"     = [math]::Round($VDatastore.FreeSpaceGB,2)
            "% Free"     = $DatastorePctFree
            "Type"       = $VDatastore.Type
            "FSVer"      = $VDatastore.FilesystemVersion
            "Folder"     = $VDatastore.ParentFolder
            "State"      = $VDatastore.State
            "Datacenter" = $VDatastore.Datacenter
            "vCenter"    = $VCServer
            "Path"       = $VDatastore.DatastoreBrowserPath
        }
    }
    #endregion

    Disconnect-VIServer -Server $VCServer -Confirm:$False
}

Write-Progress -Activity "vCenter server $VCServer" -Completed

#region Output to Excel
# Create Excel standard configuration properties
$ExcelProps = @{
    Autosize = $true;
    FreezeTopRow = $true;
    BoldTopRow = $true;
}

$ExcelProps.Path = $LogFile

# vCenter sheet
$vCenterObjectLastRow = ($vCenterObject | Measure-Object).Count + 1
If($vCenterObjectLastRow -gt 1){
    $vCenterObjectHeaderCount = Get-ColumnName ($vCenterObject | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $vCenterObjectHeaderRow = "'vCenter Servers'!`$A`$1:`$$vCenterObjectHeaderCount`$1"

    $vCenterObjectStyle = New-ExcelStyle -Range $vCenterObjectHeaderRow -HorizontalAlignment Center

    $vCenterObject | Sort-Object "Name" | Export-Excel @ExcelProps -WorksheetName "vCenter Servers" -Style $vCenterObjectStyle
}

# Datacenter sheet
$DatacenterLastRow = ($DatacenterData | Measure-Object).Count + 1
If($DatacenterLastRow -gt 1){
    $DatacenterHeaderCount = Get-ColumnName ($DatacenterData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $DatacenterHeaderRow = "'Datacenters'!`$A`$1:`$$DatacenterHeaderCount`$1"

    $DatacenterStyle = New-ExcelStyle -Range $DatacenterHeaderRow -HorizontalAlignment Center

    $DatacenterData | Sort-Object "vCenter","Name" | Export-Excel @ExcelProps -WorksheetName "Datacenters" -Style $DatacenterStyle
}

# Cluster sheet
$ClusterDataLastRow = ($ClusterData | Measure-Object).Count + 1
If($ClusterDataLastRow -gt 1){
    $ClusterDataHeaderCount = Get-ColumnName ($ClusterData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $ClusterDataHeaderRow = "'Clusters'!`$A`$1:`$$ClusterDataHeaderCount`$1"

    $ClusterDataStyle = New-ExcelStyle -Range $ClusterDataHeaderRow -HorizontalAlignment Center

    $ClusterData | Sort-Object "vCenter","Datacenter","Cluster" | Export-Excel @ExcelProps -WorkSheetname "Clusters" -Style $ClusterDataStyle
}

# Cluster DRS Rule sheet
$ClusterDRSDataLastRow = ($ClusterDRSData | Measure-Object).Count + 1
If($ClusterDRSDataLastRow -gt 1){
    $ClusterDRSDataHeaderCount = Get-ColumnName ($ClusterDRSData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $ClusterDRSDataHeaderRow = "'Cluster DRS'!`$A`$1:`$$ClusterDRSDataHeaderCount`$1"

    $ClusterDRSDataStyle = New-ExcelStyle -Range $ClusterDRSDataHeaderRow -HorizontalAlignment Center

    $ClusterDRSData | Sort-Object "vCenter", "Datacenter", "Cluster", "Rule" | Export-Excel @ExcelProps -WorkSheetname "Cluster DRS" -Style $ClusterDRSDataStyle
}

# Licensing sheet
$LicensingLastRow = ($LicenseCustomObject | Measure-Object).Count + 1
If($LicensingLastRow -gt 1){
    $LicensingHeaderCount = Get-ColumnName ($LicenseCustomObject | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $LicensingHeaderRow = "Licenses!`$A`$1:`$$LicensingHeaderCount`$1"

    $LicensingDataStyle = @()
    $LicensingDataStyle += New-ExcelStyle -Range $LicensingHeaderRow -HorizontalAlignment Center

    $LicenseCustomObject | Sort-Object "License Host","Product" | Export-Excel @ExcelProps -WorksheetName "Licenses" -Style $LicensingDataStyle
}

# Assigned Licenses sheet
$AssignedLicensesLastRow = ($AssignedLicenseObject | Measure-Object).Count + 1
If($AssignedLicensesLastRow -gt 1){
    $AssignedLicensesHeaderCount = Get-ColumnName ($AssignedLicenseObject | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $AssignedLicensesHeaderRow = "Assigned Licenses!`$A`$1:`$$AssignedLicensesHeaderCount`$1"

    $AssignedLicensesDataStyle = @()
    $AssignedLicensesDataStyle += New-ExcelStyle -Range $AssignedLicensesHeaderRow -HorizontalAlignment Center

    $AssignedLicenseObject | Sort-Object "License Name","Entity" | Export-Excel @ExcelProps -WorksheetName "Assigned Licenses" -Style $AssignedLicensesDataStyle
}

# Host sheet
$HostDataLastRow = ($HostData | Measure-Object).Count + 1
If($HostDataLastRow -gt 1){
    $HostDataHeaderCount   = Get-ColumnName ($HostData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $HostDataHeaderRow     = "Hosts!`$A`$1:`$$HostDataHeaderCount`$1"
    $MMColumn              = "Hosts!`$D`$2:`$D`$$HostDataLastRow"
    $LockdownColumn        = "Hosts!`$E`$2:`$E`$$HostDataLastRow"
    $NTPColumn             = "Hosts!`$O`$2:`$O`$$HostDataLastRow"
    $DaysCertExpiresColumn = "Hosts!`$R`$2:`$R`$$HostDataLastRow"

    $HostDataStyle = New-ExcelStyle -Range $HostDataHeaderRow -HorizontalAlignment Center

    $HostDataConditionalFormatting = @()
    $HostDataConditionalFormatting += New-ConditionalText -Range $MMColumn -ConditionalType ContainsText "TRUE" -ConditionalTextColor Brown -BackgroundColor Yellow
    $HostDataConditionalFormatting += New-ConditionalText -Range $LockdownColumn -ConditionalType ContainsText "lockdownDisabled" -ConditionalTextColor Brown -BackgroundColor Yellow
    $HostDataConditionalFormatting += New-ConditionalText -Range $NTPColumn -ConditionalType ContainsBlanks -BackgroundColor Yellow
    $HostDataConditionalFormatting += New-ConditionalText -Range $DaysCertExpiresColumn -ConditionalType LessThanOrEqual '30' -ConditionalTextColor Maroon -BackgroundColor Pink
    $HostDataConditionalFormatting += New-ConditionalText -Range $DaysCertExpiresColumn -ConditionalType LessThanOrEqual '60' -ConditionalTextColor Brown -BackgroundColor Yellow

    $HostData | Sort-Object "vCenter Server","Datacenter","Cluster","Name" | Export-Excel @ExcelProps -WorkSheetname "Hosts" -Style $HostDataStyle -ConditionalText $HostDataConditionalFormatting
}

# Host NIC sheet
$HostNicDataLastRow = ($HostNicData | Measure-Object).Count + 1
If($HostNicDataLastRow -gt 1){
    $HostNicDataHeaderCount = Get-ColumnName ($HostNicData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $HostNicDataHeaderRow   = "'Host NICs'!`$A`$1:`$$HostNicDataHeaderCount`$1"
    $HostNicDataLinkColumn  = "'Host NICs'!`$E`$2:`$E`$$HostNicDataLastRow"

    $HostNicDataStyle = New-ExcelStyle -Range $HostNicDataHeaderRow -HorizontalAlignment Center
    
    $HostNicDataConditionalFormatting = New-ConditionalText -Range $HostNicDataLinkColumn -ConditionalType ContainsText "Up" -ConditionalTextColor DarkGreen -BackgroundColor LightGreen

    $HostNicData | Sort-Object "Host","Name" | Export-Excel @ExcelProps -WorkSheetname "Host NICs" -Style $HostNicDataStyle -ConditionalFormat $HostNicDataConditionalFormatting
}

# Host VMK sheet
$HostVMKDataLastRow = ($HostVMKData | Measure-Object).Count + 1
If($HostVMKDataLastRow -gt 1){
    $HostVMKDataHeaderCount = Get-ColumnName ($HostVMKData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $HostVMKDataHeaderRow   = "'Host VMKs'!`$A`$1:`$$HostVMKDataHeaderCount`$1"

    $HostVMKDataStyle = New-ExcelStyle -Range $HostVMKDataHeaderRow -HorizontalAlignment Center

    $HostVMKData | Sort-Object "vCenter Server","Datacenter","Cluster","Host","Name" | Export-Excel @ExcelProps -WorksheetName "Host VMKs" -Style $HostVMKDataStyle
}

# VM sheet
$VMDataLastRow = ($VMData | Measure-Object).Count + 1
If($VMDataLastRow -gt 1){
    $VMDataHeaderCount = Get-ColumnName ($VMData | Get-Member | Where-Object{$_.MemberType -match 'NoteProperty'} | Measure-Object).Count
    $VMDataHeaderRow = "'VMs'!`$A`$1:`$$VMDataHeaderCount`$1"
    $VMUSBAttachedRow = "'VMs'!`$P`$2:`$P`$$VMDataLastRow"
    $VMDataUsedSpaceRawColumn = "'VMs'!`$R`$2:`$R`$$VMDataLastRow"
    $VMSnapshotColumn = "'VMs'!`$T`$2:`$T`$$VMDataLastRow"
    $VMConsolidationColumn = "'VMs'!`$U`$2:`$U`$$VMDataLastRow"
    $VMOrphanedColumn = "'VMs'!`$AB`$2:`$AB`$$VMDataLastRow"

    $VMDataStyle = @()
    $VMDataStyle += New-ExcelStyle -Range $VMDataHeaderRow -HorizontalAlignment Center
    $VMDataStyle += New-ExcelStyle -Range $VMDataUsedSpaceRawColumn -NumberFormat '0'
    $VMDataStyle += New-ExcelStyle -Range $VMSnapshotColumn -NumberFormat '0'

    $VMDataConditionalFormatting = @()
    $VMDataConditionalFormatting += New-ConditionalText -Range $VMSnapshotColumn -ConditionalType GreaterThanOrEqual '1' -ConditionalTextColor Brown -BackgroundColor Yellow
    $VMDataConditionalFormatting += New-ConditionalText -Range $VMConsolidationColumn -ConditionalType ContainsText 'TRUE' -ConditionalTextColor Brown -BackgroundColor Yellow
    $VMDataConditionalFormatting += New-ConditionalText -Range $VMUSBAttachedRow -ConditionalType ContainsText 'TRUE' -ConditionalTextColor Brown -BackgroundColor Yellow
    $VMDataConditionalFormatting += New-ConditionalText -Range $VMOrphanedColumn -ConditionalType ContainsText 'orphaned' -ConditionalTextColor Maroon -BackgroundColor Pink
    $VMDataConditionalFormatting += New-ConditionalText -Range $VMOrphanedColumn -ConditionalType ContainsText 'inaccessible' -ConditionalTextColor Brown -BackgroundColor Yellow
    $VMDataConditionalFormatting += New-ConditionalText -Range $VMOrphanedColumn -ConditionalType ContainsText 'disconnected' -ConditionalTextColor Brown -BackgroundColor Yellow

    $VMData | Sort-Object 'Name' | Export-Excel @ExcelProps -WorksheetName 'VMs' -Style $VMDataStyle -ConditionalFormat $VMDataConditionalFormatting
}

# VM NIC sheet
$VMNicDataLastRow = ($VMNicData | Measure-Object).Count + 1
If($VMNicDataLastRow -gt 1){
    $VMNicDataHeaderCount = Get-ColumnName ($VMNicData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $VMNicDataHeaderRow   = "'VM NICs'!`$A`$1:`$$VMNicDataHeaderCount`$1"

    $VMNicDataStyle = New-ExcelStyle -Range $VMNicDataHeaderRow -HorizontalAlignment Center

    $VMNicData | Sort-Object "VM" | Export-Excel @ExcelProps -WorkSheetname "VM NICs" -Style $VMNicDataStyle
}

# VM Disk sheet
$VMHDDataLastRow = ($VMHardDiskData | Measure-Object).Count + 1
If($VMHDDataLastRow -gt 1){
    $VMHDDataHeaderCount    = Get-ColumnName ($VMHardDiskData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $VMHDDataHeaderRow      = "'VM Disks'!`$A`$1:`$$VMHDDataHeaderCount`$1"
    $VMHDDataCapacityColumn = "'VM Disks'!`$C`$2:`$C`$$VMHDDataLastRow"
    $VMHDDataUsedColumn     = "'VM Disks'!`$E`$2:`$E`$$VMHDDataLastRow"
    $VMHDFormatColumn       = "'VM Disks'!`$H`$2:`$H`$$VMHDDataLastRow"

    $VMHDDataStyle = @()
    $VMHDDataStyle += New-ExcelStyle -Range $VMHDDataHeaderRow -HorizontalAlignment Center
    $VMHDDataStyle += New-ExcelStyle -Range $VMHDDataCapacityColumn -NumberFormat '0'
    $VMHDDataStyle += New-ExcelStyle -Range $VMHDDataUsedColumn -NumberFormat '0'

    $VMHDDataConditionalFormatting = New-ConditionalText -Range $VMHDFormatColumn -ConditionalType NotContainsText "Thin" -ConditionalTextColor Maroon -BackgroundColor Pink

    $VMHardDiskData | Sort-Object "VM","Disk" | Export-Excel @ExcelProps -WorkSheetname "VM Disks" -Style $VMHDDataStyle -ConditionalFormat $VMHDDataConditionalFormatting
}

# VM Drive sheet
$VMDriveDataLastRow = ($VMDriveData | Measure-Object).Count + 1
If($VMDriveDataLastRow -gt 1){
    $VMDriveDataHeaderCount    = Get-ColumnName ($VMDriveData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $VMDriveDataHeaderRow      = "'VM Drives'!`$A`$1:`$$VMDriveDataHeaderCount`$1"
    $VMDriveDataCapacityColumn = "'VM Drives'!`$C`$2:`$C`$$VMDriveDataLastRow"
    $VMDriveDataFreeColumn     = "'VM Drives'!`$E`$2:`$E`$$VMDriveDataLastRow"
    $VMDriveDataUsedColumn     = "'VM Drives'!`$G`$2:`$G`$$VMDriveDataLastRow"

    $VMDriveDataStyle = @()
    $VMDriveDataStyle += New-ExcelStyle -Range $VMDriveDataHeaderRow -HorizontalAlignment Center
    $VMDriveDataStyle += New-ExcelStyle -Range $VMDriveDataCapacityColumn -NumberFormat '0'
    $VMDriveDataStyle += New-ExcelStyle -Range $VMDriveDataFreeColumn -NumberFormat '0'
    $VMDriveDataStyle += New-ExcelStyle -Range $VMDriveDataUsedColumn -NumberFormat '0'

    $VMDriveData | Sort-Object "VM" | Export-Excel @ExcelProps -WorkSheetname "VM Drives" -Style $VMDriveDataStyle
}

# Snapshot sheet
$SnapshotLastRow = ($SnapshotData | Measure-Object).Count + 1
If($SnapshotLastRow -gt 1){
    $SnapshotHeaderCount = Get-ColumnName ($SnapshotData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $SnapshotHeaderRow   = "Snapshots!`$A`$1:`$$SnapshotHeaderCount`$1"
    $SnapshotRawsizeColumn = "'Snapshots'!`$D`$2:`$D`$$SnapshotLastRow"

    $SnapshotDataStyle = @()
    $SnapshotDataStyle += New-ExcelStyle -Range $SnapshotHeaderRow -HorizontalAlignment Center
    $SnapshotDataStyle += New-ExcelStyle -Range $SnapshotRawsizeColumn -NumberFormat '0'

    $SnapshotData | Export-Excel @ExcelProps -WorksheetName "Snapshots" -Style $SnapshotDataStyle
}

# Datastore sheet
$DatastoreLastRow = ($DatastoresData | Measure-Object).Count + 1
If($DatastoreLastRow -gt 1){
    $DatastoresDataHeaderCount = Get-ColumnName ($DatastoresData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $DatastoresDataHeaderRow   = "Datastores!`$A`$1:`$$DatastoresDataHeaderCount`$1"
    $DatastoresDataCapacityRow = "Datastores!`$B`$2:`$B`$$DatastoreLastRow"
    $DatastoresDataFreeRow     = "Datastores!`$C`$2:`$C`$$DatastoreLastRow"
    $DatastorePctFreeColumn    = "Datastores!`$D`$2:`$D`$$DatastoreLastRow"
    $DatastoreAvailableColumn  = "Datasotres!`$H`$2:`$H`$$DatastoreLastRow"

    $DatastoresDataStyle = @()
    $DatastoresDataStyle += New-ExcelStyle -Range $DatastoresDataHeaderRow -HorizontalAlignment Center
    $DatastoresDataStyle += New-ExcelStyle -Range $DatastoresDataCapacityRow -NumberFormat '0.00'
    $DatastoresDataStyle += New-ExcelStyle -Range $DatastoresDataFreeRow -NumberFormat '0.00'
    $DatastoresDataStyle += New-ExcelStyle -Range $DatastorePctFreeColumn -NumberFormat '0.00'

    $DatastoresConditionalFormatting = @()
    $DatastoresConditionalFormatting += New-ConditionalText -Range $DatastorePctFreeColumn -ConditionalType LessThanOrEqual "10" -ConditionalTextColor Maroon -BackgroundColor Pink
    $DatastoresConditionalFormatting += New-ConditionalText -Range $DatastorePctFreeColumn -ConditionalType LessThanOrEqual "20" -ConditionalTextColor Brown -BackgroundColor Yellow
    $DatastoresConditionalFormatting += New-ConditionalText -Range $DatastoreAvailableColumn -ConditionalType NotContainsText "Available" -ConditionalTextColor Maroon -BackgroundColor Pink

    $DatastoresData | Sort-Object "Path","Store" | Export-Excel @ExcelProps -WorkSheetname "Datastores" -Style $DatastoresDataStyle -ConditionalFormat $DatastoresConditionalFormatting
}

# Error sheet
$vCenterErrorLastRow = ($vCenterError | Measure-Object).Count + 1
If($vCenterErrorLastRow -gt 1){
    $vCenterErrorHeaderCount = Get-ColumnName ($vCenterError | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $vCenterErrorHeaderRow = "Errors!`$A`$1:`$$vCenterErrorHeaderCount`$1"

    $vCenterErrorStyle = @()
    $vCenterErrorStyle += New-ExcelStyle -Range $vCenterErrorHeaderRow -HorizontalAlignment Center

    $vCenterError | Sort-Object "Object","Name" | Export-Excel @ExcelProps -WorksheetName "Errors" -Style $vCenterErrorStyle
}
#endregion
