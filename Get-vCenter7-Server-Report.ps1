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
$LogFile = "C:\Logs\vCenter\vCenter_Report_$Date.xlsx"
$VCServers = @("vcenter-1.acme.com","vcenter-2.acme.com")
$LoginAccount = "vcenterer@acme.com" #NOTE: if login account changes a new credential file will be created!
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

#region Configure arrays
$DatacenterData = @()
$ClusterData = @()
$HostData = @()
$HostNicData = @()
$HostVMKData = @()
$VMData = @()
$VMNicData = @()
$VMDriveData = @()
$VMHardDiskData = @()
$DatastoresData = @()
$SnapshotData = @()
$VCLicenseServers = @()
$LicenseCustomObject = @()
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
ForEach($VCServer in $VCServers){
    $vCenterError = $null
    Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$False
    Try{
        Connect-VIServer -Server $VCServer -Credential $Credentials
    }
    Catch{
        $vCenterError = "The server $VCServer did not accept the connection request. This vCenter server will be skipped."
        Write-Warning $vCenterError
    }

    If($null -eq $vCenterError){

        #region Datacenters
        Try{
            $VDatacenters = Get-Datacenter -Server $VCServer -ErrorAction Stop
        }
        Catch{
            $VDatacenters = $null
        }

        If($null -ne $VDatacenters){
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

        If($null -ne $VClusters){
            ForEach($VCluster in $VCLusters){
                $ClusterData += [PSCustomObject]@{
                    "Cluster"    = $VCluster.Name
                    "Datacenter" = (Get-Datacenter -Cluster $VCluster).Name
                    "vCenter"    = $VCServer
                    "HA"         = $VCluster.HAEnabled
                    "DRS"        = $VCluster.DrsEnabled
                    "AutoLevel"  = $VCluster.DrsAutomationLevel
                    "Hosts"      = (Get-VMHost -Location $VCluster.Name | Measure-Object).Count
                    "VMs"        = (Get-VM -Location $VCluster.Name | Measure-Object).Count
                }
            }
        }
        #endregion

        #region Licensing
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
        }
        #endregion

        #region Hosts
        Try{
            $VMHosts = Get-VMHost -Server $VCServer -ErrorAction Stop
        }
        Catch{
            $VMHosts = $null
        }

        If($null -ne $VMHosts){
            ForEach($VMHost in $VMHosts){
                $ErrorCount = 0
                Try{
                    $HostObjView = $VMHost | Get-View -ErrorAction Stop
                }
                Catch{
                    $HostObjView = $null
                    $ErrorCount ++
                }

                Try{
                    $NTPServers =  Get-VMHostNtpServer -VMHost $VMHost.Name -ErrorAction Stop
                }
                Catch{
                    $NTPServers = "Error"
                    $ErrorCount ++
                }

                If($null -eq $HostObjView.Hardware.SystemInfo.OtherIdentifyingInfo){
                        $HostSerialNumber = "N/A"
                }
                Else{
                    $HostSerialNumber = $HostObjView.Hardware.SystemInfo.OtherIdentifyingInfo[1].IdentifierValue
                }

                Try{
                    $ESXCli = Get-EsxCli -VMHost $VMHost.Name -V2 -ErrorAction Stop
                }
                Catch{
                    $ESXCli = $null
                }
                If($ESXCli){
                    $PCINIC = $ESXCli.Network.NIC.List.Invoke()
                }
                Else{
                    $PCINIC = $null
                }
                $NetworkSystem = $HostObjView.ConfigManager.NetworkSystem
                $NetworkSystemView = Get-View $NetworkSystem
                $DVSwitchInfo = Get-VDSwitch -VMHost $VMHost
                If($null -ne $DVSwitchInfo){
                    $DVSwitchHost = $DVSwitchInfo.ExtensionData.Config.Host
                    $DVNIC = $DVSwitchHost.Config.Backing.PCINICPropspec.PnicDevice
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
                        "CPU Count"        = $HostObjView.Hardware.CpuInfo.NumCpuPackages
                        "Cores"            = $HostObjView.Hardware.CpuInfo.NumCpuCores
                        "RAM"              = ("" + [math]::round($HostObjView.Hardware.MemorySize/1GB,0) + "GB")
                        "BIOS"             = $HostObjView.Hardware.BiosInfo.BiosVersion
                        "Days Up"          = New-TimeSpan -Start $VMHost.ExtensionData.Summary.Runtime.BootTime -End (Get-Date) | Select-Object -ExpandProperty Days
                        "NTP Servers"      = $NTPServers -join ", "
                        "DNS Servers"      = $VMHost.ExtensionData.Config.Network.DnsConfig.Address -join ", "
                        "VMs"              = ($VMHost | Get-VM | Measure-Object).Count
                        "Cluster"          = ($VMHost | Get-Cluster).Name
                        "Datacenter"       = ($VMHost | Get-Datacenter).Name
                        "vCenter Server"   = $VCServer
                    }
                }
                Else{
                    Write-Warning "Host error count equals $ErrorCount on host $($VMHost.Name)"
                }
                #endregion

                #region Host NICs
                    Try{
                        $HostNICs = Get-VMHostNetworkAdapter -VMHost $VMHost.Name -Physical -ErrorAction Stop | Select-Object Name,Mac,DhcpEnabled
                    }
                    Catch{
                        $HostNICs = $null
                    }

                    ForEach($HostNic in $HostNICs){
                        $NetworkHint = $null
                        $CDPExtended = $null
                        $vSwitch = $null
                        $vSwitchName = $null

                        If($PCINIC){
                            $PCINICProps = $PCINIC | Where-Object{$HostNic.Name -eq $_.Name} | Select-Object "Description","Link","Duplex","MTU","Driver","Speed"
                        }
                        Else{
                            $PCINICProps = $null
                        }

                        $NetworkHint = $NetworkSystemView.QueryNetworkHint($HostNic.Name)
                        $CDPExtended = $NetworkHint.ConnectedSwitchPort
                        # Check if NIC is connected to distributed switch or standard switch
                        If($HostNic.Name -eq $DVNIC){
                            $vSwitch = $DVSwitchInfo | Where-Object{$HostNic.Name -eq $DVNIC} | Select-Object -ExpandProperty "Name"
                        }
                        Else{
                            $vSwitchName = $VMHost | Get-VirtualSwitch -Standard | Where-Object{$_.NIC -eq $HostNic.DeviceName}
                            $vSwitch = $vSwitchName.Name
                        }

                        $HostNicData += [PSCustomObject]@{
                            "Host"        = $VMHost.Name
                            "Name"        = $HostNic.Name
                            "MAC"         = $HostNic.MAC
                            "DHCP"        = $HostNic.DhcpEnabled
                            "Link"        = $PCINICProps.Link
                            "Speed"       = $PCINICProps.Speed
                            "Duplex"      = $PCINICProps.Duplex
                            "Vendor"      = $PCINICProps.Description
                            "Driver"      = $PCINICProps.Driver
                            "vSwitch"     = $vSwitch
                            "Switch"      = $CDPExtended.DevID
                            "Switch IP"   = $CDPExtended.Address
                            "Switch Port" = $CDPExtended.PortID
                            "Cluster"     = ($VMHost | Get-Cluster).Name
                            "Datacenter"  = ($VMHost | Get-Datacenter).Name
                        }
                    }
                    #endregion

                    #region VMKernel Adapters
                    Try{
                        $HostVMKs = Get-VMHostNetworkAdapter -VMHost $VMHost.Name -VMKernel -ErrorAction Stop | Select-Object Name,DeviceName,Mac,IP,DhcpEnabled,SubnetMask,MTU,PortGroupName,VMotionEnabled
                    }
                    Catch{
                        $HostVMKs = $null
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
                            "Cluster"     = ($VMHost | Get-Cluster).Name
                            "Datacenter"  = ($VMHost | Get-Datacenter).Name
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

        If($null -ne $VMachines){
            ForEach($VMachine in $VMachines){
                Try{
                    $VMProps = Get-VM -Name $VMachine.Name | Get-View -ErrorAction Stop
                }
                Catch{
                    $VMProps = $null
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
                }
                Try{
                    $VMHardDiskProps = Get-HardDisk -VM $VMachine.Name -ErrorAction Stop
                }
                Catch{
                    $VMHardDiskProps = $null
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
                    "Name"          = $VMachine.Name
                    "OS"            = $VMProps.Summary.Config.GuestFullName
                    "OS Family"     = $VMProps.Guest.GuestFamily
                    "Tools Version" = $VMProps.Guest.ToolsVersion
                    "Tools Status"  = $VMProps.Guest.ToolsVersionStatus
                    "Tools Policy"  = $VMProps.Config.Tools.ToolsUpgradePolicy
                    "HardwareVer"   = $VMachine.HardwareVersion
                    "Key ID"        = $EncryptedVM
                    "State"         = $VMachine.PowerState
                    "IP"            = $VMProps.Guest.IPAddress -join ", "
                    "CPUs"          = $VMachine.NumCpu
                    "RAM"           = ("" + [math]::round($VMachine.MemoryGB) + "GB")
                    "NICs"          = ($VMNicProps | Measure-Object).Count
                    "Disks"         = ($VMHardDiskProps | Measure-Object).Count
                    "UsedSpace Raw" = $VMachine.UsedSpaceGB*1GB
                    "UsedSpace"     = Get-Size ($VMachine.UsedSpaceGB*1GB)
                    "Snapshots"     = ($VSnapshots | Measure-Object).Count
                    "Consolidate"   = $VMachine.ExtensionData.Runtime.ConsolidationNeeded
                    "Folder"        = $VMProps.Folder.Name
                    "Host"          = (Get-VMHost -VM $VMachine.Name).Name
                    "Cluster"       = (Get-Cluster -VM $VMachine.Name).Name
                    "Datacenter"    = (Get-Datacenter -VM $VMachine.Name).Name
                    "Notes"         = $VMNotes
                    "VM Path"       = $VMachine.ExtensionData.Config.Files.VmPathName
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
                            "VM"          = $VMachine.Name
                            "Snapshot"    = $VSnapshot.Name
                            "Created"     = $VSnapshot.Created
                            "State"       = $VSnapshot.PowerState
                            "Description" = $VSnapshot.Description
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

                    $VMHardDiskData += [PSCustomObject]@{
                        "VM"           = $VMachine.Name
                        "Disk"         = $VMHardDisk.Name
                        "Raw Capacity" = $VMHardDisk.CapacityGB*1GB
                        "Capacity"     = Get-Size ($VMHardDisk.CapacityGB*1GB)
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
        }
        Catch{
            $VDatastores = $null
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
}

#region Output to Excel
# Create Excel standard configuration properties
$ExcelProps = @{
    Autosize = $true;
    FreezeTopRow = $true;
    BoldTopRow = $true;
}

$ExcelProps.Path = $LogFile

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

# Licensing sheet
$LicensingLastRow = ($LicenseCustomObject | Measure-Object).Count + 1
If($LicensingLastRow -gt 1){
    $LicensingHeaderCount = Get-ColumnName ($LicenseCustomObject | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $LicensingHeaderRow = "Licenses!`$A`$1:`$$LicensingHeaderCount`$1"

    $LicensingDataStyle = @()
    $LicensingDataStyle += New-ExcelStyle -Range $LicensingHeaderRow -HorizontalAlignment Center

    $LicenseCustomObject | Sort-Object "License Host","Product" | Export-Excel @ExcelProps -WorksheetName "Licenses" -Style $LicensingDataStyle
}

# Host sheet
$HostDataLastRow = ($HostData | Measure-Object).Count + 1
If($HostDataLastRow -gt 1){
    $HostDataHeaderCount = Get-ColumnName ($HostData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $HostDataHeaderRow   = "Hosts!`$A`$1:`$$HostDataHeaderCount`$1"
    $MMColumn            = "Hosts!`$D`$2:`$D`$$HostDataLastRow"
    $LockdownColumn      = "Hosts!`$E`$2:`$E`$$HostDataLastRow"
    $NTPColumn           = "Hosts!`$N`$2:`$N`$$HostDataLastRow"

    $HostDataStyle = New-ExcelStyle -Range $HostDataHeaderRow -HorizontalAlignment Center

    $HostDataConditionalFormatting = @()
    $HostDataConditionalFormatting += New-ConditionalText -Range $MMColumn -ConditionalType ContainsText "TRUE" -ConditionalTextColor Brown -BackgroundColor Yellow
    $HostDataConditionalFormatting += New-ConditionalText -Range $LockdownColumn -ConditionalType ContainsText "lockdownDisabled" -ConditionalTextColor Brown -BackgroundColor Yellow
    $HostDataConditionalFormatting += New-ConditionalText -Range $NTPColumn -ConditionalType ContainsBlanks -BackgroundColor Yellow

    $HostData | Sort-Object "vCenter Server","Datacenter","Cluster","Name" | Export-Excel @ExcelProps -WorkSheetname "Hosts" -Style $HostDataStyle -ConditionalText $HostDataConditionalFormatting
}

# Host NIC sheet
$HostNicDataLastRow = ($HostNicData | Measure-Object).Count + 1
If($HostNicDataLastRow -gt 1){
    $HostNicDataHeaderCount = Get-ColumnName ($HostNicData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $HostNicDataHeaderRow   = "'Host NICs'!`$A`$1:`$$HostNicDataHeaderCount`$1"

    $HostNicDataStyle = New-ExcelStyle -Range $HostNicDataHeaderRow -HorizontalAlignment Center

    $HostNicData | Sort-Object "Host","Name" | Export-Excel @ExcelProps -WorkSheetname "Host NICs" -Style $HostNicDataStyle
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
    $VMDataHeaderCount        = Get-ColumnName ($VMData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
    $VMDataHeaderRow          = "'VMs'!`$A`$1:`$$VMDataHeaderCount`$1"
    $VMDataUsedSpaceRawColumn = "'VMs'!`$O`$2:`$O`$$VMDataLastRow"
    $VMSnapshotColumn         = "'VMs'!`$Q`$2:`$Q`$$VMDataLastRow"
    $VMConsolidationColumn    = "'VMs'!`$R`$2:`$R`$$VMDataLastRow"

    $VMDataStyle = @()
    $VMDataStyle += New-ExcelStyle -Range $VMDataHeaderRow -HorizontalAlignment Center
    $VMDataStyle += New-ExcelStyle -Range $VMDataUsedSpaceRawColumn -NumberFormat '0'
    $VMDataStyle += New-ExcelStyle -Range $VMSnapshotColumn -NumberFormat '0'

    $VMDataConditionalFormatting = @()
    $VMDataConditionalFormatting += New-ConditionalText -Range $VMSnapshotColumn -ConditionalType GreaterThanOrEqual "1" -ConditionalTextColor Brown -BackgroundColor Yellow
    $VMDataConditionalFormatting += New-ConditionalText -Range $VMConsolidationColumn -ConditionalType ContainsText "TRUE" -ConditionalTextColor Brown -BackgroundColor Yellow

    $VMData | Sort-Object "Name" | Export-Excel @ExcelProps -WorkSheetname "VMs" -Style $VMDataStyle -ConditionalFormat $VMDataConditionalFormatting
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
#endregion
