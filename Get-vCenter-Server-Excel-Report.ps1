<#
Author: Stan Crider
Date: 7October2019
What this crap does:
Gather information from specified vCenter servers and outputs to Excel
### Must have at least read-access to vCenter
### vCenter must have a Datacenter/Cluster/Host heirarchy
### Must have VMware PowerCLI module installed!!!
### Must have ImportExcel module installed!!!
### https://github.com/dfinke/ImportExcel
#>

#region Configure variables
$Date = Get-Date -Format yyyyMMdd
$LogFile = "C:\Temp\vCenter\vCenter_Report_$Date.xlsx"
$VCServers = "server1.acme.com","server2.acme.com"
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
$ClusterData = @()
$HostData = @()
$HostNicData = @()
$VMData = @()
$VMNicData = @()
$VMDriveData = @()
$VMHardDiskData = @()
$DatastoresData = @()
#endregion

# Connect to vCenters and retrieve data
ForEach($VCServer in $VCServers){
    Connect-VIServer -Server $VCServer
    $VDatacenters = Get-Datacenter -Server $VCServer
    
# Datacenters
    ForEach($VDatacenter in $VDatacenters){
        $VClusters = Get-Cluster -Location $VDatacenter.Name
#region Clusters
        ForEach($VCluster in $VCLusters){
            $VMHosts = Get-VMHost -Location $VCluster.Name -ErrorAction SilentlyContinue
            $ClusterData += [PSCustomObject]@{
                "Cluster" = $VCluster.Name
                "Datacenter" = $VDatacenter.Name
                "HA" = $VCluster.HAEnabled
                "DRS" = $VCluster.DrsEnabled
                "AutoLevel" = $VCluster.DrsAutomationLevel
                "Hosts" = ($VMHosts | Measure-Object).Count
                "VMs" = (Get-VM -Location $VCluster.Name | Measure-Object).Count
            }
#endregion

#region Hosts
            ForEach($VMHost in $VMHosts){
                $HostObjView = $VMHost | Get-View -ErrorAction SilentlyContinue
                $VMachines = Get-VM -Location $VMHost.Name -ErrorAction SilentlyContinue
                $NTPServers =  Get-VMHostNtpServer -VMHost $VMHost.Name
                $HostData += [PSCustomObject]@{
                    "Name" = $VMHost.Name
                    "ESXi" = $VMHost.Version
                    "Build" = $VMHost.Build
                    "Maintenance Mode" = $HostObjView.Runtime.InMaintenanceMode
                    "Lockdown Mode" = $HostObjView.Config.LockdownMode
                    "Vendor" = $HostObjView.Hardware.SystemInfo.Vendor
                    "Model" = $HostObjView.Hardware.SystemInfo.Model
                    "Serial" = $HostObjView.Hardware.SystemInfo.OtherIdentifyingInfo[1].IdentifierValue
                    "CPU Count" = $HostObjView.Hardware.CpuInfo.NumCpuPackages
                    "Cores" = $HostObjView.Hardware.CpuInfo.NumCpuCores
                    "RAM" = ("" + [math]::round($HostObjView.Hardware.MemorySize/1GB,0) + "GB")
                    "BIOS" = $HostObjView.Hardware.BiosInfo.BiosVersion
                    "Days Up" = New-TimeSpan -Start $VMHost.ExtensionData.Summary.Runtime.BootTime -End (Get-Date) | Select-Object -ExpandProperty Days
                    "NTP Servers" = $NTPServers -join ", "
                    "VMs" = ($VMachines | Measure-Object).Count
                    "Cluster" = $VCluster.Name
                    "Datacenter" = $VDatacenter.Name
                }
#endregion

#region Host NICs
                $HostNICs = Get-VMHostNetworkAdapter -VMHost $VMHost.Name -VMKernel -ErrorAction SilentlyContinue | Select-Object Name,DeviceName,Mac,IP,DhcpEnabled,SubnetMask,MTU,PortGroupName,VMotionEnabled
                ForEach($HostNic in $HostNICs){
                    $HostNicData += [PSCustomObject]@{
                        "Host" = $VMHost.Name
                        "Name" = $HostNic.Name
                        "Device" = $HostNic.DeviceName
                        "IP" = $HostNic.IP
                        "Subnet Mask" = $HostNic.SubnetMask
                        "MAC" = $HostNic.MAC
                        "DHCP" = $HostNic.DhcpEnabled
                        "MTU" = $HostNic.Mtu
                        "Port Group" = $HostNic.PortGroupName
                        "vMotion" = $HostNic.VMotionEnabled
                        "Cluster" = $VCluster.Name
                        "Datacenter" = $VDatacenter.Name
                    }
                }
#endregion

#region VMs
                ForEach($VMachine in $VMachines){
                    $VMProps = Get-VM -Name $VMachine.Name -ErrorAction SilentlyContinue
                    $VMNicProps = Get-NetworkAdapter -VM $VMachine.Name -ErrorAction SilentlyContinue
                    $VMHardDiskProps = Get-HardDisk -VM $VMachine.Name -ErrorAction SilentlyContinue
                    $VMData += [PSCustomObject]@{
                        "Name" = $VMachine.Name
                        "OS" = $VMProps.Guest.OSFullName
                        "OS Family" = $VMProps.Guest.GuestFamily
                        "Tools" = $VMProps.Guest.ToolsVersion
                        "HardwareVer" = $VMProps.HardwareVersion
                        "State" = $VMProps.PowerState
                        "IP" = $VMProps.Guest.IPAddress -join ", "
                        "CPUs" = $VMProps.NumCpu
                        "RAM" = ("" + [math]::round($VMProps.MemoryGB) + "GB")
                        "NICs" = ($VMNicProps | Measure-Object).Count
                        "Disks" = ($VMHardDiskProps | Measure-Object).Count
                        "UsedSpace Raw" = $VMProps.UsedSpaceGB*1GB
                        "UsedSpace" = Get-Size ($VMProps.UsedSpaceGB*1GB)
                        "Folder" = $VMProps.Folder.Name
                        "Host" = $VMHost.Name
                        "Cluster" = $VCluster.Name
                        "Datacenter" = $VDatacenter.Name
                        "Notes" = $VMProps.Notes
                    }
#endregion

#region VM NICs
                    ForEach($VNic in $VMNicProps){
                        $VMNicData += [PSCustomObject]@{
                            "VM" = $VMachine.Name
                            "NIC" = $VNic.Name
                            "Type" = $VNic.Type
                            "Connected" = $VNic.ConnectionState.Connected
                            "ConnectAtStart" = $VNic.ConnectionState.StartConnected
                            "Network" = $VNic.NetworkName
                            "MAC" = $VNic.MacAddress
                        }
                    }
#endregion

#region VM Hard Disk
                    ForEach($VMHardDisk in $VMHardDiskProps){
                        $VMHardDiskData += [PSCustomObject]@{
                            "VM" = $VMachine.Name
                            "Disk" = $VMHardDisk.Name
                            "Raw Capacity" = $VMHardDisk.CapacityGB*1GB
                            "Capacity" = Get-Size ($VMHardDisk.CapacityGB*1GB)
                            "Persistence" = $VMHardDisk.Persistence
                            "Format" = $VMHardDisk.StorageFormat
                            "Type" = $VMHardDisk.DiskType
                            "Datastore" = ($VMHardDisk.Filename).Split("]")[0].Split("[")[1]
                            "File Name" = $VMHardDisk.FileName
                        }
                    }

                    ForEach($VMDrive in $VMProps.Guest.Disks){
                        $VMDriveData += [PSCustomObject]@{
                            "VM" = $VMachine.Name
                            "Path" = $VMDrive.Path
                            "Raw Capacity" = $VMDrive.Capacity
                            "Capacity" = Get-Size $VMDrive.Capacity
                            "Raw Free" = $VMDrive.FreeSpace
                            "Free" = Get-Size $VMDrive.FreeSpace
                            "Raw Used" = $VMDrive.Capacity - $VMDrive.FreeSpace
                            "Used" = Get-Size ($VMDrive.Capacity - $VMDrive.FreeSpace)
                        }
                    }
                }
            }
        }
#endregion

#region Datastores
        $VDatastores = Get-Datastore
        ForEach($VDatastore in $VDatastores){
            If($VDatastore.CapacityGB -eq 0){
                $DatastorePctFree = 0
            }
            Else{
                $DatastorePctFree = [math]::Round((($VDatastore.FreeSpaceGB/$VDatastore.CapacityGB)*100),2)
            }
            $DatastoresData += [PSCustomObject]@{
                "Store" = $VDatastore.Name
                "CapacityGB" = [math]::Round($VDatastore.CapacityGB,2)
                "FreeGB" = [math]::Round($VDatastore.FreeSpaceGB,2)
                "% Free" = $DatastorePctFree
                "Type" = $VDatastore.Type
                "FSVer" = $VDatastore.FilesystemVersion
                "Folder" = $VDatastore.ParentFolder
                "State" = $VDatastore.State
                "Datacenter" = $VDatastore.Datacenter
                "Path" = $VDatastore.DatastoreBrowserPath
            }
        }
#endregion
    }
    Disconnect-VIServer -Confirm:$False
}

#region Output to Excel

# Cluster sheet
$ClusterColumnCount = Get-ColumnName ($ClusterData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
$ClusterHeader = "`$A`$1:`$$ClusterColumnCount`$1"
$ClusterDataStyle = New-ExcelStyle -Range "Clusters!$ClusterHeader" -HorizontalAlignment Center
$ClusterData | Sort-Object Datacenter,Cluster | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "Clusters" -Style $ClusterDataStyle

# Host sheet
$HostsColumnCount = Get-ColumnName ($HostData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
$HostHeader = "`$A`$1:`$$HostsColumnCount`$1"
$HostDataLastRow = ($HostData | Measure-Object).Count + 1
$HostDataStyle = New-ExcelStyle -Range "Hosts!$HostHeader" -HorizontalAlignment Center
If($HostDataLastRow -gt 1){
    $HostDataConditionalText = @()
    $HostDataConditionalText += New-ConditionalText -Range "Hosts!`$D`$2:`$D`$$HostDataLastRow" -ConditionalType ContainsText "TRUE" -ConditionalTextColor Brown -BackgroundColor Wheat
    $HostDataConditionalText += New-ConditionalText -Range "Hosts!`$E`$2:`$E`$$HostDataLastRow" -ConditionalType ContainsText "lockdownDisabled" -ConditionalTextColor Brown -BackgroundColor Wheat
    $HostDataConditionalText += New-ConditionalText -Range "Hosts!`$N`$2:`$N`$$HostDataLastRow" -ConditionalType NotContainsText "tick.ia.doi.net, tock.ia.doi.net" -ConditionalTextColor Brown -BackgroundColor Wheat
    $HostData | Sort-Object Name | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "Hosts" -Style $HostDataStyle -ConditionalText $HostDataConditionalText
}

# Host NIC sheet
$HostNicColumnCount = Get-ColumnName ($HostNicData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
$HostNicHeader = "`$A`$1:`$$HostNicColumnCount`$1"
$HostNicDataStyle = New-ExcelStyle -Range "'Host NICs!'$HostNicHeader" -HorizontalAlignment Center
$HostNicData | Sort-Object Host,Name | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "Host NICs" -Style $HostNicDataStyle

# VM sheet
$VMDataColumnCount = Get-ColumnName ($HostNicData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
$VMDataHeader = "`$A`$1:`$$VMDataColumnCount`$1"
$VMDataLastRow = ($VMData | Measure-Object).Count + 1
If($VMDataLastRow -gt 1){
    $VMDataUsedSpaceColumn = "'VMs'!`$L`$2:`$L`$$VMDataLastRow"
    $VMDataStyle = @()
    $VMDataStyle += New-ExcelStyle -Range "VMs!$VMDataHeader" -HorizontalAlignment Center
    $VMDataStyle += New-ExcelStyle -Range $VMDataUsedSpaceColumn -NumberFormat '0'
    $VMData | Sort-Object Name | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "VMs" -Style $VMDataStyle
}

# VM NIC sheet
$VMNicDataColumnCount = Get-ColumnName ($VMNicData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
$VMNicHeader = "`$A`$1:`$$VMNicDataColumnCount`$1"
$VMNicDataStyle = New-ExcelStyle -Range "'VM NICs!'$VMNicHeader" -HorizontalAlignment Center
$VMNicData | Sort-Object VM | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "VM NICs" -Style $VMNicDataStyle

# VM Disk sheet
$VMHardDiskColumnCount = Get-ColumnName ($VMHardDiskData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
$VMHardDiskHeader = "`$A`$1:`$$VMHardDiskColumnCount`$1"
$VMHDDataLastRow = ($VMHardDiskData | Measure-Object).Count + 1
If($VMHDDataLastRow -gt 1){
    $VMHDDataStyle = @()
    $VMHDDataStyle += New-ExcelStyle -Range "'VM Disks!'$VMHardDiskHeader" -HorizontalAlignment Center
    $VMHDDataStyle += New-ExcelStyle -Range "'VM Disks'!`$C`$2:`$C`$$VMHDDataLastRow" -NumberFormat '0'
    $VMHDDataConditionalText = @()
    $VMHDDataConditionalText += New-ConditionalText -Range "'VM Disks'!`$F`$2:`$F`$$VMHDDataLastRow" -ConditionalType NotContainsText "Thin" -ConditionalTextColor Maroon -BackgroundColor Pink
    $VMHardDiskData | Sort-Object VM | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "VM Disks" -Style $VMHDDataStyle -ConditionalFormat $VMHDDataConditionalText
}

# VM Drive sheet
$VMDriveDataColumnCount = Get-ColumnName ($VMDriveData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
$VMDriveDataHeader = "`$A`$1:`$$VMDriveDataColumnCount`$1"
$VMDriveDataLastRow = ($VMDriveData | Measure-Object).Count + 1
If($VMDriveDataLastRow -gt 1){
    $VMDriveDataStyle = @()
    $VMDriveDataStyle += New-ExcelStyle -Range "'VM Drives!'$VMDriveDataHeader" -HorizontalAlignment Center
    $VMDriveDataStyle += New-ExcelStyle -Range "'VM Drives'!`$C`$2:`$C`$$VMDriveDataLastRow" -NumberFormat '0'
    $VMDriveDataStyle += New-ExcelStyle -Range "'VM Drives'!`$E`$2:`$E`$$VMDriveDataLastRow" -NumberFormat '0'
    $VMDriveDataStyle += New-ExcelStyle -Range "'VM Drives'!`$G`$2:`$G`$$VMDriveDataLastRow" -NumberFormat '0'
    $VMDriveData | Sort-Object VM | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "VM Drives" -Style $VMDriveDataStyle
}

# Datastore sheet
$DatastoresDataColumnCount = Get-ColumnName ($DatastoresData | Get-Member | Where-Object{$_.MemberType -match "NoteProperty"} | Measure-Object).Count
$DatastoresDataHeader = "`$A`$1:`$$DatastoresDataColumnCount`$1"
$DatastoreLastRow = ($DatastoresData | Measure-Object).Count + 1
If($DatastoreLastRow -gt 1){
    $DatastoresDataStyle = New-ExcelStyle -Range "Datastores!$DatastoresDataHeader" -HorizontalAlignment Center
    $DatastorePctFreeColumn = "Datastores!`$D`$2:`$D`$$DatastoreLastRow"
    $DatastoresDataConditionalText = @()
    $DatastoresDataConditionalText += New-ConditionalText -Range $DatastorePctFreeColumn -ConditionalType LessThanOrEqual "10" -ConditionalTextColor Maroon -BackgroundColor Pink
    $DatastoresDataConditionalText += New-ConditionalText -Range $DatastorePctFreeColumn -ConditionalType LessThanOrEqual "20" -ConditionalTextColor Brown -BackgroundColor Wheat
    $DatastoresData | Sort-Object Store | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "Datastores" -Style $DatastoresDataStyle -ConditionalFormat $DatastoresDataConditionalText
}
#endregion
