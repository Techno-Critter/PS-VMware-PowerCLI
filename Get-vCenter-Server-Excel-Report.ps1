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

#Requires -Module ImportExcel
#Requires -Module VMware.PowerCLI

#region Configure variables
$Date = Get-Date -Format yyyyMMdd
$LogFile = "C:\Temp\vCenter\vCenter_Report_$Date.xlsx"
$VCServers = "server1.acme.com","server2.acme.com"
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
$HeaderRow = ("!`$A`$1:`$ZZ`$1")

# Cluster sheet
$ClusterDataStyle = New-ExcelStyle -Range "Clusters$HeaderRow" -HorizontalAlignment Center
$ClusterData | Sort-Object Datacenter,Cluster | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "Clusters" -Style $ClusterDataStyle

# Host sheet
$HostDataLastRow = ($HostData | Measure-Object).Count + 1
$HostDataStyle = New-ExcelStyle -Range "Hosts$HeaderRow" -HorizontalAlignment Center
If($HostDataLastRow -gt 1){
    $MMColumn = "Hosts!`$D`$2:`$D`$$HostDataLastRow"
    $LockdownColumn = "Hosts!`$E`$2:`$E`$$HostDataLastRow"
    $NTPColumn = "Hosts!`$N`$2:`$N`$$HostDataLastRow"
    $HostData | Sort-Object Name | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "Hosts" -Style $HostDataStyle -ConditionalText $(
        New-ConditionalText -Range $MMColumn -ConditionalType ContainsText "TRUE" -ConditionalTextColor Brown -BackgroundColor Wheat
        New-ConditionalText -Range $LockdownColumn -ConditionalType ContainsText "lockdownDisabled" -ConditionalTextColor Brown -BackgroundColor Wheat
        New-ConditionalText -Range $NTPColumn -ConditionalType NotContainsText "172.16.127.253, 172.16.255.253" -ConditionalTextColor Brown -BackgroundColor Wheat
    )
}

# Host NIC sheet
$HostNicDataStyle = New-ExcelStyle -Range "'Host NICs'$HeaderRow" -HorizontalAlignment Center
$HostNicData | Sort-Object Host,Name | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "Host NICs" -Style $HostNicDataStyle

# VM sheet
$VMDataLastRow = ($VMData | Measure-Object).Count + 1
If($VMDataLastRow -gt 1){
    $VMDataUsedSpaceColumn = "'VMs'!`$L`$2:`$L`$$VMDataLastRow"
    $VMDataStyle = @()
    $VMDataStyle += New-ExcelStyle -Range "VMs$HeaderRow" -HorizontalAlignment Center
    $VMDataStyle += New-ExcelStyle -Range $VMDataUsedSpaceColumn -NumberFormat '0'
    $VMData | Sort-Object Name | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "VMs" -Style $VMDataStyle
}

# VM NIC sheet
$VMNicDataStyle = New-ExcelStyle -Range "'VM NICs'$HeaderRow" -HorizontalAlignment Center
$VMNicData | Sort-Object VM | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "VM NICs" -Style $VMNicDataStyle

# VM Disk sheet
$VMHDDataLastRow = ($VMHardDiskData | Measure-Object).Count + 1
If($VMHDDataLastRow -gt 1){
    $VMHDDataCapacityColumn = "'VM Disks'!`$C`$2:`$C`$$VMHDDataLastRow"
    $VMHDDataStyle = @()
    $VMHDDataStyle += New-ExcelStyle -Range "'VM Disks'$HeaderRow" -HorizontalAlignment Center
    $VMHDDataStyle += New-ExcelStyle -Range $VMHDDataCapacityColumn -NumberFormat '0'
    $VMHDFormatColumn = "'VM Disks'!`$F`$2:`$F`$$VMHDDataLastRow"
    $VMHardDiskData | Sort-Object VM | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "VM Disks" -Style $VMHDDataStyle -ConditionalFormat $(
        New-ConditionalText -Range $VMHDFormatColumn -ConditionalType NotContainsText "Thin" -ConditionalTextColor Maroon -BackgroundColor Pink
    )
}

# VM Drive sheet
$VMDriveDataLastRow = ($VMDriveData | Measure-Object).Count + 1
If($VMDriveDataLastRow -gt 1){
    $VMDriveDataCapacityColumn = "'VM Drives'!`$C`$2:`$C`$$VMDriveDataLastRow"
    $VMDriveDataFreeColumn = "'VM Drives'!`$E`$2:`$E`$$VMDriveDataLastRow"
    $VMDriveDataUsedColumn = "'VM Drives'!`$G`$2:`$G`$$VMDriveDataLastRow"
    $VMDriveDataStyle = @()
    $VMDriveDataStyle += New-ExcelStyle -Range "'VM Drives'$HeaderRow" -HorizontalAlignment Center
    $VMDriveDataStyle += New-ExcelStyle -Range $VMDriveDataCapacityColumn -NumberFormat '0'
    $VMDriveDataStyle += New-ExcelStyle -Range $VMDriveDataFreeColumn -NumberFormat '0'
    $VMDriveDataStyle += New-ExcelStyle -Range $VMDriveDataUsedColumn -NumberFormat '0'
    $VMDriveData | Sort-Object VM | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "VM Drives" -Style $VMDriveDataStyle
}

# Datastore sheet
$DatastoreLastRow = ($DatastoresData | Measure-Object).Count + 1
If($DatastoreLastRow -gt 1){
    $DatastorePctFreeColumn = "Datastores!`$D`$2:`$D`$$DatastoreLastRow"
    $DatastoresDataStyle = New-ExcelStyle -Range "Datastores$HeaderRow" -HorizontalAlignment Center
    $DatastoresData | Sort-Object Store | Export-Excel -Path $LogFile -AutoSize -FreezeTopRow -BoldTopRow -WorkSheetname "Datastores" -Style $DatastoresDataStyle -ConditionalFormat $(
        New-ConditionalText -Range $DatastorePctFreeColumn -ConditionalType GreaterThanOrEqual "90" -ConditionalTextColor Maroon -BackgroundColor Pink
        New-ConditionalText -Range $DatastorePctFreeColumn -ConditionalType GreaterThanOrEqual "80" -ConditionalTextColor Brown -BackgroundColor Yellow
    )
}
#endregion
