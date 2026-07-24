<#
Author: Stan Crider
Date: 21July2026
What this crap does:
Gather VM allocated resource information from specified vCenter servers,
calculates monthly charge based on preset cost variables, and outputs to Excel.
Rates are based on fixed resources allotted to VM's, not usage based. However,
storage usage and allocation are included in the spreadsheet for comparison and
can be changed in calculations if desired. CPU and RAM are fixed.
### Must have at least read-access to vCenter
### Must have VMware.PowerCLI.VCenter module installed!!!
### Must have ImportExcel module installed!!!
### https://github.com/dfinke/ImportExcel
#>

#Requires -Modules ImportExcel, VMware.PowerCLI.VCenter

#region Configure variables
$Date = Get-Date -Format yyyyMMdd-hhmm
# File name and folder location where spreadsheet will be created
$LogFile = "C:\Temp\VM-Info $Date.xlsx"
# vCenter servers
$VCServers = @(
    'vcenter-1.acme.com'
    'vcenter-2.acme.com'
)
# If login account changes a new credential file will be created!
$LoginAccount = 'techno@acme.com'
# Folder location where credential file will be created and stored; file name will be created automatically
# User must have write access to credential file directory!!!
$CredentialFileDirectory = 'C:\Temp\Credentials'
#endregion

#region Cost variables
#Costs for charge back model used in calculating monthly rate per VM
# CPU rate is per core
# RAM is per GB
# Storage rate is per TB
$CPURate = 5.52
$RAMRate = 1.41
$StorageRate = 59.4
$FlatRate = 7.55
# Convert flat rate to currency format
$FlatCurrency = $FlatRate.ToString("C",[System.Globalization.CultureInfo]::CurrentCulture)
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
            $FirstNumber = [Math]::Floor(($ColumnCount) / 26)
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
        $CharacterOutput = 'ZZ'
    }

    # Output column name
    $CharacterOutput
}
#endregion

#region Configure arrays and counters
[System.Collections.ArrayList]$vCenterError = @()
[System.Collections.ArrayList]$VMData = @()
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

#region Connect vCenter
# Connect to vCenters and retrieve data
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$False | Out-Null
ForEach($VCServer in $VCServers){
    $VCServerCounter ++
    Write-Progress -Activity "vCenter server $VCServer" -Status ('Connecting to vCenter server ' + $VCServerCounter + ' of ' + (($VCServers | Measure-Object).Count) + '.')
    Try{
        Connect-VIServer -Server $VCServer -Credential $Credentials -ErrorAction Stop | Out-Null
        #Connect-VIServer -Server $VCServer -AllLinked -Credential $Credentials -ErrorAction Stop | Out-Null
    }
    Catch{
        $vCenterError.Add([PSCustomObject]@{
                'Object' = 'vCenter'
                'Name'   = $VCServer
                'Error'  = "The server $VCServer did not accept the connection request. This vCenter server will be skipped."
            }) | Out-Null
        Continue
    }
#endregion

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
            Write-Progress -Activity "vCenter server $VCServer" -Status 'Gathering virtual machine information...' -CurrentOperation $VMachine.Name -PercentComplete ($VMCounter * 100 / ($VMachines | Measure-Object).Count)

            Try{
                $VMHardDiskProps = Get-HardDisk -VM $VMachine.Name -ErrorAction Stop
            }
            Catch{
                $VMHardDiskProps = $null
                $vCenterError.Add([PSCustomObject]@{
                        'Object' = 'Virtual Machine'
                        'Name'   = $VMachine.Name
                        'Error'  = "The Get-HardDisk command failed on $($VMachine.Name)"
                    }) | Out-Null
            }

            #region VM Hard Disk
            $VMTotalHardDiskCapacityRaw = $null

            ForEach($VMHardDisk in $VMHardDiskProps){
                If($VMHardDisk.CapacityGB){
                    $VMDKRawCapacity = $VMHardDisk.CapacityGB * 1GB
                }
                Else{
                    $VMDKRawCapacity = '0'
                }

                $VMTotalHardDiskCapacityRaw += $VMDKRawCapacity
            }

            $CPUCost = $VMachine.NumCpu * $CPURate
            $RAMCost = $VMachine.MemoryGB * $RAMRate 
            $StorageCost = $StorageRate * ($VMTotalHardDiskCapacityRaw  / [double]1TB
            $CPUCurrency = $CPUCost.ToString("C",[System.Globalization.CultureInfo]::CurrentCulture)
            $RAMCurrency = $RAMCost.ToString("C",[System.Globalization.CultureInfo]::CurrentCulture)
            $StorageCurrency = $StorageCost.ToString("C",[System.Globalization.CultureInfo]::CurrentCulture)
            $MonthlyRate = ($CPUCost + $RAMCost + $StorageCost + $FlatRate)
            $MonthlyCurrency = $MonthlyRate.ToString("C",[System.Globalization.CultureInfo]::CurrentCulture)

            $VMData.Add([PSCustomObject]@{ 
                'Name'         = $VMachine.Name                                 # Column A
                'CPUs'         = $VMachine.NumCpu                               # Column B
                'RAM GB'       = $VMachine.MemoryGB                             # Column C
                'Disks'        = ($VMHardDiskProps | Measure-Object).Count      # Column D
                'Used Raw'     = $VMachine.UsedSpaceGB * 1GB                    # Column E
                'Used GB'      = $VMachine.UsedSpaceGB                          # Column F
                'Capacity Raw' = $VMTotalHardDiskCapacityRaw                    # Column G
                'Capacity GB'  = ($VMTotalHardDiskCapacityRaw / [double]1GB)    # Column H
                'Host'         = $VMachine.VMHost.Name                          # Column I
                'Cluster'      = $VMachine.VMHost.Parent.Name                   # Column J
                'CPU Cost'     = $CPUCurrency                                   # Column K
                'RAM Cost'     = $RAMCurrency                                   # Column L
                'Storage'      = $StorageCurrency                               # Column M
                'Flat Rate'    = $FlatCurrency                                  # Column N
                'Monthly Cost' = $MonthlyCurrency                               # Column O
            }) | Out-Null
            #endregion
        }
    }
    #endregion

#region Disconnect vCenter
    Disconnect-VIServer -Server * -Confirm:$False
}

Write-Progress -Activity "vCenter server $VCServer" -Completed
#endregion

#region Output to Excel
# Create Excel standard configuration properties
$ExcelProps = @{
    Autosize     = $true
    FreezeTopRow = $true
    BoldTopRow   = $true
}

$ExcelProps.Path = $LogFile

# VM sheet
$VMDataLastRow = ($VMData | Measure-Object).Count + 1
If($VMDataLastRow -gt 1){
    $VMDataHeaderCount = Get-ColumnName ($VMData | Get-Member | Where-Object{$_.MemberType -match 'NoteProperty'} | Measure-Object).Count
    $VMDataHeaderRow = "'VMs'!`$A`$1:`$$VMDataHeaderCount`$1"
    $VMCPUsColumn = "'VMs'!`$B`$2:`$B`$$VMDataLastRow"
    $VMRAMGBColumn = "'VMs'!`$C`$2:`$C`$$VMDataLastRow"
    $VMDisksColumn = "'VMs'!`$D`$2:`$D`$$VMDataLastRow"
    $VMUsedRawColumn = "'VMs'!`$E`$2:`$E`$$VMDataLastRow"
    $VMUsedGBColumn = "'VMs'!`$F`$2:`$F`$$VMDataLastRow"
    $VMCapacityRawColumn = "'VMs'!`$G`$2:`$G`$$VMDataLastRow"
    $VMCapacityGBColumn = "'VMs'!`$H`$2:`$H`$$VMDataLastRow"
    $VMCPUCostColumn = "'VMs'!`$K`$2:`$K`$$VMDataLastRow"
    $VMRAMCostColumn = "'VMs'!`$L`$2:`$L`$$VMDataLastRow"
    $VMStorageCostColumn = "'VMs'!`$M`$2:`$M`$$VMDataLastRow"
    $VMFlatCostColumn = "'VMs'!`$N`$2:`$N`$$VMDataLastRow"
    $VMMonthlyCostColumn = "'VMs'!`$O`$2:`$O`$$VMDataLastRow"

    $VMDataStyle = @()
    $VMDataStyle += New-ExcelStyle -Range $VMDataHeaderRow -HorizontalAlignment Center
    $VMDataStyle += New-ExcelStyle -Range $VMCPUsColumn -NumberFormat '0'
    $VMDataStyle += New-ExcelStyle -Range $VMRAMGBColumn -NumberFormat '0'
    $VMDataStyle += New-ExcelStyle -Range $VMDisksColumn -NumberFormat '0'
    $VMDataStyle += New-ExcelStyle -Range $VMUsedRawColumn -NumberFormat '0'
    $VMDataStyle += New-ExcelStyle -Range $VMUsedGBColumn -NumberFormat '0.00'
    $VMDataStyle += New-ExcelStyle -Range $VMCapacityRawColumn -NumberFormat '0'
    $VMDataStyle += New-ExcelStyle -Range $VMCapacityGBColumn -NumberFormat '0.00'
    $VMDataStyle += New-ExcelStyle -Range $VMCPUCostColumn -NumberFormat 'Currency'
    $VMDataStyle += New-ExcelStyle -Range $VMRAMCostColumn -NumberFormat 'Currency'
    $VMDataStyle += New-ExcelStyle -Range $VMStorageCostColumn -NumberFormat 'Currency'
    $VMDataStyle += New-ExcelStyle -Range $VMFlatCostColumn -NumberFormat 'Currency'
    $VMDataStyle += New-ExcelStyle -Range $VMMonthlyCostColumn -NumberFormat 'Currency' -Bold

    $VMDataConditionalFormatting = @()

    $VMData | Sort-Object 'Cluster', 'Name' | Export-Excel @ExcelProps -WorksheetName 'VMs' -Style $VMDataStyle -ConditionalFormat $VMDataConditionalFormatting
}

# Error sheet
$vCenterErrorLastRow = ($vCenterError | Measure-Object).Count + 1
If($vCenterErrorLastRow -gt 1){
    $vCenterErrorHeaderCount = Get-ColumnName ($vCenterError | Get-Member | Where-Object{$_.MemberType -match 'NoteProperty'} | Measure-Object).Count
    $vCenterErrorHeaderRow = "Errors!`$A`$1:`$$vCenterErrorHeaderCount`$1"

    $vCenterErrorStyle = @()
    $vCenterErrorStyle += New-ExcelStyle -Range $vCenterErrorHeaderRow -HorizontalAlignment Center

    $vCenterError | Sort-Object 'Object', 'Name' | Export-Excel @ExcelProps -WorksheetName 'Errors' -Style $vCenterErrorStyle
}
#endregion
