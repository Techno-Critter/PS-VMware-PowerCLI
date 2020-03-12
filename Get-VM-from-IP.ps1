<#
Author: Stan Crider
Date: 10Apr2017
What this crap does:
Take user-input IP address and search specified vCenter for virtual machines with that IP.

!!! Requirements for this script to work: !!!
1. VMware PowerCLI must be installed on the computer running this script
2. VMware Tools must be installed on the VM for this script to recognize the IP address
3. Permissions on specified vCenter to read VM properties
#>

# Specify vCenter and DataCenter
$vCenter = "VC-Server"

# Initialize counter
$FindCounter = 0

# Get user-input IP address
$VIP = Read-Host "Enter IP address to search for or 'exit' to cancel"

If($VIP -eq "exit"){
# Exit script if cancelled
    Write-Output "Script cancelled."
    Break
}
Else{
# Validate IP Address format
    $ip_address = $null
    [System.Net.IPAddress]::TryParse($VIP, [ref]$ip_address) | Out-Null
    If($null -ne $ip_address){

# Import PowerCLI Module and connect to vCenter server
        Write-Output ("Connecting to vCenter " + $vCenter + ". . .")
        Import-Module VMware.vimautomation.core
        Connect-VIserver $vCenter

# Get virtual machines
        $VSrvrs = Get-VM
#        $VSrvrs = Get-DataCenter $VMDataCenter | Get-VM
        Write-Output ("Searching for IP address " + $VIP + ". . .")

# Find each IP address of each VM; report name and increment counter upon match
        ForEach($VSrv in $VSrvrs){
            $VMIPList = $VSrv.Guest.IPAddress
            ForEach($VMIP in $VMIPList){
                If($VMIP -eq $VIP){
                    Write-Output ($VSrv.Name + ": " + $VMIP)
                    $FindCounter++
                }
            }
        }

# Report number of matches
        If ($FindCounter -eq 0){
            Write-Warning ("No VM's found with IP address " + $VIP)
        }
        Else{
            Write-Output ("VM's found with IP address " + $VIP + ": " + $FindCounter)
        }

# Close vCenter connection
        Disconnect-VIserver -Confirm:$false
    }
# If IP format not valid, exit script
    Else{
        Write-Error ("The entry " + $VIP + " is not a valid IP Address. Script terminated.")
        Break
    }
}
