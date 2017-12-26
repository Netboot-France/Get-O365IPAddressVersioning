<#  
    .SYNOPSIS  
        This script provides a list of Office365 IP address changes

    .DESCRIPTION
        This script gets the Office 365 IP address and URL, in XML format, and returns the changes in CSV format,
        you can filter the products you want with the configuration file "Config.json"

    .NOTES  
        File Name   : Get-O365IPAddressVersioning.ps1
        Author      : Thomas ILLIET, contact@thomas-illiet.fr
        Date	    : 2017-10-23
        Last Update : 2017-10-23
        Test Date   : 2017-10-23
        Version	    : 1.0.0 

    .LINK
        Invoke-WebRequest
        https://support.office.com/en-us/article/Office-365-URLs-and-IP-address-ranges-8548a211-3fe7-47cb-abb1-355ea5aa88a2
        https://support.content.office.net/en-us/static/O365IPAddresses.xml
#>

# Paramaters #######################################################################################
[CmdletBinding()]
Param (
    [Parameter(Mandatory=$False)]
    [ValidateSet('Continue','SilentContinue')]
    $DebugPreference = "Continue"
)
# END Paramaters ###################################################################################

Clear-Host

# GLOBAL VARIABLES #################################################################################
$BaseName       = "O365IPAddress"
$LogTime        = Get-Date -Format "yyyy.MM.dd_HH.mm.ss"
$ExportFile    = Join-Path -Path $PSScriptRoot -ChildPath "Export/$BaseName-$LogTime.csv"
$DatabaseFile   = (Join-Path -Path $PSScriptRoot -ChildPath "Config/Database.json")
$ConfigFile     = (Join-Path -Path $PSScriptRoot -ChildPath "Config/Config.json")
# END GLOBAL VARIABLES #############################################################################

# FUNCTIONS ########################################################################################
Function Get-O365IPAddress {
    <#  
        .SYNOPSIS  
            The function gets the Office 365 IP address & URL information.

        .DESCRIPTION
            The function gets the Office 365 IP address & URL information, in XML format, and returns the Product (O365, SPO, etc...),
            you can filter the products you want with the configuration file "Config.json"

        .NOTES  
            File Name   : Get-O365IPAddress.ps1
            Author      : Thomas ILLIET, contact@thomas-illiet.fr
            Date	    : 2017-10-23
            Last Update : 2017-10-23
            Test Date   : 2017-10-23
            Version	    : 1.0.0

        .LINK
            Invoke-WebRequest
            https://support.office.com/en-us/article/Office-365-URLs-and-IP-address-ranges-8548a211-3fe7-47cb-abb1-355ea5aa88a2
            https://support.content.office.net/en-us/static/O365IPAddresses.xml
    #>
    Try
    {
        Write-Verbose "Getting the Office 365 IP address & URL information from $O365IPAddresses."
        $O365IPAddresses = "https://support.content.office.net/en-us/static/O365IPAddresses.xml"
        [XML]$O365XMLData = Invoke-WebRequest -Uri $O365IPAddresses -DisableKeepAlive -ErrorAction Stop

        $O365IPAddressObj = @() 
        ForEach($Product in $O365XMLData.Products.Product)
        {
          ForEach($AddressList in $Product.AddressList)
          {
            $O365IPAddressProps = @{
                Product     = $Product.Name;
                AddressType = $AddressList.Type;
                Addresses   = $AddressList.Address
            }
            $O365IPAddressObj += New-Object -TypeName PSObject -Property $O365IPAddressProps
          }
        }
        return $O365IPAddressObj
    }
    Catch
    {
      Write-Error "Failed to get the Office 365 IP address & URL information from $O365IPAddresses."
    }
}
# END FUNCTIONS ####################################################################################

# MONITOR ##########################################################################################

# Start Watch
$sw = New-Object Diagnostics.Stopwatch
$sw.Start()

# END MONITOR ######################################################################################


# MAIN #############################################################################################

#----------------------------------------------
# Get Configuration
#-----------------------------------------------
$Config         = Get-Content $ConfigFile | ConvertFrom-Json
$ALLO365Address = Get-O365IPAddress
if (Test-Path -Path $DatabaseFile) {
    $LocalDatabase  = Get-Content $DatabaseFile | ConvertFrom-Json
}

#----------------------------------------------
# Get 365 Database
#----------------------------------------------
write-debug "| + Get 365 Database"
$365Database = @()
foreach($O365Address in $ALLO365Address){
    
    # If AddressType is enable on product
    if($Config.($O365Address.Product).($O365Address.AddressType) -eq $True) {
   
        foreach($Addresses in $O365Address.Addresses -split ","){
            $365Database += [PSCustomObject]@{
                Address     = $Addresses
                AddressType = $O365Address.AddressType
                Product   = $O365Address.Product
            }
        }
    } elseif($Config.($O365Address.Product).($O365Address.AddressType) -ne $False){
        Write-Warning "$($O365Address.AddressType) Is not configured on $($O365Address.Product)"
    }
}

if($LocalDatabase -ne $null) {

    #----------------------------------------------
    # Compare 365 Database & Local Database
    #----------------------------------------------
    write-debug "| + Compare 365 Database & Local Database"
    $CompareObject = Compare-Object -ReferenceObject $LocalDatabase -DifferenceObject $365Database -Property Address,AddressType,Product

    #----------------------------------------------
    # Create Export
    #----------------------------------------------
    write-debug "| + Create Export"
    $Export = @()
    foreach($Object in $CompareObject){

        if($Object.SideIndicator -eq "=>") 
        { 
            $lineOperation = "To added" 
        } 
        elseif($Object.SideIndicator -eq "<=") 
        { 
            $lineOperation = "To deleted" 
        }

        $Export += [PSCustomObject][ordered] @{
            Alias       = $Config.($Object.Product).Alias
            Product     = $Object.Product
            AddressType = $Object.AddressType
            Address     = $Object.Address
            Operation   = $lineOperation
        } 
    }
    if($Export -ne $null) {
        $Export | Export-Csv -Path $ExportFile -Encoding UTF8 -NoTypeInformation -Delimiter ";"
    } else {
        Write-Debug "|  + No change Detected"
    }

} else {

    $Export = @()
    foreach($Object in $365Database){
        $Export += [PSCustomObject][ordered] @{
            Alias       = $Config.($Object.Product).Alias
            Product     = $Object.Product
            AddressType = $Object.AddressType
            Address     = $Object.Address
            Operation   = "To added"
        }
    }
    $Export | Export-Csv -Path $ExportFile -Encoding UTF8 -NoTypeInformation -Delimiter ";"
}

if($Export -ne $null) {
    #----------------------------------------------
    # Update Local Database
    #----------------------------------------------
    write-debug "| + Update Local Database"
    $365Database | ConvertTo-Json | Set-Content $DatabaseFile
}

# END MAIN #########################################################################################

# MONITOR ##########################################################################################

# Stop Watch
$sw.Stop()
$TotalSeconds = [math]::Round(($sw.Elapsed).TotalSeconds)
write-host "Script executed on $TotalSeconds secondes" -ForegroundColor Yellow

# END MONITOR ######################################################################################