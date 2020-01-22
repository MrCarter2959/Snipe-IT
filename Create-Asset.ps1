#################################################
# Notes
# 
# - Need To Build GUI For User Prompts
# - Need To Build GUI For User Assignment
# - Need To Figure Out obtaining winPE mac addresses
# - Need to tweak and build a little more error capturing and try/catch loops
# - In SnipeIT You need to create Custom Fields For Wireless MAC Address and Ethernet MAC Addresses
# - To run in WinPE you need the SnipeIT POSH Module LINK = https://github.com/snazy2000/SnipeitPS embedded
# - Also to run in WinPE you need to embed PowerShell in the boot image
Function Check-Manufacturer {

    param (
        [Parameter (Mandatory = $true)]
        [string] $URL,
        [string] $APIKey,
        [string] $Search
    )
    
    $Details = New-Object System.Collections.Generic.List[object]

    Try {
            $manufacturerExists = Get-Manufacturer -url $URL -apiKey $APIKey -search $Search | Select Name,ID
        }
    Catch
        {
            Write-Host $_.Exception.Message
        }

    If ($manufacturerExists -eq $null)
        {
            $manufacturerName = ($manufacturerExists.Name)
            $manufacturerID = ($manufacturerExists.ID)
            Write-Host "Manufacturer : $manufacturerName Doesn't Exist" -BackgroundColor "Black" -ForegroundColor "Yellow"
        }
    Else
        {
            $manufacturerName = ($manufacturerExists.Name)
            $manufacturerID = ($manufacturerExists.ID)
            #Write-Host "Manufacturer : $Manufacturer Already Exists" -BackgroundColor "Black" -ForegroundColor "Yellow"
             $Details.Add(
                [PSCUSTOMOBJECT]@{ManufactuerName=$manufacturerName ; ManufactuerID=$manufacturerID}
            )
            $manufacturerArray = Tee-Object -Variable 'manufactuerList' -InputObject $Details | Format-Table -Wrap
            Return $manufactuerList
        }
    }

Function Check-Model {
    
    param (
        [Parameter (Mandatory = $true)]
        [string] $URL,
        [string] $APIKey,
        [string] $Search
    )
    
    $ModelDetails = New-Object System.Collections.Generic.List[object]

    Try {
        $modelExists = Get-Model -url $url -apiKey $APIKey -search $Search | Select Name , ID
        }
    Catch
        {
            Write-Host $_.Exception.Message
        }

    If ($modelExists -eq $null)
        {
            $modelName = ($modelExists.Name)
            $modelID = ($modelExists.ID)
            Write-Host "Model : $ComputerSystem Doesn't Exist" -BackgroundColor "Black" -ForegroundColor "Yellow"
            #Return $modelName , $modelID
        }
    Else
        {
            $modelName = ($modelExists.Name)
            $modelID = ($modelExists.ID)
            Write-Host "Model : $($modelExists.Name) Already Exists" -BackgroundColor "Black" -ForegroundColor "Yellow"
            $ModelDetails.Add(
                [PSCUSTOMOBJECT]@{ModelName=$modelName ; ModelID=$ModelID}
            )
            Return $ModelDetails
        }
    }

Function Check-Asset {

    param (
        [Parameter (Mandatory = $true)]
        [string] $URL,
        [string] $APIKey,
        [string] $Search
    )
    
    $AssetDetails = New-Object System.Collections.Generic.List[object]

    Try {
            $checkAsset = Get-Asset -url $url -apiKey $APIKey -search $Search | Select Name , ID
        }
    Catch
        {
            Write-Host $_.Exception.Message
        }

    if ($checkAsset -eq $null)
        {
            $assetName = ($modelExists.Name)
            $assetID = ($modelExists.ID)
            Write-Host "Asset : $Search Doesn't Exist" -BackgroundColor "Black" -ForegroundColor "Yellow"
        }
    Else
        {
            $assetName = ($checkAsset.Name)
            $assetID = ($checkAsset.ID)
            Write-Host "Asset : $($checkAsset.Name) Already Exists" -BackgroundColor "Black" -ForegroundColor "Yellow"
            $AssetDetails.Add(
                [PSCUSTOMOBJECT]@{AssetName=$assetName ; AssetID=$assetID}
            )
            Return $AssetDetails
        }
    }

Function Detect-Laptop {
        $isLaptop = $false
        #The chassis is the physical container that houses the components of a computer. Check if the machine’s chasis type is 9.Laptop 10.Notebook 14.Sub-Notebook
        
        if (Get-WmiObject -Class win32_systemenclosure | Where-Object { $_.chassistypes -eq 9 -or $_.chassistypes -eq 10 -or $_.chassistypes -eq 14})
            {
                $isLaptop = $true
            }
        Else
            {
                $isLaptop = $false
            }

        #Shows battery status , if true then the machine is a laptop.
        
        if(Get-WmiObject -Class win32_battery)
            {
                $isLaptop = $true
            }

    $isLaptop
}

###################################
#NOTE: Work in progress, need to build GUI for prompts at beginning and end
#To ask if computer is new, or replacing another one in inventory
#
###################################
#                                 #
#           Set Variables         #
#                                 #
###################################
#URL
$URL = "server_URL"
#API Key
$secret = "api-key"


$SerialNumber = get-ciminstance win32_bios | Select -ExpandProperty serialnumber
###################################
# This works, but trying to get syntax correct for all ethernet cards
$HardMACAddress = Get-NetAdapter -name "ethernet" | Select -ExpandProperty  MacAddress
#
$WirelessMACAddress = Get-WmiObject win32_networkadapterconfiguration | Select description, macaddress | where {$_ -like "*Wireless*"}
#
$ComputerName = Get-WmiObject Win32_computerSystem | Select -ExpandProperty Name
#
$ComputerModel = Get-WmiObject win32_computersystemproduct | Select -ExpandProperty Name
#
$ComputerModelNumber = $ComputerModel.Split(" ")[1]
#
$ComputerManufacturer = Get-WmiObject Win32_ComputerSystem | Select -ExpandProperty Manufacturer
#
$ComputerSystem = Get-WmiObject Win32_ComputerSystem | Select -ExpandProperty SystemFamily

###################################
#                                 #
#       Check For Existing        #
#                                 #
###################################

#Asset
$check_asset = Check-Asset -URL $url -APIKey $secret -Search $ComputerName

#Manufacturer
$check_Manufacturer = Check-Manufacturer -URL $url -APIKey $secret -Search $ComputerManufacturer

#Model Name
$check_model_Name = Check-Model -URL $url -APIKey $secret -Search $ComputerSystem

#Model Number
$check_Model_number = Check-Model -URL $url -APIKey $secret -Search $ComputerModelNumber

###################################
#                                 #
#      Create Non Existing        #
#                                 #
###################################

#Manufacturer
If ($check_Manufacturer -eq $null)
    {
        New-Manufacturer -url $URL -apiKey $secret -Name $ComputerManufacturer
    }
#Returns data for other checks below
$check_Manufacturer = Check-Manufacturer -URL $url -APIKey $secret -Search $ComputerManufacturer

#Model / Model Number
If ($check_Model_number -eq $null)
    {
        
        If (Detect-Laptop)
            {
                Write-host “Asset : Laptop Detected” -BackgroundColor "Black" -ForegroundColor "Yellow"
                $categoryType = "4"
            }
        Else 
            {
                Write-host “Asset : Desktop Detected” -BackgroundColor "Black" -ForegroundColor "Yellow"
                $categoryType = "2"
            }
        New-Model -url $URL -apiKey $secret -name $ComputerSystem -model_number $check_Model_number -fieldset_id 1 -manufacturer_id $($check_Manufacturer.ManufactuerID) -category_id $categoryType
    }
#Returns Data for other checks below
$check_Model = Check-Model -URL $url -APIKey $secret -Search $ComputerModelNumber | Select ModelID

###################################
#                                 #
#          Create Asset           #
#                                 #
###################################
If ($check_asset -eq $null)
    {
        $new_asset = New-Asset -url $url -apiKey $secret -Name $ComputerName -Model_id $($check_Model.ModelID) -Status_id 4 -customfields @{
        '_snipeit_ethernet_mac_address_3' = $HardMACAddress
        'serial' = $SerialNumber
        }
    }
Else
    {
        Write-Host "Computer : $ComputerName Already exists" -BackgroundColor "Black" -ForegroundColor "Yellow"
    }
###################################
#                                 #
#          Assign Asset           #
#                                 #
###################################
$assignedAsset = Get-Asset -search $ComputerName -url $url -apiKey $secret | Select name, assigned_to ,id
#-id = $assignedAsset.id
#-assigned_id = -assigned_id "1"
if ($assignedAsset.assigned_to -eq $null)
    {
        Write-Host "$assignedAsset is not assigned"
        Set-AssetOwner -id $assignedAsset.id -assigned_id "1" -checkout_to_type user -url $url -apiKey $secret -Verbose

    }
Else
    {
        Write-Host "$($assignedAsset.name) is assigned to $($assignedAsset.assigned_to.name)"
    }
