<#
===============================================================================================================================================================
In Place Upgrade - Policy Engine
===============================================================================================================================================================
Author:     Brian Thorp
Date:       2018-12-19
Updated:    2018-12-04
===============================================================================================================================================================
#>

# Log File
$Global:OutPutFile                  =   "C:\SCCM\IPU\ExistingUserReg.log"           # Text file to log to

# Items in C:\Users Excluded
$Global:Exclusionsfile              =   ".\exclusions.csv"                          # Exluded folders / user profiles~

# Instruction File
$CSV_Path                           =   ".\KeyData.csv"

# 
$Global:IncludeDefaultUser          =   $true                                       # Determines if we should include C:\Users\Default\ in our profiles

$UserDir                            =   "C:\Users\"


# Logging function
function Out-Logger
{
    [cmdletbinding(SupportsShouldProcess=$True)]
    Param(
        [string]$Message
    )


    # Create the file if it doesnt exist
    $fold = Split-Path $Global:OutPutFile -Parent
    
    if (!(test-path -path $fold))
    {
        New-Item -Path $fold -ItemType Container
    }
    
    
    # Create the file if it doesnt exist
    if (!(test-path -path $Global:OutPutFile))
    {
        New-Item -Path $Global:OutPutFile -ItemType File
    }

    # write to console
    Write-Host "Log Message: $Message"

    # write to log file, append
    # Out-File -FilePath $Global:OutPutFile $Message -Append
    $message | Out-File -FilePath $Global:OutPutFile -Append
}

# Function just kicks off some of the logging~
function Start-Script
{
    [cmdletbinding(SupportsShouldProcess=$True)]
    $date = Get-Date

    Out-Logger -Message "***************************************************************************************************************************************************************"
    Out-Logger -Message "|                                                                      Script Start                                                                           |"
    Out-Logger -Message "***************************************************************************************************************************************************************"
    Out-Logger -Message "Start Time: 													$date"
    Out-Logger -Message " "

    return $date
}

# Carbon Registry Test Module Function
# https://stackoverflow.com/questions/5648931/test-if-registry-value-exists
function Test-RegistryKeyValue
{
    <#
    .SYNOPSIS
    Tests if a registry value exists.

    .DESCRIPTION
    The usual ways for checking if a registry value exists don't handle when a value simply has an empty or null value.  This function actually checks if a key has a value with a given name.

    .EXAMPLE
    Test-RegistryKeyValue -Path 'hklm:\Software\Carbon\Test' -Name 'Title'

    Returns `True` if `hklm:\Software\Carbon\Test` contains a value named 'Title'.  `False` otherwise.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]
        # The path to the registry key where the value should be set.  Will be created if it doesn't exist.
        $Path,

        [Parameter(Mandatory=$true)]
        [string]
        # The name of the value being set.
        $Name
    )

    if( -not (Test-Path -Path $Path -PathType Container) )
    {
        return $false
    }

    $properties = Get-ItemProperty -Path $Path 
    if( -not $properties )
    {
        return $false
    }

    $member = Get-Member -InputObject $properties -Name $Name
    if( $member )
    {
        return $true
    }
    else
    {
        return $false
    }
}

# Function Add
function New-Registry
{
    [cmdletbinding(SupportsShouldProcess=$True)]
    Param
    (
        $key,
        $name,
        $regtype,
        $value,
        $description,
        $prefix
    )

    # Full Key Path
    $ps_key = $prefix + $key


    Out-Logger -Message "Testing for existing registry path:"
    if (Test-RegistryKeyValue -Path $ps_key -Name $name)
    {
        Out-Logger -Message "   Key found, modifying..."
        Set-ItemProperty -Path $ps_key -Name $name -Value $value
    }
    else
    {
        Out-Logger -Message "Key not found..."

        $Key_Exists = Test-Path -Path "$ps_key"

        Out-Logger -Message "Testing if key exists: $Key_Exists"

        if (!$Key_Exists)
        {
            Out-Logger -Message "Key doesnt exist, creating..."
            New-Item -Path $ps_key -Force | Out-Null
        }

        New-ItemProperty -Path $ps_key -Name $name -Value $value -PropertyType $regtype
    }

    Out-Logger -Message "Key Changed Completed."
    Out-Logger -Message " "
    
    Out-Logger -Message "-------------------------------------------------------------"
    Out-Logger -Message "Begin validation of change"
    Out-Logger -Message "-------------------------------------------------------------"

    $ps_keyroot = Get-ItemProperty -Path "$ps_key"
    $FPath = ($ps_keyroot.$name)

    # Validation
    Out-Logger -Message "Key: $ps_key"
    Out-Logger -Message "Checking value: $FPath"
    if( $value -eq $FPath )
    {
        Out-Logger -Message "$Description is set properly."
    }
    else
    {
        Out-Logger -Message "$Description FAILED to set properly."
        throw "Configuration Change Issue Detected @ $Description"
    }
    Out-Logger -Message "-------------------------------------------------------------"
    Out-Logger -Message " "

    # Not supported on PS 2.0
    #$regresult.Handle.Close()  
}

# Function Delete
function Remove-Registry
{
    [cmdletbinding(SupportsShouldProcess=$True)]
    Param
    (
        $key,
        $name,
        $regtype,
        $value,
        $description,
        $prefix
    )

    $ps_key = $prefix + $key

    Out-Logger -Message "-------------------------------------------------------------"
    Out-Logger -Message "Begin registry change -- Remove"
    Out-Logger -Message "Key:           $ps_key"
    Out-Logger -Message "Name:          $name"
    Out-Logger -Message "Type:          $regtype"
    Out-Logger -Message "Value:         $value"
    Out-Logger -Message "Description:   $description"
    Out-Logger -Message "-------------------------------------------------------------"
    # =======================================================================================================================
    # Check if we're just deleting the root key, or deleting a property. Null = Whole Key
    #if ($null -eq $name)
    $DeleteType = [string]::IsNullOrEmpty($name)

    if ($DeleteType)
    {
        Out-Logger -Message "Testing for existing registry key path..."
        if (Test-Path -Path $ps_key)
        {
            Out-Logger -Message "Key found, deleting..."
            Remove-Item -Path $ps_key -Recurse

            $rootkeydelete = $true
        }
        else
        {
            Out-Logger -Message "Key not found - nothing to do."
        }
    }
    else # Deleting a property
    {

        Out-Logger -Message "Testing for existing registry path..."
        if (Test-RegistryKeyValue -Path $ps_key -Name $name)
        {
            Out-Logger -Message "Key found, deleting..."
            Remove-ItemProperty -Path $ps_key -Name $name
        }
        else
        {
            Out-Logger -Message "Property not found - nothing to do."
        }
    }

    # =======================================================================================================================
    Out-Logger -Message "Key Changed Completed."
    Out-Logger -Message " "
    Out-Logger -Message "-------------------------------------------------------------"
    Out-Logger -Message "Begin validation of change"
    Out-Logger -Message "-------------------------------------------------------------"

    if (!$DeleteType)
    {
        $ps_keyroot = Get-ItemProperty -Path "$ps_key"
        $FPath = ($ps_keyroot.$name)
    }
    

    # Validation
    # Value is NULL if deleted - note that this is currently untested
    Out-Logger -Message "Key: $ps_key"
    Out-Logger -Message "Checking value: $FPath"

    if( $null -eq $FPath )
    {
        Out-Logger -Message "$Description deleted properly."
    }
    else
    {
        Out-Logger -Message "$Description FAILED to delete properly."
        throw "Configuration Change Issue Detected"
    }

    if ($rootkeydelete)
    {
        Out-Logger -Message "$Description - Testing for key deletion"
        $isDeleted = Test-Path -Path $ps_key

        if (!$isDeleted)
        {
            Out-Logger -Message "$Description - Success"
        }
    }

    # Doesnt work on PS 2.00~
    # $regresult.Handle.Close()  
}

# Get the user profiles on the system, return it
function Get-UserProfiles
{
    Param(
        [string]$UsersDir
    )

    $userlist = $null
    $userlist = New-Object System.Collections.Generic.List[System.Object]

    # Get a list of the user profiles
    $users = Get-ChildItem -Path $UsersDir -Force

    # Load the Exclusions List
    $excludedprofiles = Import-CSV $Global:Exclusionsfile
    # Get Just the profile names from the array
    $exclusions = $excludedprofiles.name
    # Allow us to edit the array if we need to add or remove
    $exclusions = [System.Collections.ArrayList]$exclusions

    # If We arent including default user, add it to the exceptionslist
    if (!($Global:IncludeDefaultUser -eq $true))
    {
        $exclusions.Add("Default") | Out-Null   # Out null to absorb the function outputting array index on add
    }

    # Cycle through the users
    ForEach($user in $users)
    {
        # Exclude the ones we want to ignore
        if ($exclusions -notcontains [system.string]$user)
        {
            # Add our filtered user to our output
            $userlist.add("$user")
        }
    }
    return $userlist
}

# Loads ntuser.dat files into the registry hive for our editing
function Start-UserHive
{
    [cmdletbinding(SupportsShouldProcess=$True)]
    Param
    (
        $prefix,
        $UsersDir
    )

    # Since we are weird we're going to use CMD tools and they dont like :
    $reg_prefix = $prefix -replace '[:]'

    & cmd.exe /c reg.exe load $reg_prefix $UsersDir + $user\ntuser.dat

    return 0;
}

function Stop-UserHive
{
    [cmdletbinding(SupportsShouldProcess=$True)]
    Param
    (
        $prefix
    )

    # This hard tries to close out open registry in the default prefix path w/o inputs or outputs
    [gc]::Collect()

    if ( Test-Path $prefix )
    {
        # Since we are weird we're going to use CMD tools and they dont like :
        $reg_prefix = $prefix -replace '[:]'

        & cmd.exe /c reg.exe unload $reg_prefix
    }    

    return 0;
}



# Import CSV File with instructions
$csv = Import-CSV "$CSV_Path"

# Row (instruction) -> Execute per scope and return results

# Loop through CSV Instructions and branch >
$starttime = Start-Script
$users = Get-UserProfiles -UsersDir $UserDir


# Main Function Body
ForEach($row in $csv)
{
    $import_mode            =   $row.mode
    $import_zone            =   $row.zone
    $import_function        =   $row.function       # Determines what function to use
    $import_path            =   $row.path
    $import_name            =   $row.Name
    $import_type            =   $row.type
    $import_value           =   $row.value
    $import_description     =   $row.description    # For logging 

    # Flip type to PowerShell
    switch ($import_type)
    {
        REG_Binary          { $ps_type = "Binary"       }
        REG_DWORD           { $ps_type = "DWord"        }
        REG_EXPAND_SZ       { $ps_type = "ExpandString" }
        REG_MULTI_SZ        { $ps_type = "MultiString"  }
        REG_QWORD           { $ps_type = "QWord"        }
        REG_SZ              { $ps_type = "String"       }
        $null               { $ps_type = $null          }
    }
    # Out-Logger -Message "Registry type changed from $import_type to $ps_type"

    switch ($import_mode)
    {
        registry
        {
            #write-host "Registry"
            switch ($import_zone)
            {
                user
                {
                    switch ($import_function)
                    {
                        Add     
                        {
                            Out-Logger -Message "==============================================================================================================================================================="
                            Out-Logger -Message "Begin registry change -- Add"
                            Out-Logger -Message "==============================================================================================================================================================="
                            Out-Logger -Message "Key:           $ps_key"
                            Out-Logger -Message "Name:          $name"
                            Out-Logger -Message "Type:          $regtype"
                            Out-Logger -Message "Value:         $value"
                            Out-Logger -Message "Description:   $description"
                            Out-Logger -Message "---------------------------------------------------------------------------------------------------------------------------------------------------------------"
                            Out-Logger -Message " "
                            
                            # Since this is a user hive we'll use our temp prefix
                            $prefix = "HKLM:\UserLand"

                            # For Each User
                            foreach($user in $users)
                            {

                                Out-Logger -Message "Current User:  $user"
                                
                                # Load the User Hive
                                Out-Logger -Message "   Loading Registry"
                                Start-UserHive -prefix $prefix -UsersDir $UserDir
                                
                                # Make Registry changes or additions
                                Out-Logger -Message "   Applying Change"
                                New-Registry -key $import_key -name $import_name -regtype $ps_type -value $import_value -description $import_description -prefix $prefix
                                
                                # Unload the user hive
                                Stop-UserHive -prefix $prefix
                            }

                        }
                        Delete
                        {
                            $prefix = "HKLM:\UserLand"

                            foreach($user in $users)
                            {
                                
                                # Load the User Hive
                                Start-UserHive -prefix $prefix -UsersDir $UserDir
                                
                                # Make Registry changes or additions
                                Remove-Registry -key $import_path -name $import_name -regtype $ps_type -value $import_value -description $import_description -prefix $prefix
                                
                                # Unload the user hive
                                Stop-UserHive -prefix $prefix
                            }
                        }
                    }
                }
                global # Global System-Wide Registry Changes
                {
                    switch ($import_function)
                    {
                        Add
                        {
                            $prefix = "HKLM:"

                            New-Registry -key $import_path -name $import_name -regtype $ps_type -value $import_value -description $import_description -prefix $prefix
                        }
                        Delete
                        {
                            $prefix = "HKLM:"

                            Remove-Registry -key $import_path -name $import_name -regtype $ps_type -value $import_value -description $import_description -prefix $prefix
                        }
                    }
                }
            }
        }
        file
        {
            switch ($import_zone)
            {
                Global
                {
                    write-host "File"
                    switch ($import_function)
                    {
                        Add
                        {
                            <# This function is currently not supported at all #>
                        }
                        Delete
                        {
                            # Delete the defined files
                        }
                    }
                }
                User
                {
                    write-host "File"
                    switch ($import_function)
                    {
                        Add
                        {
                            <# This function is currently not supported at all #>
                        }
                        Delete
                        {
                            ForEach($user in $users)
                            {
                                # Used for example to purge out shortcuts or pinned items on Windows 7 task bar prior to upgrade
                                # C:\Users | %Username% | <PATH>
                                # TaskPin V10
                                # Migrate this to a delete function~ Remove-Files -Path $Import_Path -User 1 

                                Test-Path -Path $import_path
                                Remove-Item $FullPin -Recurse -confirm:$false -ErrorAction SilentlyContinue
                            }
                            # Delete the defined files
                        }
                }

                }
            }
        }
    }
}