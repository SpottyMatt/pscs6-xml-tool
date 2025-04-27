<#
.SYNOPSIS
    Updates the BridgeTalkVersion value in the Adobe Photoshop CS6 application.xml file.

.DESCRIPTION
    This script modifies the Adobe Photoshop CS6 application.xml file to increment the BridgeTalkVersion value.
    It temporarily disables network connectivity during the update to prevent license validation checks.
    The script will automatically request administrator privileges if needed.

.PARAMETER XmlFilePath
    Full path to the Adobe application.xml file. Defaults to the standard installation path.

.PARAMETER LogFile
    Path to the log file. Defaults to update_log.txt in the script directory.

.PARAMETER DataKey
    The specific data key to locate in the XML. Defaults to "BridgeTalkVersion".

.PARAMETER AdobeCode
    The adobeCode attribute value to locate in the XML. Defaults to "/adobe/bridgetalk/photoshop-60.064".

.PARAMETER Elevated
    Internal parameter used when the script elevates itself. Should not be specified manually.

.NOTES
    Version:        1.1
    Author:         Script Author
    Creation Date:  2025-04-26
#>
param(
    [Parameter(Position = 0)]
    [string]$XmlFilePath = "$env:ProgramFiles\Adobe\Adobe Photoshop CS6 (64 Bit)\AMT\application.xml",
    
    [Parameter(Position = 1)]
    [string]$LogFile = "$PSScriptRoot\update_log.txt",
    
    [Parameter(Position = 2)]
    [string]$DataKey = "BridgeTalkVersion",
    
    [Parameter(Position = 3)]
    [string]$AdobeCode = "/adobe/bridgetalk/photoshop-60.064",
    
    [Parameter(Position = 4)]
    [switch]$Elevated
)

# Clear log file only at the start of the non-elevated process
if (-not $Elevated) {
    try {
        # Clear the log file or create a new one
        Set-Content -Path $LogFile -Value "Log started at $(Get-Date)" -Force
    } catch {
        # If we can't write to the log file, use a temporary one
        $LogFile = "$env:TEMP\update_xml_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
        Set-Content -Path $LogFile -Value "Log started at $(Get-Date)" -Force
    }
}

# Helper function for logging with better file handling
function Write-Log {
    <#
    .SYNOPSIS
        Writes a log message to both the console and a log file.
    
    .DESCRIPTION
        This function writes a timestamped log message to both the console and a log file.
        It handles errors gracefully and tries alternative methods to write to the log file if needed.
    
    .PARAMETER Message
        The message to log. Can be empty.
    
    .PARAMETER IsError
        If specified, the message is treated as an error message.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true)]
        [string]$Message = "",
        
        [Parameter()]
        [switch]$IsError
    )
    
    # Generate timestamp
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Format message with optional ERROR prefix
    if ($IsError) {
        $logMessage = "[$timestamp] ERROR: $Message"
        # Write to error stream for console
        Write-Error $Message
    } else {
        $logMessage = "[$timestamp] $Message"
        # Write to standard output for console
        Write-Host $logMessage
    }
    
    # Always write to log file with better error handling
    try {
        Add-Content -Path $LogFile -Value $logMessage -ErrorAction Stop
    } catch {
        # Try opening the file with shared access
        try {
            $fileStream = [System.IO.File]::Open($LogFile, 'Append', 'Write', 'ReadWrite')
            $streamWriter = New-Object System.IO.StreamWriter($fileStream)
            $streamWriter.WriteLine($logMessage)
            $streamWriter.Close()
            $fileStream.Close()
        } catch {
            # If we still can't write, log to console only
            Write-Host "WARNING: Could not write to log file: $_"
        }
    }
}

Write-Log "Script started with parameters:"
Write-Log "  XML Path: $XmlFilePath"
Write-Log "  Log File: $LogFile"
Write-Log "  Data Key: $DataKey"
Write-Log "  Adobe Code: $AdobeCode"
Write-Log "  Elevated: $Elevated"
Write-Log "  PowerShell Version: $($PSVersionTable.PSVersion)"
Write-Log "  Process ID: $PID"

# Check if running as administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Log "This script requires administrator privileges."
    Write-Log "Attempting to restart with elevated permissions..."
    
    # Simple elevation using Start-Process
    $scriptPath = $MyInvocation.MyCommand.Path
    $arguments = "-ExecutionPolicy Bypass -File `"$scriptPath`" -Elevated -LogFile `"$LogFile`""
    
    if ($XmlFilePath) {
        $arguments += " -XmlFilePath `"$XmlFilePath`""
    }
    
    if ($DataKey) {
        $arguments += " -DataKey `"$DataKey`""
    }
    
    if ($AdobeCode) {
        $arguments += " -AdobeCode `"$AdobeCode`""
    }
    
    Write-Log "Launching with arguments: $arguments"
    try {
        # Wait a moment to ensure logs are written before elevation
        Start-Sleep -Seconds 1
        
        # Start elevated process and wait for it
        Write-Log "Starting elevated process..."
        Start-Process powershell -Verb RunAs -ArgumentList $arguments -Wait
        Write-Log "Elevated process completed."
    }
    catch {
        Write-Log "Failed to elevate: $_" -IsError
    }
    
    # Wait a moment to ensure any final logs are written
    Start-Sleep -Seconds 1
    exit
}

Write-Log "Running with administrator privileges"

# Check if the XML file exists
if (-not (Test-Path $XmlFilePath)) {
    Write-Log "Error: Could not find the XML file at: $XmlFilePath" -IsError
    exit 1
}

Write-Log "Found XML file at: $XmlFilePath"

# Test if we can read the file
try {
    $testContent = Get-Content -Path $XmlFilePath -TotalCount 5
    Write-Log "Successfully read the first few lines of the XML file."
    Write-Log "First line: $($testContent[0])"
} catch {
    Write-Log "Could not read the XML file: $_" -IsError
    exit 1
}

# Save the active network adapters to a file so we can restore them even if variables are lost
$adaptersInfoFile = "$env:TEMP\network_adapters_info_$PID.xml"

# Function to get active network adapters
function Get-ActiveNetworkAdapters {
    <#
    .SYNOPSIS
        Gets the list of currently active network adapters.
    
    .DESCRIPTION
        This function retrieves a list of active (status "Up") network adapters
        and saves their information to a file for later restoration.
    
    .OUTPUTS
        System.Array. An array of NetAdapter objects representing active network adapters.
    #>
    [CmdletBinding()]
    [OutputType([System.Array])]
    param()
    
    Write-Log "Getting active network adapters..."
    try {
        # Use Get-NetAdapter with more verbose output and force it to return an array
        $adapters = @(Get-NetAdapter -ErrorAction Stop | Where-Object { $_.Status -eq 'Up' })
        
        # Log detailed adapter info
        Write-Log "Found $($adapters.Count) active adapters"
        foreach ($adapter in $adapters) {
            Write-Log "  - Name: $($adapter.Name)"
            Write-Log "    InterfaceDescription: $($adapter.InterfaceDescription)"
            Write-Log "    Status: $($adapter.Status)"
            Write-Log "    MacAddress: $($adapter.MacAddress)"
            Write-Log "    InterfaceIndex: $($adapter.InterfaceIndex)"
        }
        
        # Save adapter info to file for later restoration
        if ($adapters.Count -gt 0) {
            try {
                $adapters | Export-Clixml -Path $adaptersInfoFile -Force
                Write-Log "Saved adapter information to: $adaptersInfoFile"
            } catch {
                Write-Log "Failed to save adapter information: $_" -IsError
            }
        }
        
        return $adapters
    }
    catch {
        Write-Log "Error getting network adapters: $_" -IsError
        return @()
    }
}

# Function to disable network adapters
function Disable-NetworkAdapters {
    <#
    .SYNOPSIS
        Disables a list of network adapters.
    
    .DESCRIPTION
        This function disables each of the provided network adapters.
        
    .PARAMETER Adapters
        An array of NetAdapter objects to disable.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Adapters
    )
    
    if ($Adapters.Count -eq 0) {
        Write-Log "No adapters to disable."
        return
    }
    
    Write-Log "Disabling network adapters..."
    foreach ($adapter in $Adapters) {
        try {
            Write-Log "Attempting to disable: $($adapter.Name) (Index: $($adapter.InterfaceIndex))"
            Disable-NetAdapter -Name $adapter.Name -Confirm:$false -ErrorAction Stop
            Start-Sleep -Seconds 1
            Write-Log "  - Successfully disabled: $($adapter.Name)"
        }
        catch {
            Write-Log "Failed to disable adapter $($adapter.Name): $_" -IsError
        }
    }
}

# Function to enable network adapters
function Enable-NetworkAdapters {
    <#
    .SYNOPSIS
        Enables a list of network adapters.
    
    .DESCRIPTION
        This function enables each of the provided network adapters and verifies
        that they are successfully brought online.
        
    .PARAMETER Adapters
        An array of NetAdapter objects to enable.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Adapters
    )
    
    if ($Adapters.Count -eq 0) {
        Write-Log "No adapters to enable. WILL NOT attempt to randomly enable adapters."
        return
    }
    
    Write-Log "Enabling network adapters..."
    foreach ($adapter in $Adapters) {
        try {
            Write-Log "Attempting to enable: $($adapter.Name) (Index: $($adapter.InterfaceIndex))"
            Enable-NetAdapter -Name $adapter.Name -Confirm:$false -ErrorAction Stop
            Start-Sleep -Seconds 1
            Write-Log "  - Successfully enabled: $($adapter.Name)"
            
            # Double check that it actually enabled
            $currentStatus = (Get-NetAdapter -Name $adapter.Name -ErrorAction Stop).Status
            Write-Log "  - Current status: $currentStatus"
            
            if ($currentStatus -ne "Up") {
                Write-Log "  - Adapter is not up, trying again..."
                Start-Sleep -Seconds 2
                Enable-NetAdapter -Name $adapter.Name -Confirm:$false -ErrorAction Stop
                $currentStatus = (Get-NetAdapter -Name $adapter.Name -ErrorAction Stop).Status
                Write-Log "  - Status after retry: $currentStatus"
            }
        }
        catch {
            Write-Log "Failed to enable adapter $($adapter.Name): $_" -IsError
            
            # Try enabling by InterfaceIndex as a fallback
            if ($adapter.InterfaceIndex) {
                try {
                    Write-Log "  - Trying to enable by InterfaceIndex: $($adapter.InterfaceIndex)"
                    Get-NetAdapter -InterfaceIndex $adapter.InterfaceIndex | Enable-NetAdapter -Confirm:$false
                    Write-Log "  - Enabled adapter by InterfaceIndex"
                } catch {
                    Write-Log "  - Failed to enable by InterfaceIndex: $_" -IsError
                }
            }
        }
    }
}

# Function to update the BridgeTalkVersion value in the XML file
function Update-BridgeTalkVersion {
    <#
    .SYNOPSIS
        Updates the specified version value in the XML file.
    
    .DESCRIPTION
        This function finds the specified Data node in the XML file,
        increments its value, and saves the file.
    
    .PARAMETER XmlDocument
        The XML document object to update.
    
    .PARAMETER XmlFilePath
        The path to the XML file to save changes to.
        
    .PARAMETER DataKey
        The key attribute value to search for in Data nodes.
        
    .PARAMETER AdobeCode
        The adobeCode attribute to search for in Other nodes.
    
    .OUTPUTS
        System.Boolean. Returns $true if successful, $false otherwise.
    #>
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param(
        [Parameter(Mandatory = $true)]
        [xml]$XmlDocument,
        
        [Parameter(Mandatory = $true)]
        [string]$XmlFilePath,
        
        [Parameter(Mandatory = $false)]
        [string]$DataKey = "BridgeTalkVersion",
        
        [Parameter(Mandatory = $false)]
        [string]$AdobeCode = "/adobe/bridgetalk/photoshop-60.064"
    )
    
    # Find the node with the specific adobeCode
    Write-Log "Searching for target node in XML..."
    # Store the original XPath for reference
    $xpath = "//Configuration/Other[@adobeCode='$AdobeCode']/Data[@key='$DataKey']"
    $node = $XmlDocument.SelectSingleNode($xpath)

    if ($node) {
        # Convert current version to integer, increment, and convert back to string
        try {
            $currentVersion = [System.Numerics.BigInteger]::Parse($node.InnerText)
            $newVersion = $currentVersion + 1
            $node.InnerText = $newVersion.ToString()
            
            Write-Log "Found target node with current version: $currentVersion"
            Write-Log "Updating to new version: $newVersion"
            
            # Save the changes
            try {
                $XmlDocument.Save($XmlFilePath)
                Write-Log "Successfully saved updated XML file"
                Write-Log "XML successfully updated from version $currentVersion to $newVersion"
                return $true
            }
            catch {
                Write-Log "Error saving XML file: $_" -IsError
                return $false
            }
        }
        catch {
            Write-Log "Error processing version number: $_" -IsError
            return $false
        }
    } else {
        Write-Log "Could not find the specified XML node in the file" -IsError
        Write-Log "XML file structure may not be as expected" -IsError
        try {
            Write-Log "First 50 characters of XML content: $($XmlDocument.OuterXml.Substring(0, [Math]::Min(50, $XmlDocument.OuterXml.Length)))"
            
            # Check if we can find any nodes that might be similar
            $otherDataNodes = $XmlDocument.SelectNodes("//Data[@key]")
            if ($otherDataNodes -and $otherDataNodes.Count -gt 0) {
                Write-Log "Found $($otherDataNodes.Count) Data nodes with keys"
                foreach ($dataNode in ($otherDataNodes | Select-Object -First 3)) {
                    Write-Log "  - Data node with key='$($dataNode.key)', value='$($dataNode.InnerText)'"
                }
            }
            
            $otherNodes = $XmlDocument.SelectNodes("//Other[@adobeCode]")
            if ($otherNodes -and $otherNodes.Count -gt 0) {
                Write-Log "Found $($otherNodes.Count) Other nodes with adobeCode"
                foreach ($otherNode in ($otherNodes | Select-Object -First 3)) {
                    Write-Log "  - Other node with adobeCode='$($otherNode.adobeCode)'"
                }
            }
        } catch {
            Write-Log "Could not extract XML content: $_" -IsError
        }
        return $false
    }
}

try {
    # Get list of active adapters before disabling
    Write-Log "Attempting to get network adapters..."
    $activeAdapters = Get-ActiveNetworkAdapters
    
    if ($activeAdapters.Count -eq 0) {
        Write-Log "WARNING: No active network adapters found!"
    } else {
        # Ensure we have the adapter objects in memory
        Write-Log "Storing $($activeAdapters.Count) network adapters for later restoration"
        $global:adaptersToRestore = $activeAdapters
    }

    # Disable all active adapters if we found any
    Write-Log "About to disable network adapters..."
    Disable-NetworkAdapters -Adapters $activeAdapters

    # Load the XML file
    Write-Log "Loading XML file..."
    try {
        $xml = [xml](Get-Content $XmlFilePath)
        Write-Log "XML file loaded successfully"
    }
    catch {
        Write-Log "Error loading XML file: $_" -IsError
        exit 1
    }

    # Update the BridgeTalkVersion value in the XML file
    if (Update-BridgeTalkVersion -XmlDocument $xml -XmlFilePath $XmlFilePath -DataKey $DataKey -AdobeCode $AdobeCode) {
        # ===========================================================================
        # IMPORTANT: This pause is required before re-enabling network adapters!
        # Do not remove this pause - the user must be allowed to verify the changes
        # before network connectivity is restored.
        # ===========================================================================
        Write-Log " "
        Write-Log "===================================================="
        Write-Log "!!! XML modification complete !!!"
        Write-Log "Press Enter in this window to restore network connectivity..."
        Write-Log "===================================================="
        Write-Log " "
        $null = Read-Host
        Write-Log "User confirmed. Proceeding to restore network connectivity."
    } else {
        exit 1
    }
}
catch {
    # Only write the error if there actually is one
    if ($_) {
        Write-Log "An unexpected error occurred: $_" -IsError
    }
    exit 1
}
finally {
    # Always try to re-enable the adapters, even if the XML operation failed
    Write-Log "Re-enabling network adapters..."
    
    # First try to load from saved file (most reliable in case variables were lost)
    if (Test-Path $adaptersInfoFile) {
        try {
            Write-Log "Loading saved adapter information from file..."
            $restoredAdapters = Import-Clixml -Path $adaptersInfoFile
            Write-Log "Loaded $($restoredAdapters.Count) adapters from saved file"
            Enable-NetworkAdapters -Adapters $restoredAdapters
        } catch {
            Write-Log "Failed to load adapter information from file: $_" -IsError
            # Fall through to next option
        }
    } 
    # Then try using the global variable
    elseif ($global:adaptersToRestore -and $global:adaptersToRestore.Count -gt 0) {
        Write-Log "Using stored adapter objects from memory"
        Enable-NetworkAdapters -Adapters $global:adaptersToRestore
    }
    # Then try the local variable
    elseif ($activeAdapters -and $activeAdapters.Count -gt 0) {
        Write-Log "Using local adapter objects"
        Enable-NetworkAdapters -Adapters $activeAdapters
    }
    # If all else fails, we might have lost track of the adapters
    else {
        Write-Log "WARNING: Could not find any record of disabled adapters!" -IsError
        Write-Log "This script will NOT attempt to enable random adapters for safety reasons." -IsError
        Write-Log "You may need to manually re-enable your network adapters." -IsError
    }
    
    Write-Log "Script completed at $(Get-Date)"
    
    # Clean up
    if (Test-Path $adaptersInfoFile) {
        try {
            Remove-Item -Path $adaptersInfoFile -Force
            Write-Log "Removed temporary adapter info file"
        } catch {
            Write-Log "Failed to remove temporary adapter info file: $_" -IsError
        }
    }
}