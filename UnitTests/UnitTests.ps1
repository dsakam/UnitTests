######## LICENSE ####################################################################################################################################
<#
 # Copyright (c) 2013-2014, Daiki Sakamoto
 # All rights reserved.
 #
 # Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
 #   - Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
 #   - Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer
 #     in the documentation and/or other materials provided with the distribution.
 #
 # THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
 # THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
 # HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
 # LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
 # ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE
 # USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 #
 #>
 # http://opensource.org/licenses/BSD-2-Clause
#####################################################################################################################################################

######## HISTORY ####################################################################################################################################
<#
 # Unit Tests for PowerShell Modules
 #
 # 2013/09/02  Version 0.0.0.1
 # 2013/09/04  Update
 # 2013/09/09  Update
 # 2013/09/18  Update
 # 2013/09/23  Update
 # 2013/10/24  Update
 #     :
 #     :
 # 2014/05/06  Version 1.0.0.0
 # 2014/05/09  Version 1.0.1.0
 # 2014/05/22  Version 1.1.0.0    Always import all modules (2014/05/24)
 #                                Add test cases for 'New-ZipFile' cmdlet. (2014/05/24)
 #                                Change process of exception. (2014/05/25)
 #
 #>
#####################################################################################################################################################
'[Unit Tests for PowerShell Moduels] Script Version ' + ($version = '1.1.0.0')

#####################################################################################################################################################
# PARAMETERS

# Update (Only Importing Modules)
[bool]$Update = $false

# Clean (Cleaning Only)
[bool]$Clean = $false

# Remote (Remote Computer Test Only)
[bool]$Remote = $false


# Verbose
[bool]$Verbose = $true

# Skip GUI Test option
[bool]$Skip_GUI = $false


# Additional Tests
[bool]$test_of_TestCommand = $true
[bool]$test_of_HelpContent = $false

#####################################################################################################################################################
# Target Module
$TargetModule = @(
    @{
        Path = "PackageBuilder\PackageBuilder.psd1";
        Target = $true;
        Commands = @(

            # PackageBuilder.Utilities
            @{ Name = "New-GUID"; Target = $true },
            @{ Name = "New-HR"; Target = $true; },
            @{ Name = "Write-Title"; Target = $true },
            @{ Name = "Write-Boolean"; Target = $true },
            @{ Name = "Show-Message"; Target = $true; "GUI" = $true },
            @{ Name = "Get-DateString"; Target = $true },
            @{ Name = "Get-FileVersionInfo"; Target = $true },
            @{ Name = "Get-ProductName"; Target = $true },
            @{ Name = "Get-FileDescription"; Target = $true },
            @{ Name = "Get-FileVersion"; Target = $true },
            @{ Name = "Get-ProductVersion"; Target = $true },
            @{ Name = "Get-HTMLString"; Target = $true },
            @{ Name = "Get-PrivateProfileString"; Target = $true },
            @{ Name = "Update-Content"; Target = $true },
            @{ Name = "Get-WindowHandler"; Target = $true },
            @{ Name = "New-StructArray"; Target = $true },
            @{ Name = "Get-ByteArray"; Target = $false },
            @{ Name = "ConvertFrom-ByteArray"; Target = $true },
            @{ Name = "ConvertTo-ByteArray"; Target = $true },

            # PackageBuilder.Win32
            @{ Name = "Invoke-LoadLibraryEx"; Target = $true },
            @{ Name = "Invoke-FreeLibrary"; Target = $false },
            @{ Name = "Invoke-LoadString"; Target = $false },
            @{ Name = "Get-ResourceString"; Target = $true },
            @{ Name = "Invoke-HtmlHelp"; Target = $true },

            # PackageBuilder.Core
            @{ Name = "Get-MD5"; Target = $true },
            @{ Name = "Start-Command"; Target = $true; "GUI" = $true },
            @{ Name = "Test-SameFile"; Target = $true },
            @{ Name = "New-ISOImageFile"; Target = $true },

            # PackageBuilder.Remote
            @{ Name = "Stop-Host"; Target = $true; "Remote" = $true },
            @{ Name = "Restart-Host"; Target = $true; "Remote" = $true },
            @{ Name = "Start-Computer"; Target = $true; "Remote" = $true },

            # PackageBuilder.Legacy
            @{ Name = "Invoke-LoadLibrary"; Target = $true },
            @{ Name = "Get-CheckSum"; Target = $true },
            @{ Name = "Send-Mail"; Target = $false }
        );
    },
    @{
        Path = "ZipFile\ZipFile.psd1";
        Target = $true;
        Commands = @(

            # ZipFile
            @{ Name = "Expand-ZipFile"; Target = $true },
            @{ Name = "New-ZipFile"; Target = $true }
        );
    }
)

#####################################################################################################################################################
# Path

# Current Directory (this Test Script in this Fodler)
$CurrentDirectory = $PSScriptRoot | Convert-Path
$current_FolderPath = $CurrentDirectory

# Root Directory
$RootDirectory = $CurrentDirectory | Split-Path -Parent | Convert-Path
$root_FolderPath = $RootDirectory

# Test Data Folder
$TestDataFolder = $CurrentDirectory | Join-Path -ChildPath 'TestData' | Convert-Path
$testdata_FolderPath = $TestDataFolder

#####################################################################################################################################################
# Private Data

$test_Hostname   = Get-Content -Path ($RootDirectory | Join-Path -ChildPath '..\TestData\Private\HostName')
$test_MacAddress = Get-Content -Path ($RootDirectory | Join-Path -ChildPath '..\TestData\Private\MacAddress')
$test_UserName   = Get-Content -Path ($RootDirectory | Join-Path -ChildPath '..\TestData\Private\UserName')
$test_Password   = Get-Content -Path ($RootDirectory | Join-Path -ChildPath '..\TestData\Private\Password')

#####################################################################################################################################################
# Functions

Function Test-Command {

    [CmdletBinding ()]
    Param (
        [Parameter (Mandatory=$true, Position=0, ParameterSetName="with_Message")][string]$Message,

        [Parameter (Mandatory=$true, Position=1, ParameterSetName="with_Message", ValueFromPipeline=$true)]
        [Parameter (Mandatory=$true, Position=0, ParameterSetName="only_Code", ValueFromPipeline=$true)]
        [scriptblock]$TestCode
    )

    Process
    {
        # Print Message (or Test Code)
        if ($Message) { Write-Host ($Message + "... ") -NoNewline }
        else { Write-Host ("{" + $TestCode.ToString() + "} ") -NoNewline }


        # Execute Test Code
        $result = $false
        if (-not [bool]::TryParse((& $TestCode), [ref]$result)) { Write-Host (" [-]") }
        else {
            Write-Host (" [") -NoNewline
            PASS_FAIL $result
            Write-Host "]"
        }
    }
}

Function Test-Module {

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true, Position=0)][object[]]$Modules,
        [Parameter(Mandatory=$true, Position=1)][string]$Command,
        [Parameter(Mandatory=$true, Position=2)][scriptblock]$TestCode
    )

    Process
    {
        $Modules | % {
            if ($_.Target)
            {
                if ($_.Commands | ? { $_.Target -and ($_.Name -eq $Command) } )
                {
                    # Title
                    Write-Host (LINE)
                    Write-Host ("Unit Tests of '" + $Command + "' Command")
                    Write-Host (PRINT START)

                    # Do Test(s) [*]V1.1.0.0 (2014/05/22)
                    [void](& $TestCode)
                }
            }
        }
    }
}

Function Exit-Script {

    [CmdletBinding ()]
    Param (
        [Parameter (Mandatory=$false)][object]$e
    )

    Process
    {
        if ($Script:failed)
        {
            # Failure (Fail)
            Write-Host ("`n" + (New-HR -Char "#")) -ForegroundColor Red
            Write-Host "Unit Tests for Package Builder Toolkit is completed, but including some failures..." -ForegroundColor Red
            Write-Host "Please Check!" -ForegroundColor Red
            Write-Host
            throw $e
        }
        else
        {
            # Success (Pass)
            Write-Host ("`n" + (New-HR -Char "#")) -ForegroundColor Green
            Write-Host "Unit Tests for Package Builder Toolkit is completed successfully!" -ForegroundColor Green
            exit 0
        }
    }
}

#####################################################################################################################################################
# Exception
[bool]$Script:failed = $false

trap
{
    $Script:failed = $true
    Exit-Script $_
}

#####################################################################################################################################################
# Verbose
$default_VerbosePreference = $VerbosePreference

if ($Verbose -eq $true) { $VerbosePreference = "Continue" }
else { $VerbosePreference = "SilentlyContinue" }

#####################################################################################################################################################
# Update Modules
$TargetModule | % {

    # Always import all modules / [*]V1.1.0.0 (2014/05/24)
    if ($true)
    {
        # Validate Module Path
        if (-not (Test-Path -Path ($p = $root_FolderPath | Join-Path -ChildPath $_.'Path'))) { throw New-Object System.IO.FileNotFoundException }

        # Remove Module
        if (Get-Module -Name ($m = (Get-Item -Path $p).BaseName))
        {
            Write-Host ("Removing Module '" + $m + "'...")
            if ($VerbosePreference -eq "SilentlyContinue") { Remove-Module -Name $m }
            else { Remove-Module -Name $m -Verbose } # ("Continue", "Stop" or "Inquire")
        }

        # Import Module
        Write-Host ("Importing Module '" + $m + "'...")
        if ($VerbosePreference -eq "SilentlyContinue") { Import-Module -Name $p }
        else { Import-Module -Name $p -Verbose } # ("Continue", "Stop" or "Inquire")
    }
}

# Update (Only Importing Modules)
if ($Update) { Exit-Script }

#####################################################################################################################################################
# Clean (Cleaning Only)
if ($Clean)
{
    # .\TestData\ISOImageFile\*.iso
    if (($p = $testdata_FolderPath | Join-Path -ChildPath "ISOImageFile\*") | Test-Path -Include "*.iso")
    {
        Remove-Item -Path $p -Force -Include "*.iso"
        Write-Warning ( "'$p' were cleaned up.")
    }

    # .\TestData\ZipFile\*
    if (($p = $testdata_FolderPath | Join-Path -ChildPath "ZipFile\*") | Test-Path -Exclude "Archive", "*.*")
    {
        $p | Get-ChildItem -Recurse -Exclude "Archive", "*.*" | Remove-Item -Force -Recurse
        Write-Warning ( "'$p' were cleaned up.")
    }

    Exit-Script
}

#####################################################################################################################################################
# Print Message
MAIN_TITLE "Unit Tests for BUILDLet PowerShell Commands"

# Print Root Path
Write-Host
Write-Host "Root Directory Path is..."
Write-Host ("`t" + $root_FolderPath)

# Print Current Path (this Test Script)
Write-Host
Write-Host "Test Script Directory Path is..."
Write-Host ("`t" + $current_FolderPath)

# Print Test Data Path
Write-Host
Write-Host "Test Data Directory Path is..."
Write-Host ("`t" + $testdata_FolderPath)

# Print Target Module(s)
Write-Host
Write-Host "Target Module(s) are..."
$TargetModule | % {
    if ($_.Target) 
    {
        Write-Host ("`t" + ($root_FolderPath | Join-Path -ChildPath $_.Path))
    }
}

# Print Test Targets
Write-Host
Write-Host "Target Command(s) are..."
$TargetModule | % {
    if ($_.Target)
    {
        $mod = ((Get-Item -Path ($root_FolderPath | Join-Path -ChildPath $_.Path)).BaseName)
        Write-Host ("`n`t'" + $mod + "' Module:")

        # Check Target Command
        $_.Commands | % {
            Write-Host ("`t`t" + $_.Name + " [") -NoNewline
            if ((Get-Module -Name $mod).ExportedCommands.Keys -contains $_.Name)
            {
                YES_NO $_.Target
            }
            else { Write-Host $_.Target -NoNewline }
            Write-Host ("]")
        }


        $commands = $_.Commands

        # Double Check
        (Get-Module -Name $mod).ExportedCommands.Keys | % {
            $found = $false
            $cmd = $_
            $commands | % {
                if ($_.Name -eq $cmd) { $found = $true }
            }
            if (-not $found) { Write-Warning ("`t" + "Test of '" + $cmd + "' command is not including test cases.") }
        }
    }
}

# Print Skip GUI Test option
Write-Host
Write-Host "Skip GUI Test option is... " -NoNewline
TRUE_FALSE $Skip_GUI
Write-Host
if ($Skip_GUI)
{
    $TargetModule | % {
        $_.Commands | % {
            if ($_.GUI)
            {
                $_.Target = $false
                Write-Host ("`t" + "Test of '" + $_.Name + "' Command is to be skipped...")
            }
        }
    }
}

# Print target of Remote Computer Test
Write-Host
Write-Host "Remote Computer Test option is... " -NoNewline
TRUE_FALSE $Remote
Write-Host
$TargetModule | % {
    $_.Commands | % {
        if ($Remote)
        {
            if ($_.Remote)
            {
                Write-Host ("`t" + "Test of '" + $_.Name + "' Command is to be [") -NoNewline
                Write-Boolean -TestObject $_.Target -Green Done -Red Skipped
                Write-Host "]"
            }
            else
            {
                $_.Target = $false
                Write-Host ("`t" + "Test of '" + $_.Name + "' Command is to be skipped...")
            }
        }
        else
        {
            if ($_.Remote)
            {
                $_.Target = $false
                Write-Host ("`t" + "Test of '" + $_.Name + "' Command is to be skipped...")
            }
        }
    }
}

# Print Verbose option
Write-Host
Write-Host "'Verbose' option is... " -NoNewline
TRUE_FALSE $Verbose
Write-Host

# Start Message
Write-Host
Write-Host (New-HR -Char "#")
Write-Host "Starting Unit Tests for BUILDLet PowerShell Commands..!"

#####################################################################################################################################################
# Test-Command
if ($test_of_TestCommand)
{
    Write-Host (LINE)
    Write-Host ("Unit Tests of 'Test-Command' Command")
    Write-Host (PRINT START)
    $i = 0
    
    Test-Command -Message Test-Command -TestCode {
        Write-Host This test result should be `" -NoNewline
        Write-Host Fail -ForegroundColor Red -NoNewline
        Write-Host `". -NoNewline
        return $false
    }

    Test-Command Test-Command {
        Write-Host This test result should be `" -NoNewline
        Write-Host Pass -ForegroundColor Green -NoNewline
        Write-Host `". -NoNewline
        return $true
    }

    Test-Command -TestCode { return $false }
    Test-Command { return $true }

    Test-Command (MESSAGE Test-Command, $i) { Write-Host This test result should be `"-`". -NoNewline }

    Test-Command (MESSAGE Test-Command, (++$i), (++$i), (++$i)) {
        Write-Host "This is test of combination of 'PRINT' function + 'MESSAGE' function." -NoNewline
    }
}


#####################################################################################################################################################
# Unit Tests for PackageBuilder Module
#####################################################################################################################################################

#####################################################################################################################################################
# New-GUID
Test-Module $TargetModule New-GUID {

    Test-Command New-GUID { Write-Host (New-GUID) -NoNewline }
}

#####################################################################################################################################################
# New-HR
Test-Module $TargetModule New-HR {
    $i = 0

    Test-Command (MESSAGE New-HR, (++$i)) { Write-Host ("`n" + (New-HR)) -ForegroundColor Yellow }
    Test-Command (MESSAGE New-HR, (++$i)) { Write-Host ("`n" + (New-HR -Char "#")) -ForegroundColor Yellow }
    Test-Command (MESSAGE New-HR, (++$i)) { Write-Host ("`n" + (New-HR -Char "*" -Length 100)) -ForegroundColor Yellow }

    Test-Command (MESSAGE New-HR, (++$i)) {
        return ((New-HR -Char "#" -Length 50) -eq "##################################################")
    }
}

#####################################################################################################################################################
# Write-Title
Test-Module $TargetModule Write-Title {
    $i = 0

    Test-Command (MESSAGE Write-Title, (++$i)) { Write-Title -Text "This is Test Case $i of 'Write-Title' Command" -Color Yellow }
    Test-Command (MESSAGE Write-Title, (++$i)) {
        Write-Title -Text "This is Test Case $i of 'Write-Title' Command" `
            -Char "*" -Width 100 -Padding 1 -ColumnWidth 1 -MinWidth 10 -MaxWidth 255 -Color Yellow
    }

    Test-Command (MESSAGE Write-Title, (++$i)) {
        Write-Title -Text @("Write-Title", "", "This is No. $i of Test Case of", "'Write-Title' Command.") `
            -Char "+" -Width 42 -Padding 3 -ColumnWidth 1 -MinWidth 10 -MaxWidth 255 -Color Yellow
    }

    Test-Command (MESSAGE Write-Title, (++$i)) { Write-Title -Text "" -Color Yellow }

    Write-Host
    Test-Command (MESSAGE Write-Title, (++$i)) { Write-Title -Text "" -Width 6 -Color Yellow }
    Test-Command (MESSAGE Write-Title, (++$i)) { Write-Title -Text "" -Width 6 -ColumnWidth 1 -Color Yellow }

    Write-Host
    Test-Command (MESSAGE Write-Title, (++$i)) { Write-Title -Text "hogehoge" -Width 12 -Color Yellow }
    Test-Command (MESSAGE Write-Title, (++$i)) { Write-Title -Text "hogeho" -Width 12 -Color Yellow }
    Test-Command (MESSAGE Write-Title, (++$i)) { Write-Title -Text "hogeh" -Width 12 -Color Yellow }
    Test-Command (MESSAGE Write-Title, (++$i)) { Write-Title -Text "hoge" -Width 12 -Color Yellow }

    Write-Host
    Test-Command (MESSAGE Write-Title, (++$i)) { Write-Title -Text "hoge" -Width 13 -Color Yellow }
    Test-Command (MESSAGE Write-Title, (++$i)) { Write-Title -Text "hoge" -Width 14 -Color Yellow }
    Test-Command (MESSAGE Write-Title, (++$i)) { Write-Title -Text "hoge" -Width 15 -Color Yellow }

    Write-Host
    Test-Command (MESSAGE Write-Title, (++$i)) { Write-Title -Text "hoge" -Width 7 -Color Yellow }
    Test-Command (MESSAGE Write-Title, (++$i)) { Write-Title -Text "hog" -Width 7 -Color Yellow }
    Test-Command (MESSAGE Write-Title, (++$i)) { Write-Title -Text "ho" -Width 7 -Color Yellow }
    Test-Command (MESSAGE Write-Title, (++$i)) { Write-Title -Text "h" -Width 7 -Color Yellow }
}

#####################################################################################################################################################
# Write-TestResult
Test-Module $TargetModule Write-Boolean {
    $i = 0

    Test-Command { Write-Boolean -TestObject $false }
    Test-Command { Write-Boolean -TestObject $true }

    Test-Command (MESSAGE Write-Boolean, (++$i)) { Write-Boolean -TestObject $true -Green "This is Green" -Red "This is Red" }
    Test-Command (MESSAGE Write-Boolean, (++$i)) { Write-Boolean -TestObject $false -Green "This is Green" -Red "This is Red" }
}

#####################################################################################################################################################
# Show-Message
Test-Module $TargetModule Show-Message {
    $i = 0

    Test-Command (MESSAGE Show-Message, (++$i)) {
        Write-Host "ダイアログの OK ボタンをクリックしてください。" -ForegroundColor Magenta -NoNewline
        Show-Message -Text "メッセージ" -Caption "タイトル"
    }

    Test-Command (MESSAGE Show-Message, (++$i)) {
        Write-Host "ダイアログの OK ボタンをクリックしてください。" -ForegroundColor Magenta -NoNewline
        Show-Message "Package Builder Toolkit for PowerShell"
    }
}

#####################################################################################################################################################
# Get-DateString
Test-Module $TargetModule Get-DateString {
    $i = 0

    Write-Host
    Test-Command { Write-Host (Get-DateString) -ForegroundColor Yellow -NoNewline }
    Test-Command { Write-Host (Get-DateString -Date 9月4日 -LCID en-US) -ForegroundColor Yellow -NoNewline}
    Test-Command { Write-Host (Get-DateString -Date 2013.9.4 -Format yyyy/MM/dd) -ForegroundColor Yellow -NoNewline}

    Write-Host
    Test-Command { Write-Host (Get-DateString -LCID zh-CN -Format D) -ForegroundColor Yellow -NoNewline}
    Test-Command { Write-Host (Get-DateString -LCID en-US -Format D) -ForegroundColor Yellow -NoNewline}
    Test-Command { Write-Host (Get-DateString -LCID fr -Format D) -ForegroundColor Yellow -NoNewline}
    Test-Command { Write-Host (Get-DateString -LCID it -Format D) -ForegroundColor Yellow -NoNewline}
    Test-Command { Write-Host (Get-DateString -LCID de -Format D) -ForegroundColor Yellow -NoNewline}
    Test-Command { Write-Host (Get-DateString -LCID es -Format D) -ForegroundColor Yellow -NoNewline}
    Test-Command { Write-Host (Get-DateString -LCID ja) -ForegroundColor Yellow -NoNewline}
    Test-Command { Write-Host (Get-DateString -LCID pl) -ForegroundColor Yellow -NoNewline}
    Test-Command { Write-Host (Get-DateString -LCID Ru) -ForegroundColor Yellow -NoNewline}
    Test-Command { Write-Host (Get-DateString -LCID Tr) -ForegroundColor Yellow -NoNewline}
    Test-Command { Write-Host (Get-DateString -LCID Ar) -ForegroundColor Yellow -NoNewline}

    Write-Host
    Test-Command { Write-Host (Get-DateString -LCID Ar-eg) -ForegroundColor Yellow -NoNewline}
    Test-Command { Write-Host (Get-DateString -LCID Ar-sa) -ForegroundColor Yellow -NoNewline}

    Write-Host
    Test-Command (MESSAGE Get-DateString, (++$i)) { (Get-DateString -Date 2013/9/29) -eq "2013年9月29日" }
    Test-Command (MESSAGE Get-DateString, (++$i)) { (Get-DateString -Date 2013/9/29 -Format yyyy/MM/dd) -eq "2013/09/29" }
    Test-Command (MESSAGE Get-DateString, (++$i)) { (Get-DateString -Date 2013/9/29 -LCID en-US) -eq "Sunday, September 29, 2013" }
}

#####################################################################################################################################################
# Get-FileVersionInfo
Test-Module $TargetModule Get-FileVersionInfo {
    $i = 0

    # Print Test Target File Path
    Write-Host
    Write-Host '$filepath =' ($filepath = ($testdata_FolderPath | Join-Path -ChildPath "win32rsc.dll"))
    Write-Host '$filename =' ($filename = Split-Path -Path $filepath -Leaf)

    # Printout content of test file.
    Test-Command (MESSAGE Get-FileVersionInfo, (++$i), $filename) {
        Write-Host
        Write-Verbose (VERBOSE_LINE)
        Write-Verbose (Get-FileVersionInfo -Path $filepath)
        Write-Verbose (VERBOSE_LINE)
    }


    # ProductName / FileDescription
    Test-Command { Write-Host (Get-FileVersionInfo -Path $filepath -ProductName) -ForegroundColor Yellow -NoNewline }
    Test-Command { Write-Host (Get-FileVersionInfo -Path $filepath -FileDescription) -ForegroundColor Yellow -NoNewline }

    Test-Command (MESSAGE Get-FileVersionInfo, (++$i), $filename) { (Get-FileVersionInfo -Path $filepath -ProductName) -eq "This is ProductName" }
    Test-Command (MESSAGE Get-FileVersionInfo, (++$i), $filename) { (Get-FileVersionInfo -Path $filepath -FileDescription) -eq "This is FileDescription" }


    # FileVersion
    Test-Command { Write-Host (Get-FileVersionInfo -Path $filepath -FileVersion) -ForegroundColor Yellow -NoNewline }
    Test-Command { Write-Host (Get-FileVersionInfo -Path $filepath -FileVersion -Major) -ForegroundColor Yellow -NoNewline }
    Test-Command { Write-Host (Get-FileVersionInfo -Path $filepath -FileVersion -Minor) -ForegroundColor Yellow -NoNewline }
    Test-Command { Write-Host (Get-FileVersionInfo -Path $filepath -FileVersion -Build) -ForegroundColor Yellow -NoNewline }
    Test-Command { Write-Host (Get-FileVersionInfo -Path $filepath -FileVersion -Private) -ForegroundColor Yellow -NoNewline }
    Test-Command { Write-Host (Get-FileVersionInfo -Path $filepath -FileVersion -Composite) -ForegroundColor Yellow -NoNewline }

    Test-Command (MESSAGE Get-FileVersionInfo, (++$i), $filename) { (Get-FileVersionInfo -Path $filepath -FileVersion) -eq "1.2.3.4" }
    Test-Command (MESSAGE Get-FileVersionInfo, (++$i), $filename) { (Get-FileVersionInfo -Path $filepath -FileVersion -Major) -eq 1 }
    Test-Command (MESSAGE Get-FileVersionInfo, (++$i), $filename) { (Get-FileVersionInfo -Path $filepath -FileVersion -Minor) -eq 2 }
    Test-Command (MESSAGE Get-FileVersionInfo, (++$i), $filename) { (Get-FileVersionInfo -Path $filepath -FileVersion -Build) -eq 3 }
    Test-Command (MESSAGE Get-FileVersionInfo, (++$i), $filename) { (Get-FileVersionInfo -Path $filepath -FileVersion -Private) -eq 4 }
    Test-Command (MESSAGE Get-FileVersionInfo, (++$i), $filename) { (Get-FileVersionInfo -Path $filepath -FileVersion -Composite) -eq "1.2.3.4" }


    # ProductVersion
    Test-Command { Write-Host (Get-FileVersionInfo -Path $filepath -ProductVersion) -ForegroundColor Yellow -NoNewline }
    Test-Command { Write-Host (Get-FileVersionInfo -Path $filepath -ProductVersion -Major) -ForegroundColor Yellow -NoNewline }
    Test-Command { Write-Host (Get-FileVersionInfo -Path $filepath -ProductVersion -Minor) -ForegroundColor Yellow -NoNewline }
    Test-Command { Write-Host (Get-FileVersionInfo -Path $filepath -ProductVersion -Build) -ForegroundColor Yellow -NoNewline }
    Test-Command { Write-Host (Get-FileVersionInfo -Path $filepath -ProductVersion -Private) -ForegroundColor Yellow -NoNewline }
    Test-Command { Write-Host (Get-FileVersionInfo -Path $filepath -ProductVersion -Composite) -ForegroundColor Yellow -NoNewline }

    Test-Command (MESSAGE Get-FileVersionInfo, (++$i), $filename) { (Get-FileVersionInfo -Path $filepath -ProductVersion) -eq "10.11.101.1001" }
    Test-Command (MESSAGE Get-FileVersionInfo, (++$i), $filename) { (Get-FileVersionInfo -Path $filepath -ProductVersion -Major) -eq 10 }
    Test-Command (MESSAGE Get-FileVersionInfo, (++$i), $filename) { (Get-FileVersionInfo -Path $filepath -ProductVersion -Minor) -eq 11 }
    Test-Command (MESSAGE Get-FileVersionInfo, (++$i), $filename) { (Get-FileVersionInfo -Path $filepath -ProductVersion -Build) -eq 101 }
    Test-Command (MESSAGE Get-FileVersionInfo, (++$i), $filename) { (Get-FileVersionInfo -Path $filepath -ProductVersion -Private) -eq 1001 }
    Test-Command (MESSAGE Get-FileVersionInfo, (++$i), $filename) { (Get-FileVersionInfo -Path $filepath -ProductVersion -Composite) -eq "10.11.101.1001" }


    # Wrapper Commands
    Test-Command Get-ProductName { (Get-ProductName -Path $filepath) -eq "This is ProductName" }
    Test-Command Get-ProductVersion { (Get-ProductVersion -Path $filepath) -eq "10.11.101.1001" }
    Test-Command Get-FileDescription { (Get-FileDescription -Path $filepath) -eq "This is FileDescription" }
    Test-Command Get-FileVersion { (Get-FileVersion -Path $filepath) -eq "1.2.3.4" }


    # Invalid Parameter - File Not Found
    Write-Host
    Write-Host '$filepath =' ($filepath = ($testdata_FolderPath | Join-Path -ChildPath "dummy.dll"))
    Write-Host '$filename =' ($filename = Split-Path -Path $filepath -Leaf)

    Test-Command (MESSAGE Get-FileVersionInfo, (++$i), $filename) {
        try
        {
            Get-FileVersionInfo -Path $filepath
            return $false
        }
        catch
        {
            Write-Warning $_
            if (([System.Management.Automation.ErrorRecord]$_).Exception.InnerException.InnerException -is [System.IO.FileNotFoundException]) { return $true }
            else { return $false }
        }
    }

    # Invalid Parameter - Target is not File (is Folder).
    Write-Host
    Write-Host '$filepath =' ($filepath = $testdata_FolderPath) '(Directory)'

    Test-Command (MESSAGE Get-FileVersionInfo, (++$i), (Split-Path -Path $filepath -Leaf)) {
        try
        {
            Get-FileVersionInfo -Path $filepath
            return $false
        }
        catch
        {
            Write-Warning $_
            if (([System.Management.Automation.ErrorRecord]$_).Exception.InnerException.InnerException -is [System.IO.FileNotFoundException]) { return $true }
            else { return $false }
        }
    }
}

#####################################################################################################################################################
# Get-HTMLString
#
# "example.html" is copied from the following URL.
# The global structure of an HTML document <http://www.w3.org/TR/REC-html40/struct/global.html>
#
Test-Module $TargetModule Get-HTMLString {
    $i = 0

    # File 1
    Write-Host
    Write-Host '$filepath =' ($filepath = ($testdata_FolderPath | Join-Path -ChildPath "example.html"))
    Write-Host '$filename =' ($filename = Split-Path -Path $filepath -Leaf)
    Write-Verbose (VERBOSE_LINE)
    Write-Verbose (Get-Content -Path $filepath -Raw)
    Write-Verbose (VERBOSE_LINE)

    Test-Command { Write-Host (Get-HTMLString -Path $filepath -Tag "title") -ForegroundColor Yellow -NoNewline }
    Test-Command { Write-Host (Get-HTMLString -Path $filepath -Tag "p") -ForegroundColor Yellow -NoNewline }

    Test-Command (MESSAGE Get-HTMLString, (++$i), $filename) { (Get-HTMLString -Path $filepath -Tag "title") -eq "My first HTML document" }
    Test-Command (MESSAGE Get-HTMLString, (++$i), $filename) { (Get-HTMLString -Path $filepath -Tag "p") -eq "Hello world!" }


    # File 2
    Write-Host
    Write-Host '$filepath =' ($filepath = ($testdata_FolderPath | Join-Path -ChildPath "TheProject.html"))
    Write-Host '$filename =' ($filename = Split-Path -Path $filepath -Leaf)
    Write-Verbose (VERBOSE_LINE)
    Write-Verbose (Get-Content -Path $filepath -Raw)
    Write-Verbose (VERBOSE_LINE)

    Test-Command { Write-Host (Get-HTMLString -Path $filepath -Tag "h1") -ForegroundColor Yellow -NoNewline }
    Test-Command { (Get-HTMLString -Path $filepath -Tag "DD") | % { Write-Host $_ -ForegroundColor Yellow } }


    # Invalid Parameter - File Not Found
    Write-Host
    Write-Host '$filepath =' ($filepath = ($testdata_FolderPath | Join-Path -ChildPath "dummy.html"))
    Write-Host '$filename =' ($filename = Split-Path -Path $filepath -Leaf)

    Test-Command (MESSAGE Get-HTMLString, (++$i), $filename) {
        try
        {
            Get-HTMLString -Path $filepath -Tag dummy
            return $false
        }
        catch
        {
            Write-Warning $_
            if (([System.Management.Automation.ErrorRecord]$_).Exception.InnerException.InnerException -is [System.IO.FileNotFoundException]) { return $true }
            else { return $false }
        }
    }

    # Invalid Parameter - Target is not File (is Folder).
    Write-Host
    Write-Host '$filepath =' ($filepath = $testdata_FolderPath) '(Directory)'

    Test-Command (MESSAGE Get-HTMLString, (++$i), (Split-Path -Path $filepath -Leaf)) {
        try
        {
            Get-HTMLString -Path $filepath -Tag dummy
            return $false
        }
        catch
        {
            Write-Warning $_
            if (([System.Management.Automation.ErrorRecord]$_).Exception.InnerException.InnerException -is [System.IO.FileNotFoundException]) { return $true }
            else { return $false }
        }
    }
}

#####################################################################################################################################################
# Get-PrivateProfileString
#
# "sample.inf" is copied from the following URL, on 2013/10/25.
# Sample INF File (Windows Drivers) <http://msdn.microsoft.com/en-us/library/windows/hardware/ff548081.aspx>
#
Test-Module $TargetModule Get-PrivateProfileString {
    $i = 0

    # File 1
    Write-Host
    Write-Host '$filepath =' ($filepath = ($testdata_FolderPath | Join-Path -ChildPath "Test.INI"))
    Write-Host '$filename =' ($filename = Split-Path -Path $filepath -Leaf)
    Write-Verbose (VERBOSE_LINE)
    Write-Verbose (Get-Content -Path $filepath -Raw)
    Write-Verbose (VERBOSE_LINE)

    Test-Command (MESSAGE Get-PrivateProfileString, (++$i), $filename) {
        (Get-PrivateProfileString -Path $filepath -Section Number -Key Two) -eq "2"
    }

    Test-Command (MESSAGE Get-PrivateProfileString, (++$i), $filename) {
        (Get-PrivateProfileString -Path $filepath -Section Number -Key Three) -eq "3"
    }

    Test-Command (MESSAGE Get-PrivateProfileString, (++$i), $filename) {
        (Get-PrivateProfileString -Path $filepath -Section Alphabet -Key First) -eq "ABC"
    }

    Test-Command (MESSAGE Get-PrivateProfileString, (++$i), $filename) {
        (Get-PrivateProfileString -Path $filepath -Section Alphabet -Key Last) -eq "XYZ"
    }

    Test-Command (MESSAGE Get-PrivateProfileString, (++$i), $filename) {
        (Get-PrivateProfileString -Path $filepath -Section KANJI -Key NotFound) -eq [string]::Empty
    }

    Test-Command (MESSAGE Get-PrivateProfileString, (++$i), $filename) {
        (Get-PrivateProfileString -Path $filepath -Section Alphabet -Key NotFound) -eq [string]::Empty
    }


    # File 2
    Write-Host
    Write-Host '$filepath =' ($filepath = ($testdata_FolderPath | Join-Path -ChildPath "Sample.INF"))
    Write-Host '$filename =' ($filename = Split-Path -Path $filepath -Leaf)
    Write-Verbose (VERBOSE_LINE)
    Write-Verbose (Get-Content -Path $filepath -Raw)
    Write-Verbose (VERBOSE_LINE)

    Test-Command (MESSAGE Get-PrivateProfileString, (++$i), $filename) {
        (Get-PrivateProfileString -Path $filepath -Section Version -Key DriverVer) -eq "MM/DD/YYYY,n.n.n.n"
    }

    Test-Command (MESSAGE Get-PrivateProfileString, (++$i), $filename) {
        (Get-PrivateProfileString -Path $filepath -Section Strings -Key USB\MyDevice.DeviceDesc) -eq """My Device Description"""
    }


    # Invalid Parameter - File Not Found
    Write-Host
    Write-Host '$filepath =' ($filepath = ($testdata_FolderPath | Join-Path -ChildPath "dummy.INF"))
    Write-Host '$filename =' ($filename = Split-Path -Path $filepath -Leaf)

    Test-Command (MESSAGE Get-PrivateProfileString, (++$i), $filename) {
        try
        {
            Get-PrivateProfileString -Path $filepath -Section dummy_Section -Key dummy_Key
            return $false
        }
        catch
        {
            Write-Warning $_
            if (([System.Management.Automation.ErrorRecord]$_).Exception.InnerException.InnerException -is [System.IO.FileNotFoundException]) { return $true }
            else { return $false }
        }
    }

    # Invalid Parameter - Target is not File (is Folder).
    Write-Host
    Write-Host '$filepath =' ($filepath = $testdata_FolderPath) '(Directory)'

    Test-Command (MESSAGE Get-PrivateProfileString, (++$i), $filename) {
        try
        {
            Get-PrivateProfileString -Path $filepath -Section dummy_Section -Key dummy_Key
            return $false
        }
        catch
        {
            Write-Warning $_
            if (([System.Management.Automation.ErrorRecord]$_).Exception.InnerException.InnerException -is [System.IO.FileNotFoundException]) { return $true }
            else { return $false }
        }
    }
}

#####################################################################################################################################################
# Update-Content
Test-Module $TargetModule Update-Content {
    $i = 0

    # Print content of the file
    Write-Host
    Write-Host '$filepath =' ($filepath = ($testdata_FolderPath | Join-Path -ChildPath "example.html"))
    Write-Host '$filename =' ($filename = Split-Path -Path $filepath -Leaf)
    Write-Verbose (VERBOSE_LINE)
    Write-Verbose (Get-Content -Path $filepath -Raw)
    Write-Verbose (VERBOSE_LINE)


    Test-Command (MESSAGE Update-Content, (++$i)) {
        Write-Host
        Write-Host (New-HR)
        Update-Content -SearchText "Hello" -UpdateText "Good Morning" -InputObject (Get-Content -Path $filepath) | % { Write-Host $_ -ForegroundColor Yellow }
        Write-Host (New-HR)
    }

    Test-Command (MESSAGE Update-Content, (++$i)) {
        Write-Host
        Write-Host (New-HR)
        Update-Content -Line 8 -UpdateText "`t`t<H1>HELLO, WORLD!</H1>" -InputObject (Get-Content -Path $filepath) | % { Write-Host $_ -ForegroundColor Yellow }
        Write-Host (New-HR)
    }

    Test-Command (MESSAGE Update-Content, (++$i)) {
        (Compare-Object `
            -DifferenceObject (Update-Content "hoge" "Good morning" (Get-Content $filepath)) `
            -ReferenceObject (Get-Content $filepath) `
            -CaseSensitive
        ) -eq $null
    }

    Test-Command (MESSAGE Update-Content, (++$i)) {
        (Compare-Object `
            -DifferenceObject (Get-Content $filepath | Update-Content "Hello" "Good morning") `
            -ReferenceObject (Get-Content $filepath | Update-Content 8 "      <P>Good morning world!") `
            -CaseSensitive
        ) -eq $null
    }

    # Invalid Parameter - Line is over
    Test-Command (MESSAGE Update-Content, (++$i)) {
        try
        {
            Update-Content -Line 20 -UpdateText "dummy" -InputObject (Get-Content -Path $filepath)
            return $false
        }
        catch
        {
            Write-Warning $_
            if (([System.Management.Automation.ErrorRecord]$_).Exception -is [System.ArgumentOutOfRangeException]) { return $true }
            else { return $false }
        }
    }


    # Print content of the file / [+]V1.0.1.0 (2014/05/09)
    Write-Host
    Write-Host '$filepath =' ($filepath = ($testdata_FolderPath | Join-Path -ChildPath "Test.INI"))
    Write-Host '$filename =' ($filename = Split-Path -Path $filepath -Leaf)
    Write-Verbose (VERBOSE_LINE)
    Write-Verbose (Get-Content -Path $filepath -Raw)
    Write-Verbose (VERBOSE_LINE)

    # Including empty line / [+]V1.0.1.0 (2014/05/09)
    Test-Command (MESSAGE Update-Content, (++$i)) {
        Write-Host
        Write-Host (New-HR)
        Update-Content -SearchText "ABC" -UpdateText "EFG" -InputObject (Get-Content -Path $filepath) | % { Write-Host $_ -ForegroundColor Yellow }
        Write-Host (New-HR)
    }
}

#####################################################################################################################################################
# Invoke-LoadLibraryEx
Test-Module $TargetModule Invoke-LoadLibraryEx {

    Write-Host '$filepath =' ($filepath = ($testdata_FolderPath | Join-Path -ChildPath "win32rsc.dll"))
    Write-Host '$filename =' ($filename = Split-Path -Path $filepath -Leaf)

    Test-Command { Write-Host (Invoke-LoadString (Invoke-LoadLibraryEx $filepath) 201) -ForegroundColor Yellow -NoNewline }
}

#####################################################################################################################################################
# Get-ResourceString
Test-Module $TargetModule Get-ResourceString {

    Write-Host '$filepath =' ($filepath = ($testdata_FolderPath | Join-Path -ChildPath "win32rsc.dll"))
    Write-Host '$filename =' ($filename = Split-Path -Path $filepath -Leaf)

    Test-Command (MESSAGE Get-ResourceString) {
        Write-Host
        (Get-ResourceString -Path $filepath -uID (201..203)) | Write-Host -ForegroundColor Yellow
    }
}

#####################################################################################################################################################
# Invoke-HtmlHelp
Test-Module $TargetModule Invoke-HtmlHelp {
    $i = 0

    Write-Host '$filepath =' ($filepath = "C:\Windows\Help\mui\0411\mmc.CHM")
    Write-Host '$filename =' ($filename = Split-Path -Path $filepath -Leaf)

    Test-Command (MESSAGE Invoke-HtmlHelp, (++$i)) { Invoke-HtmlHelp -Path $filepath -hwndCaller (Get-WindowHandler) }

    Write-Host "Wait 3 second(s)..." -ForegroundColor Magenta
    Start-Sleep -Seconds 3

    Test-Command (MESSAGE Invoke-HtmlHelp, (++$i)) { Invoke-HtmlHelp -uCommand 0x0012 -hwndCaller (Get-WindowHandler) }
}

#####################################################################################################################################################
# Get-WindowHandler
Test-Module $TargetModule Get-WindowHandler {

    Test-Command { Write-Host (Get-WindowHandler) -NoNewline }
}

#####################################################################################################################################################
# New-StructArray
Test-Module $TargetModule New-StructArray {
    $i = 0

    Write-Host
    Test-Command (MESSAGE New-StructArray, (++$i)) {
        Write-Host
        ($obj = New-StructArray -Members a=1,b=2,c=3,x,y,z) | Get-Member -MemberType NoteProperty | Write-Host

        Write-Host ".ToString() = " -NoNewline
        $obj.ToString() | Write-Host -ForegroundColor Yellow
    }

    Write-Host
    Test-Command (MESSAGE New-StructArray, (++$i)) {
        Write-Host
        ($obj = New-StructArray -Members first=1,second=2,third=XYZ -Count 3) | % { $_ | Get-Member -MemberType NoteProperty | Write-Host }

        Write-Host ".ToString() = " -NoNewline
        $obj | % { $_.ToString("Table border=1", "Item", "Name", "Value") | Write-Host -ForegroundColor Yellow }
    }

    Write-Host
    Test-Command (MESSAGE New-StructArray, (++$i)) {
        Write-Host
        ($obj = New-StructArray -Members Label=Language, Version=V1.0.0.0 -Count 2) | % { $_ | Get-Member -MemberType NoteProperty | Write-Host }

        Write-Host ".ToString() = " -NoNewline
        $obj | % { $_.ToString("Table border=1", "tr") | Write-Host -ForegroundColor Yellow }
    }
}

#####################################################################################################################################################
# ConvertFrom-ByteArray
Test-Module $TargetModule ConvertFrom-ByteArray {
    $i = 0

    Write-Host
    Write-Host '$filename =' ($filename = "test.bin")
    Write-Host '$filepath =' ($filepath = ($testdata_FolderPath | Join-Path -ChildPath $filename))

    $bin = Get-ByteArray -Path $filepath

    Write-Host
    Test-Command (MESSAGE ConvertFrom-ByteArray, (++$i)) { ConvertFrom-ByteArray -InputObject $bin | Write-Host -ForegroundColor Yellow -NoNewline }
    Test-Command (MESSAGE ConvertFrom-ByteArray, (++$i)) { ConvertFrom-ByteArray -InputObject $bin -Hex | Write-Host -ForegroundColor Yellow -NoNewline }
    Test-Command (MESSAGE ConvertFrom-ByteArray, (++$i)) { $bin | ConvertFrom-ByteArray -Separator ',' | Write-Host -ForegroundColor Yellow -NoNewline }
    Test-Command (MESSAGE ConvertFrom-ByteArray, (++$i)) { $bin | ConvertFrom-ByteArray -Separator '-' -Hex | Write-Host -ForegroundColor Yellow -NoNewline }

    Write-Host
    Test-Command (MESSAGE ConvertFrom-ByteArray, (++$i)) {
        (ConvertFrom-ByteArray -InputObject $bin) -eq '0.1.2.3.4.5.6.7.8.9.10.11.12.13.14.15.240.241.242.242.243.245.246.247.248.249.250.251.252.253.254.255'
    }

    Test-Command (MESSAGE ConvertFrom-ByteArray, (++$i)) {
        (ConvertFrom-ByteArray -InputObject $bin -Hex) -eq '00:01:02:03:04:05:06:07:08:09:0A:0B:0C:0D:0E:0F:F0:F1:F2:F2:F3:F5:F6:F7:F8:F9:FA:FB:FC:FD:FE:FF'
    }
}

#####################################################################################################################################################
# ConvertTo-ByteArray
Test-Module $TargetModule ConvertTo-ByteArray {
    $i = 0
    $bin1 = $null
    $bin2 = $null

    # Decimal
    Write-Host
    Test-Command (MESSAGE ConvertTo-ByteArray, (++$i)) {

        $text = '1.3.6.1.2.1.1.5' # sysName
        $bin = ConvertTo-ByteArray -InputObject $text

        Write-Host
        Write-Host "Input = '" -NoNewline
        Write-Host $text -ForegroundColor Yellow -NoNewline
        Write-Host "' (" -NoNewline
        Write-Host $text.GetType() -ForegroundColor Yellow -NoNewline
        Write-Host ")  //sysName"

        Write-Host 'Output = {' -NoNewline
        foreach ($j in 0..($bin.Count - 1))
        {
            if ($j -ne 0) { Write-Host ', ' -NoNewline }
            Write-Host $bin[$j] -ForegroundColor Yellow -NoNewline
        }
        Write-Host '} (' -NoNewline
        Write-Host $bin.GetType() -ForegroundColor Yellow -NoNewline
        Write-Host ')'

        Write-Host 'Type of Output = {' -NoNewline
        foreach ($j in 0..($bin.Count - 1))
        {
            if ($j -ne 0) { Write-Host ', ' -NoNewline }
            Write-Host '(' -NoNewline
            Write-Host $bin[$j].GetType() -ForegroundColor Yellow -NoNewline
            Write-Host ')' -NoNewline
        }
        Write-Host '}'

        $Script:temp1 = $bin
    }


    # Hexadecimal
    Write-Host
    Test-Command (MESSAGE ConvertTo-ByteArray, (++$i)) {

        $text = '00:01:02:EE:EF'
        $bin = ConvertTo-ByteArray -InputObject $text -Hex

        Write-Host
        Write-Host "Input = '" -NoNewline
        Write-Host $text -ForegroundColor Yellow -NoNewline
        Write-Host "' (" -NoNewline
        Write-Host $text.GetType() -ForegroundColor Yellow -NoNewline
        Write-Host ")"

        Write-Host 'Output = {' -NoNewline
        foreach ($j in 0..($bin.Count - 1))
        {
            if ($j -ne 0) { Write-Host ', ' -NoNewline }
            Write-Host $bin[$j] -ForegroundColor Yellow -NoNewline
        }
        Write-Host '} (' -NoNewline
        Write-Host $bin.GetType() -ForegroundColor Yellow -NoNewline
        Write-Host ')'

        Write-Host 'Output = {' -NoNewline
        foreach ($j in 0..($bin.Count - 1))
        {
            if ($j -ne 0) { Write-Host ', ' -NoNewline }
            Write-Host '(' -NoNewline
            Write-Host $bin[$j].GetType() -ForegroundColor Yellow -NoNewline
            Write-Host ')' -NoNewline
        }
        Write-Host '}'

        $Script:temp2 = $bin
    }

    Write-Host
    Test-Command (MESSAGE ConvertTo-ByteArray, (++$i)) {

        $bin = ConvertTo-ByteArray -InputObject '01-03-06-01-02-01-01-05' -Separator '-'

        if ($bin.Count -ne $Script:temp1.Count) { return $false }
        foreach ($j in 0..($bin.Count - 1))
        {
             if ($bin[$j] -ne $Script:temp1[$j]) { return $false }
        }
        return $true
    }

    Test-Command (MESSAGE ConvertTo-ByteArray, (++$i)) {

        $bin = ConvertTo-ByteArray -InputObject '00-01-02-EE-EF' -Separator '-' -Hex

        if ($bin.Count -ne $Script:temp2.Count) { return $false }
        foreach ($j in 0..($bin.Count - 1))
        {
             if ($bin[$j] -ne $Script:temp2[$j]) { return $false }
        }
        return $true
    }
}

#####################################################################################################################################################
# Get-MD5
Test-Module $TargetModule Get-MD5 {
    $i = 0

    Write-Host
    Write-Host '$filename1 =' ($filename1 = "win32rsc.dll")
    Write-Host '$filepath1 =' ($filepath1 = ($testdata_FolderPath | Join-Path -ChildPath $filename1))


    # 1 File
    Write-Host
    Test-Command (MESSAGE Get-MD5, (++$i)) { Get-MD5 -Path $filepath1 | Write-Host -ForegroundColor Yellow -NoNewline }
    Test-Command (MESSAGE Get-MD5, (++$i)) { Get-MD5 -Path $filepath1 -FileName | Write-Host -ForegroundColor Yellow -NoNewline }
    Test-Command (MESSAGE Get-MD5, (++$i)) { Get-MD5 -Path $filepath1 -FullName | Write-Host -ForegroundColor Yellow -NoNewline }
    Test-Command (MESSAGE Get-MD5, (++$i)) { (Get-MD5 -Path $filepath1) -eq (Get-CheckSum -InputObject $filepath1) }


    Write-Host
    Write-Host '$filename2 =' ($filename2 = "TheProject.html")
    Write-Host '$filepath2 =' ($filepath2 = ($testdata_FolderPath | Join-Path -ChildPath $filename2))


    # 2 Files
    Write-Host
    Test-Command (MESSAGE Get-MD5, (++$i)) {
        Write-Host
        Get-MD5 $filepath1, $filepath2 | Write-Host -ForegroundColor Yellow
    }

    Test-Command (MESSAGE Get-MD5, (++$i)) {
        Write-Host
        Get-MD5 $filepath1, $filepath2 -FileName | Write-Host -ForegroundColor Yellow
    }

    Test-Command (MESSAGE Get-MD5, (++$i)) {
        Write-Host
        Get-MD5 $filepath1, $filepath2 -FullName | Write-Host -ForegroundColor Yellow
    }


    # Folder
    Write-Host
    Test-Command (MESSAGE Get-MD5, (++$i)) {
        Write-Host
        Get-MD5 -Path $testdata_FolderPath | Write-Host -ForegroundColor Yellow
    }

    Write-Host
    Test-Command (MESSAGE Get-MD5, (++$i)) {
        Write-Host
        Get-MD5 -Path $testdata_FolderPath -FileName | Write-Host -ForegroundColor Yellow
    }

    Write-Host
    Test-Command (MESSAGE Get-MD5, (++$i)) {
        Write-Host
        Get-MD5 -Path $testdata_FolderPath -FullName | Write-Host -ForegroundColor Yellow
    }


    # File Not Found
    Write-Host
    Test-Command (MESSAGE Get-MD5, (++$i)) {
        try
        {
            Get-MD5 -Path ($filepath1 | Join-Path -ChildPath 'dummy')
            return $false
        }
        catch
        {
            Write-Warning $_
            if (([System.Management.Automation.ErrorRecord]$_).Exception.InnerException.InnerException -is [System.IO.FileNotFoundException]) { return $true }
            else { return $false }
        }
    }

    # Empty Folder
    Write-Host
    Test-Command (MESSAGE Get-MD5, (++$i)) { (Get-MD5 -Path ($testdata_FolderPath | Join-Path -ChildPath 'empty')) -eq $null }
}

#####################################################################################################################################################
# Test-SameFile
Test-Module $TargetModule Test-SameFile {
    $i = 0

    Test-Command (MESSAGE Test-SameFile, (++$i)) {
        (Test-SameFile `
            -ReferenceObject ($testdata_FolderPath | Join-Path -ChildPath "abc (lower).txt") `
            -DifferenceObject ($testdata_FolderPath | Join-Path -ChildPath "ABC (UPPER).txt")
        ) -eq $false
    }

    Test-Command (MESSAGE Test-SameFile, (++$i)) {
        (Test-SameFile `
            -ReferenceObject ($testdata_FolderPath | Join-Path -ChildPath "abc (lower).txt") `
            -DifferenceObject ($testdata_FolderPath | Join-Path -ChildPath "copy of abc (lower).txt")
        ) -eq $true
    }

    # Folder (same)
    Test-Command (MESSAGE Test-SameFile, (++$i)) {
        (Test-SameFile -ReferenceObject $testdata_FolderPath -DifferenceObject $testdata_FolderPath) -eq $true
    }

    # Folder (different)
    Test-Command (MESSAGE Test-SameFile, (++$i)) {
        (Test-SameFile `
            -ReferenceObject ($testdata_FolderPath | Join-Path -ChildPath 'ISOImageFile') `
            -DifferenceObject ($testdata_FolderPath | Join-Path -ChildPath 'ZipFile')
        ) -eq $false
    }
}

#####################################################################################################################################################
# Start-Command
Test-Module $TargetModule Start-Command {
    $i = 0

    Test-Command (MESSAGE Start-Command, (++$i)) {
        Write-Host "ダイアログの OK ボタンをクリックしてください。" -ForegroundColor Magenta -NoNewline
        if ($Verbose) { Write-Host }
        Start-Command -FilePath winver
    }

    Test-Command (MESSAGE Start-Command, (++$i)) {
        Write-Host
        $id = Start-Command `
            -FilePath notepad `
            -WorkingDirectory $testdata_FolderPath `
            -ArgumentList "このウィンドウは約3秒後に自動的に閉じられます.txt" `
            -Async
        Write-Host "Wait 3 second(s)..." -ForegroundColor Magenta
        Start-Sleep -Seconds 3
        (Get-Process -Id $id).Kill()
    }

    Test-Command (MESSAGE Start-Command, (++$i)) {
        $id = 256 + 1
        Write-Host

        try
        {
            Start-Command cmd "/K", "exit $id" -Retry 3 -Interval 2
            return $false
        }
        catch
        {
            if ($Error[0].FullyQualifiedErrorId -eq ("0x" + ($id).ToString("x8"))) { return $true }
            else { return $false }
        }
    }
}

#####################################################################################################################################################
# New-ISOImageFile
Test-Module $TargetModule New-ISOImageFile {
    $i = 0

    $iso_Input_FolderPath = ($testdata_FolderPath | Join-Path -ChildPath "ISOImageFile" | Join-Path -ChildPath "root")
    $iso_OutputFile_BaseNamePath = ($testdata_FolderPath | Join-Path -ChildPath "ISOImageFile" | Join-Path -ChildPath "test")

    Write-Host
    Write-Host "Input Folder Path =" $iso_Input_FolderPath


    $iso_Output_FilePath = $iso_OutputFile_BaseNamePath + "("+ (++$i) + ").iso"
    Write-Host
    Write-Host ("(Output) ISO File Name (" + $i + ") = " + ($iso_Output_FilePath | Split-Path -Leaf))
    Test-Command (MESSAGE New-ISOImageFile, $i) {
        Write-Host
        New-ISOImageFile `
            -InputObject $iso_Input_FolderPath `
            -Path $iso_Output_FilePath `
            -VolumeID ("This is MAX Length of Vol-ID (" + $i + ")") `
            -BinPath C:\cygwin64\bin `
        | Write-Host -ForegroundColor Yellow
    }

    $iso_Output_FilePath = $iso_OutputFile_BaseNamePath + "("+ (++$i) + ").iso"
    Write-Host
    Write-Host ("(Output) ISO File Name (" + $i + ") = " + ($iso_Output_FilePath | Split-Path -Leaf))
    Test-Command (MESSAGE New-ISOImageFile, $i) {
        Write-Host
        New-ISOImageFile `
            -InputObject $iso_Input_FolderPath `
            -Path $iso_Output_FilePath `
            -VolumeID ("Test Image (" + $i + ")") `
            -Publisher "BUILDLet" `
            -ApplicationID "Package Builder Toolkit for PowerShell" `
            -Recommended `
            -RedirectStandardError `
        | Write-Host -ForegroundColor Yellow
    }

    $iso_Output_FilePath = $iso_OutputFile_BaseNamePath + "("+ (++$i) + ").iso"
    Write-Host
    Write-Host ("(Output) ISO File Name (" + $i + ") = " + ($iso_Output_FilePath | Split-Path -Leaf))
    Test-Command (MESSAGE New-ISOImageFile, $i) {
        Write-Host
        New-ISOImageFile `
            -InputObject $iso_Input_FolderPath `
            -Path $iso_Output_FilePath `
            -VolumeID ("Test Image (" + $i + ")") `
            -ArgumentList @(
                "-publisher `"BUILDLet.com`"",
                ("-volid `"Overwritten Volume ID (" + $i + ")`""),
                "-appid `"This is ISO Image Test File.`"",
                "-omit-period",
                "-disable-deep-relocation",
                "-output-charset ascii",
                "-full-iso9660-filenames",
                "-allow-limited-size",
                "-allow-leading-dots",
                "-no-iso-translate",
                "-allow-lowercase",
                "-allow-multidot"
            ) `
            -RedirectStandardError `
        | Write-Host -ForegroundColor Yellow
    }


    # Default Settings
    Write-Host
    Write-Host ("(Output) ISO File Name (" + (++$i) + ") = Defaualt ('" + ($iso_Input_FolderPath | Split-Path -Leaf) + ".iso')")
    Test-Command (MESSAGE New-ISOImageFile, $i) {
        Write-Host
        New-ISOImageFile -InputObject $iso_Input_FolderPath | Write-Host -ForegroundColor Yellow
    }

    # Recommended
    $iso_Output_FilePath = $iso_Input_FolderPath | Split-Path -Parent | Join-Path -ChildPath "Recommended.iso"
    Write-Host
    Write-Host ("(Output) ISO File Name (" + (++$i) + ") = " + ($iso_Output_FilePath | Split-Path -Leaf))
    Test-Command (MESSAGE New-ISOImageFile, $i) {
        Write-Host
        New-ISOImageFile `
            -InputObject $iso_Input_FolderPath `
            -Path $iso_Output_FilePath `
            -VolumeID "Recommended" `
            -Recommended `
            | Write-Host -ForegroundColor Yellow
    }


    # Parameter is invalid. - Output Folder does not exist.
    $iso_Output_FilePath = $iso_OutputFile_BaseNamePath | Split-Path -Parent | Join-Path -ChildPath "dummy" | Join-Path -ChildPath "dummy.iso"
    Write-Host
    Write-Host ("(Output) ISO File Name (" + (++$i) + ") = " + ($iso_Output_FilePath | Split-Path -Leaf))
    Test-Command (MESSAGE New-ISOImageFile, $i) {
        Write-Host
        try { 
            New-ISOImageFile `
                -InputObject $iso_Input_FolderPath `
                -Path $iso_Output_FilePath `
                -VolumeID ("Test Image (" + $i + ")") `
                -Recommended

            return $false
        }
        catch
        {
            Write-Warning $_
            if (([System.Management.Automation.ErrorRecord]$_).Exception.InnerException.InnerException -is [System.IO.DirectoryNotFoundException]) { return $true }
            else { return $false }
        }
    }

    # Parameter is invalid. - Input Folder is not found.
    $iso_Input_FolderPath = $iso_Input_FolderPath | Split-Path -Parent | Join-Path -ChildPath "dummy"
    $iso_Output_FilePath = $iso_OutputFile_BaseNamePath | Split-Path -Parent | Join-Path -ChildPath "dummy.iso"
    Write-Host
    Write-Host ("(Output) ISO File Name (" + (++$i) + ") = " + ($iso_Output_FilePath | Split-Path -Leaf))
    Test-Command (MESSAGE New-ISOImageFile, $i) {
        Write-Host
        try { 
            New-ISOImageFile `
                -InputObject $iso_Input_FolderPath `
                -Path $iso_Output_FilePath `
                -VolumeID ("Test Image (" + $i + ")") `
                -Recommended

            return $false
        }
        catch
        {
            Write-Warning $_
            if (([System.Management.Automation.ErrorRecord]$_).Exception.InnerException.InnerException -is [System.IO.DirectoryNotFoundException]) { return $true }
            else { return $false }
        }
    }

}

#####################################################################################################################################################
# Start-Computer
Test-Module $TargetModule Start-Computer {

    Test-Command Start-Computer {
        if (-not (Test-Connection $test_Hostname -Quiet -Verbose))
        {
            Write-Host
            Start-Computer -MacAddress $test_MacAddress -Verbose

            # Wait (after Waking up)
            Write-Host ("Waiting " + ($interval = 30) + " seconds...")
            Start-Sleep -Seconds $interval
        }
        else { Write-Warning "Test of Start-Computer Command is skipped, because host '$test_Hostname' is already alive." }
    }
}

#####################################################################################################################################################
# Restart-Host
Test-Module $TargetModule Restart-Host {
    $i = 0

    # Wake up the host
    if (-not (Test-Connection $test_Hostname -Quiet))
    {
        Write-Host
        Start-Computer -MacAddress $test_MacAddress -Verbose

        # Wait (after Waking up)
        Write-Host ("Waiting " + ($interval = 30) + " seconds...")
        Start-Sleep -Seconds $interval
    }

    # Show message to log on the host
    Write-Host "Please log on the host '$test_Hostname'." -ForegroundColor Magenta


    Test-Command (MESSAGE Restart-Host, (++$i)) { Restart-Host -ComputerName $test_Hostname -UserName $test_UserName -Password $test_Password }

    Test-Command (MESSAGE Restart-Host, (++$i)) {
        try { Restart-Host -ComputerName $test_Hostname -UserName $test_UserName -Password $test_Password -Silent }
        catch { Write-Warning $_ }
    }

    Test-Command (MESSAGE Restart-Host, (++$i)) { Restart-Host -ComputerName $test_Hostname -UserName $test_UserName -Password $test_Password -Silent -Force }

    # Wait (after Rebooting)
    Write-Host ("Waiting " + ($interval = 30) + " seconds...")
    Start-Sleep -Seconds $interval
}

#####################################################################################################################################################
# Stop-Host
Test-Module $TargetModule Stop-Host {
    $i = 0

    # Wake up the host
    if (-not (Test-Connection $test_Hostname -Quiet))
    {
        Write-Host
        Start-Computer -MacAddress $test_MacAddress -Verbose
    }

    # Show message to log on the host
    Write-Host "Please log on the host '$test_Hostname'." -ForegroundColor Magenta


    Test-Command (MESSAGE Stop-Host, (++$i)) { Stop-Host -ComputerName $test_Hostname -UserName $test_UserName -Password $test_Password }

    Test-Command (MESSAGE Stop-Host, (++$i)) {
        try { Stop-Host -ComputerName $test_Hostname -UserName $test_UserName -Password $test_Password -Silent }
        catch { Write-Warning $_ }
    }

    Test-Command (MESSAGE Stop-Host, (++$i)) { Stop-Host -ComputerName $test_Hostname -UserName $test_UserName -Password $test_Password -Silent -Force }
}

#####################################################################################################################################################
# Invoke-LoadLibrary
Test-Module $TargetModule Invoke-LoadLibrary {

    Write-Host '$filepath =' ($filepath = "C:\Windows\System32\calc.exe")
    Write-Host '$filename =' ($filename = Split-Path -Path $filepath -Leaf)

    Test-Command { Write-Host (Invoke-LoadString (Invoke-LoadLibrary $filepath) 3) -ForegroundColor Yellow -NoNewline }
    #【 注意 】
    # (カレントディレクトリの DLL (EXE) を見つけられないため、システムフォルダ内にある EXE を指定。)
}

#####################################################################################################################################################
# Get-CheckSum
Test-Module $TargetModule Get-CheckSum {
    $i = 0

    Write-Host
    Write-Host '$filename =' ($filename = "win32rsc.dll")
    Write-Host '$filepath =' ($filepath = ($testdata_FolderPath | Join-Path -ChildPath $filename))

    # MD5
    if ($Verbose) { Write-Host }
    Test-Command (MESSAGE Get-CheckSum, (++$i)) {
        if ($Verbose) { Write-Host }
        Write-Host ("'" + $filename + "' = ") -NoNewline
        Get-CheckSum -InputObject $filepath -BinPath ($current_FolderPath | Join-Path -ChildPath "FCIV") `
            | Write-Host -ForegroundColor Yellow -NoNewline
        Write-Host " (MD5)" -NoNewline
    }

    if ($Verbose) { Write-Host }
    Test-Command (MESSAGE Get-CheckSum, (++$i)) {
        $checksum = Get-CheckSum $filepath
        return ($checksum -eq "3d7a2b1a00bfb15be9448eaf8ee045b8")
    }

    # MD5 (Folder)
    if ($Verbose) { Write-Host }
    Test-Command (MESSAGE Get-CheckSum, (++$i)) {
        Write-Host
        Write-Host "'$testdata_FolderPath' (MD5) ="
        Get-CheckSum -InputObject $testdata_FolderPath -BinPath ($current_FolderPath | Join-Path -ChildPath "FCIV") `
            | % { Write-Host "`t$_" -ForegroundColor Yellow }
    }


    # SHA1
    if ($Verbose) { Write-Host }
    Test-Command (MESSAGE Get-CheckSum, (++$i)) {
        if ($Verbose) { Write-Host }
        Write-Host ("'" + $filename + "' = ") -NoNewline
        Get-CheckSum -InputObject $filepath -BinPath ($current_FolderPath | Join-Path -ChildPath "FCIV") -SHA1 `
            | Write-Host -ForegroundColor Yellow -NoNewline
        Write-Host " (SHA1)" -NoNewline
    }

    if ($Verbose) { Write-Host }
    Test-Command (MESSAGE Get-CheckSum, (++$i)) {
        $checksum = Get-CheckSum $filepath -SHA1
        return ($checksum -eq "3b46316b2892e39ef385292bbb4a79465baa53ce")
    }

    # SHA1 (Folder)
    if ($Verbose) { Write-Host }
    Test-Command (MESSAGE Get-CheckSum, (++$i)) {
        Write-Host
        Write-Host "'$testdata_FolderPath' (SHA1) ="
        Get-CheckSum -InputObject $testdata_FolderPath -BinPath ($current_FolderPath | Join-Path -ChildPath "FCIV") -SHA1 `
            | % { Write-Host "`t$_" -ForegroundColor Yellow }
    }


    # Parameter is invalid. - File is not found.
    if ($Verbose) { Write-Host }
    Test-Command (MESSAGE Get-CheckSum, (++$i)) {
        if ($Verbose) { Write-Host }
        try
        {
            Get-CheckSum -InputObject ($filepath | Join-Path -ChildPath "dummy") -BinPath ($current_FolderPath | Join-Path -ChildPath "FCIV")
            return $false
        }
        catch
        {
            Write-Warning $_
            if (([System.Management.Automation.ErrorRecord]$_).Exception.InnerException.InnerException -is [System.IO.FileNotFoundException]) { return $true }
            else { return $false }
        }
    }
}

#####################################################################################################################################################
# Send-Mail
Test-Module $TargetModule Send-Mail {

    Write-Host Send-MailMessage "コマンドレットがあるため実施しません。"
}


#####################################################################################################################################################
# Unit Tests for ZipFile Module
#####################################################################################################################################################

#####################################################################################################################################################
# Expand-ZipFile
Test-Module $TargetModule Expand-ZipFile {


    # Load 'WindowsBase' Assembly for handling 'System.IO.FileFormatException'
    try {
        [System.Reflection.Assembly]::Load('WindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35')
    }
    catch {
        [System.Reflection.Assembly]::LoadWithPartialName('WindowsBase.dll')
    }


    for ($i = 0; $i -lt 2; $i++)
    {
        switch ($i)
        {
            0 { $zip_Destination_FolderName = "Decompressed" }
            1 { $zip_Destination_FolderName = "Decompressed (Shell mode)" }
        }

        Write-Host
        Write-Host
        Write-Host 'Source Directory      =' (
            $zip_Source_FolderPath = ($testdata_FolderPath | Join-Path -ChildPath "ZipFile" | Join-Path -ChildPath "Archive")
        )
        Write-Host 'Destination Directory =' (
            $zip_Destination_FolderPath = ($testdata_FolderPath | Join-Path -ChildPath "ZipFile" | Join-Path -ChildPath $zip_Destination_FolderName)
        )

        if (-not (Test-Path -Path $zip_Destination_FolderPath))
        {
            if ($Verbose) { New-Item -Path $zip_Destination_FolderPath -ItemType Directory -Force -Verbose }
            else { New-Item -Path $zip_Destination_FolderPath -ItemType Directory -Force }
        }

        for ($j = 1; $j -le 7; $j++)
        {
            switch ($j)
            {
                1 { $zip_filename = "01 a file (test).zip" }
                2 { $zip_filename = "02 3 files including Japanese filename.zip" }
                3 { $zip_filename = "03 a file in a folder (Lenna).zip"}
                4 { $zip_filename = "04 3 files in a folder (RGB).zip"}
                5 { $zip_filename = "05 misc.zip"}
                6 { $zip_filename = "06 text file.zip"}
                7 { $zip_filename = "07 empty.zip"}
            }

            Write-Host
            Write-Host '$zip_filepath =' ($zip_filepath = ($zip_Source_FolderPath | Join-Path -ChildPath $zip_filename))
            Write-Host '$zip_filename =' $zip_filename

            if ($j -le 5)
            {
                switch ($i)
                {
                    0 # Decompression by .NET Framework
                    {
                        # w/o -Force Option
                        Test-Command (MESSAGE Expand-ZipFile, $j) {
                            Write-Host
                            Expand-ZipFile `
                                -InputObject $zip_filepath `
                                -Path ($zip_Destination_FolderPath | Join-Path -ChildPath "Expand-ZipFile ($j)") `
                            | Write-Host -ForegroundColor Yellow
                        }

                        # w/ -Force Option
                        Test-Command (MESSAGE Expand-ZipFile, $j, "w/ Force option") {
                            Write-Host
                            Expand-ZipFile `
                                -InputObject $zip_filepath `
                                -Path ($zip_Destination_FolderPath | Join-Path -ChildPath "Expand-ZipFile ($j)") -Force `
                            | Write-Host -ForegroundColor Yellow
                        }
                    }

                    1 # Decompression by Shell mode
                    {
                        # Test only w/ -Force Option
                        Test-Command (MESSAGE Expand-ZipFile, $j) {
                            Write-Host
                            Expand-ZipFile `
                                -InputObject $zip_filepath `
                                -Path ($zip_Destination_FolderPath | Join-Path -ChildPath "Expand-ZipFile ($j)") -ShellMode -Force `
                            | Write-Host -ForegroundColor Yellow
                        }
                    }
                }
            }
            elseif ($j -eq 6)
            {
                switch ($i)
                {
                    0
                    {
                        # Zip File is not found.
                        Test-Command (MESSAGE Expand-ZipFile, $j) {
                            Write-Host
                            try
                            {
                                Expand-ZipFile -InputObject $zip_filepath -Path $zip_Destination_FolderPath | Write-Host -ForegroundColor Yellow
                                return $false
                            }
                            catch
                            {
                                Write-Warning $_
                                if (([System.Management.Automation.ErrorRecord]$_).Exception.InnerException.InnerException -is [System.IO.FileFormatException])
                                {
                                    return $true
                                }
                                else { return $false }
                            }
                        }
                    }
                    1 { Write-Warning "Test of Expand-ZipFile for '$zip_filename' is skipped." }
                }
            }
            elseif ($j -eq 7)
            {
                switch ($i)
                {
                    0
                    {
                        # Target is not File (is Folder).
                        Test-Command (MESSAGE Expand-ZipFile, $j) {
                            Write-Host
                            try
                            {
                                Expand-ZipFile -InputObject $zip_Source_FolderPath -Path $zip_Destination_FolderPath | Write-Host -ForegroundColor Yellow
                                return $false
                            }
                            catch
                            {
                                Write-Warning $_
                                if (([System.Management.Automation.ErrorRecord]$_).Exception.InnerException.InnerException -is [System.IO.FileNotFoundException])
                                {
                                    return $true
                                }
                                else { return $false }
                            }
                        }
                    }
                    1 { Write-Warning "Test of Expand-ZipFile for '$zip_filename' is skipped." }
                }
            }
            else { throw }
        }
    }
}

#####################################################################################################################################################
# New-ZipFile
Test-Module $TargetModule New-ZipFile {

    for ($i = 0; $i -lt 3; $i++)
    {
        switch ($i)
        {
            0 { $zip_Destination_FolderName = "Compressed" }
            1 { $zip_Destination_FolderName = "Compressed (Shell mode)" }

            # [+]V1.1.0.0 (2014/05/24)
            2 { $zip_Destination_FolderName = "Compressed (Additinal)" }
        }

        Write-Host
        Write-Host
        Write-Host 'Source Directory      =' (
            $zip_Source_FolderPath = ($testdata_FolderPath | Join-Path -ChildPath "ZipFile" | Join-Path -ChildPath "Decompressed")
        )
        Write-Host 'Destination Directory =' (
            $zip_Destination_FolderPath = ($testdata_FolderPath | Join-Path -ChildPath "ZipFile" | Join-Path -ChildPath $zip_Destination_FolderName)
        )

        if (-not (Test-Path -Path $zip_Destination_FolderPath))
        {
            if ($Verbose) { New-Item -Path $zip_Destination_FolderPath -ItemType Directory -Force -Verbose }
            else { New-Item -Path $zip_Destination_FolderPath -ItemType Directory -Force }
        }


        # Check for input (This comment is added at 2014/05/23.)
        if ($i -lt 2)
        {
            for ($j = 1; $j -le 5; $j++)
            {
                switch ($j)
                {
                    1 { $zip_leafpath = "01 a file (test)\test.bmp" }
                    2 { $zip_leafpath = "02 3 files including Japanese filename\日本語ファイル名.pdf" }
                    3 { $zip_leafpath = "03 a file in a folder (Lenna)"}
                    4 { $zip_leafpath = "04 3 files in a folder (RGB)"}
                    5 { $zip_leafpath = "05 misc"}
                }
                $zip_filepath = ($zip_Source_FolderPath | Join-Path -ChildPath "Expand-ZipFile ($j)" | Join-Path -ChildPath $zip_leafpath)

                Write-Host
                Write-Host '$zip_filepath =' $zip_filepath
                Write-Host '$zip_leafpath =' $zip_leafpath

                switch ($i)
                {
                    0 # Compression by .NET Framework
                    {
                        # w/o -Force Option
                        Test-Command (MESSAGE New-ZipFile, $j) {
                            Write-Host
                            (New-ZipFile -InputObject $zip_filepath -Path $zip_Destination_FolderPath) | Write-Host -ForegroundColor Yellow
                        }

                        # w/ -Force Option
                        Test-Command (MESSAGE New-ZipFile, $j, "w/ Force option") {
                            Write-Host
                            (New-ZipFile -InputObject $zip_filepath -Path $zip_Destination_FolderPath -Force) | Write-Host -ForegroundColor Yellow
                        }
                    }
                    1 # Compression by Shell mode
                    {
                        if ($j -le 2)
                        {
                            # File
                            Test-Command (MESSAGE New-ZipFile, "Shell mode", $j) {
                                Write-Host
                                (New-ZipFile -InputObject $zip_filepath -Path $zip_Destination_FolderPath -ShellMode) | Write-Host -ForegroundColor Yellow
                            }
                        }
                        else
                        {
                            # Folder
                            Test-Command (MESSAGE New-ZipFile, "Shell mode", $j) {
                                Write-Host
                                try
                                {
                                    New-ZipFile -InputObject $zip_filepath -Path $zip_Destination_FolderPath -ShellMode | Write-Host -ForegroundColor Yellow
                                    return $false
                                }
                                catch
                                {
                                    Write-Warning $Error[0].Exception
                                    if ($Error[0].CategoryInfo.Reason -eq [System.NotSupportedException].Name) { return $true }
                                    else { return $false }
                                }
                            }
                        }
                    }
                }
            }
        }

        # Additional Check for output [+]V1.1.0.0 (2014/05/23)
        else
        {
            foreach ($j in 1..3)
            {
                $k = 0

                switch ($j)
                {
                    1 { $zip_leafpath = "Expand-ZipFile (1)\01 a file (test)\test.bmp" }
                    2 { $zip_leafpath = "Expand-ZipFile (5)\05 misc" }
                    3 { $zip_leafpath = "..\Compressed\日本語ファイル名.zip" }
                }

                # Copy file
                Copy-Item `
                    -Path ($zip_Source_FolderPath | Join-Path -ChildPath $zip_leafpath) `
                    -Destination $zip_Destination_FolderPath `
                    -Recurse `
                    -Force

                Write-Host
                Write-Host
                Write-Host ("'" + ($zip_Source_FolderPath | Join-Path -ChildPath $zip_leafpath | Convert-Path) + "' is copied to '$zip_Destination_FolderPath'.")


                # Update $zip_filepath
                $zip_filepath = ($zip_Destination_FolderPath | Join-Path -ChildPath (Split-Path -Path $zip_leafpath -Leaf))

                Write-Host
                Write-Host '$zip_filepath =' $zip_filepath
                Write-Host '$zip_leafpath =' $zip_leafpath


                if ($j -le 2)
                {
                    Write-Host
                    Test-Command (MESSAGE New-ZipFile, 'Additional Test', $j, (++$k)) {
                        Write-Host
                        New-ZipFile $zip_filepath -Force | Write-Host -ForegroundColor Yellow
                    }


                    Write-Host
                    Test-Command (MESSAGE New-ZipFile, 'Additional Test', $j, (++$k)) {
                        Write-Host
                        return ((New-ZipFile $zip_filepath) -eq [string]::Empty )
                    }
                }
                else
                {
                    Write-Host
                    Test-Command (MESSAGE New-ZipFile, 'Additional Test', $j, (++$k)) {
                        Write-Host
                        try
                        {
                            New-ZipFile -InputObject $zip_filepath -Path $zip_Destination_FolderPath | Write-Host -ForegroundColor Yellow
                            return $false
                        }
                        catch
                        {
                            # Write-Warning $Error[0].Exception
                            Write-Warning $_

                            if ($Error[0].CategoryInfo.Reason -eq [System.ArgumentException].Name) { return $true }
                            else { return $false }
                        }
                    }


                    Write-Host
                    Test-Command (MESSAGE New-ZipFile, 'Additional Test', $j, (++$k)) {
                        Write-Host
                        try
                        {
                            New-ZipFile -InputObject $zip_filepath -Path $zip_filepath | Write-Host -ForegroundColor Yellow
                            return $false
                        }
                        catch
                        {
                            # Write-Warning $Error[0].Exception
                            Write-Warning $_

                            if (([System.Management.Automation.ErrorRecord]$_).Exception.InnerException.InnerException -is [System.IO.DirectoryNotFoundException])
                            {
                                return $true
                            }
                            else { return $false }
                        }
                    }
                }
            }
        }
    }
}

#####################################################################################################################################################
# Check Help Contents
if ($test_of_HelpContent)
{
    Write-Host (LINE)
    Write-Host ("Help Content Check")
    Write-Host (PRINT START)


    Out-Break
    Get-Help Get-DateString -Full

    Out-Break
    Get-Help Get-FileVersion -Full

    Out-Break
    Get-Help Get-ProductVersion -Full
}

#####################################################################################################################################################
# Restore $VerbosePreference
if ($default_VerbosePreference -ne $null) { $VerbosePreference = $default_VerbosePreference }

#####################################################################################################################################################
# Exit this test script
Exit-Script
