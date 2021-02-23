<#
    .SYNOPSIS
        Report about all computers found in AD and their software, excluding 
        servers.

    .DESCRIPTION
        The AD is queried  to retrieve all computer names and properties. This 
        list is then used to query SCCM to find out the software installed on 
        these machines.

    .PARAMETER ImportFile
        Contains all the organization units where we need to search and the 
        e-mail addresses of whom to inform.

    .PARAMETER LogFolder
        Location for the log files.
#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$LogFolder = $env:POWERSHELL_LOG_FOLDER,
    [String]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    Try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start

        #region Import file and vars
        $File = Get-Content $ImportFile -EA Stop | Remove-CommentsHC

        if (-not ($MailTo = $File | Get-ValueFromArrayHC MailTo -Delimiter ',')) {
            throw "No 'MailTo' found in the input file."
        }

        if (-not ($OUs = $File | Get-ValueFromArrayHC -Exclude MailTo)) {
            throw "No organizational units found in the input file."
        }
        #endregion

        #region Logging
        $LogParams = @{
            LogFolder    = New-FolderHC -Path $LogFolder -ChildPath "SCCM Reports\SCCM Installed software\$ScriptName"
            Name         = $ScriptName
            Date         = 'ScriptStartTime'
            NoFormatting = $true
        }
        $LogFile = New-LogFileNameHC @LogParams

        $MachineFolder = New-Item -Path "$LogFile - Machines" -ItemType Directory
        #endregion

        $MailParams = @{
            To          = $MailTo
            Bcc         = $ScriptAdmin
            LogFolder   = $LogParams.LogFolder
            Header      = $ScriptName
            Save        = $LogFile + ' - Mail.html'
            Attachments = @()
        }
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_  -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        $Computers = Get-ADComputerHC -OU $OUs -EA Stop | Where-Object { ($_.Enabled) -and ($_.OS -notlike '*server*') }
        Write-EventLog @EventVerboseParams -Message "$(@($Computers).count) enabled computers found in AD with a non server OS"

        if ($Computers) {
            #region SCCM Primary device user
            $SCCMDeviceUsers = Get-SCCMPrimaryDeviceUsersHC -EA Stop |
            Where-Object { $Computers.Name -contains $_.ComputerName } | Group-Object ComputerName
            Write-EventLog @EventVerboseParams -Message "$(@($SCCMDeviceUsers).count) devices found in SCCM with primary users"
            #endregion

            if ($SCCMsoftware = Get-SCCMInstalledSoftwareHC -ComputerName $Computers.Name -EA Stop |
                Group-Object ComputerName) {
                Write-EventLog @EventVerboseParams -Message "$(@($SCCMsoftware).count) devices found in SCCM with installed software"

                #region Export SCCM installed software for each machine
                Write-EventLog @EventVerboseParams -Message "Export SCCM installed software for each machine to an Excel file in the 'Machines' folder"

                $MachineFiles = @{}

                ForEach ($S in $SCCMsoftware) {
                    Write-Verbose "Export SCCM software to Excel file for '$($S.Name)'"

                    $ExcelParams = @{
                        Path               = Join-Path $MachineFolder "$($S.Name).xlsx"
                        FreezeTopRow       = $true
                        AutoSize           = $true
                        TableName          = 'Software'
                        WorksheetName      = 'SCCM'
                        NoNumberConversion = 'ProductVersion'
                    }
                    $S.Group | Select-Object * -ExcludeProperty 'ComputerName' |
                    Sort-Object 'ProductName' | Export-Excel @ExcelParams

                    $MachineFiles.($S.Name) = $ExcelParams.Path
                }
                #endregion

                #region Export SCCM ProductName of all machines to one sheet
                $ProductNamesAll = $SCCMsoftware.Group | Group-Object ProductName | Sort-Object Name |
                Select-Object @{N = 'ProductName'; E = { $_.Name } },
                @{N = 'ComputerCount'; E = { $_.Count } },
                @{N = 'ComputerName'; E = { $_.Group.ComputerName -join ', ' } }

                Write-EventLog @EventVerboseParams -Message "Export $($ProductNamesAll.Count) product names to Excel, this is all the software found on all the devices in one Excel file"

                $ExcelParams = @{
                    Path               = "$LogFile - SCCM installed software.xlsx"
                    FreezeTopRow       = $true
                    AutoSize           = $true
                    TableName          = 'ProductName'
                    WorksheetName      = 'ProductName'
                    NoNumberConversion = 'ProductVersion'
                }
                $ProductNamesAll | Export-Excel @ExcelParams
                #endregion

                $MailParams.Attachments += $ExcelParams.Path

                #region Export SCCM ProductName and version of all machines to one sheet
                $ProductVersionsAll = $SCCMsoftware.Group | Group-Object ProductName, ProductVersion |
                Sort-Object Name | Select-Object @{N = 'ProductName, ProductVersion'; E = { $_.Name } },
                @{N = 'ComputerCount'; E = { $_.Count } },
                @{N = 'ComputerName'; E = { $_.Group.ComputerName -join ', ' } }

                Write-EventLog @EventVerboseParams -Message "Export $($ProductVersionsAll.Count) product name and version to Excel, this is all the software found on all the devices in one Excel file"

                $ExcelParams = @{
                    Path               = "$LogFile - SCCM installed software.xlsx"
                    FreezeTopRow       = $true
                    AutoSize           = $true
                    TableName          = 'ProductVersion'
                    WorksheetName      = 'ProductVersion'
                    NoNumberConversion = 'ProductVersion'
                }
                $ProductVersionsAll | Export-Excel @ExcelParams
                #endregion
            }

            #region SCCM and DNS IP address
            Write-EventLog @EventVerboseParams -Message "Retrieve SCCM and DNS IP address details for $(@($Computers.Name).Count) computers"

            $DNSandSCCMdetails = Get-SCCMandDNSdetailsHC -ComputerName $Computers.Name

            Write-EventLog @EventVerboseParams -Message "Retrieved $($DNSandSCCMdetails.Count) SCCM and DNS IP address details"
            #endregion

            #region Hardware details: Manufacturer, model, type, ..
            $SCCMHardware = Get-SCCMHardwareHC
            #endregion

            #region Create objects for the AD computers Excel sheet
            $ADComputers = $Computers | Select-Object @{N = 'Name'; E = {
                    $Script:ComputerName = $_.Name

                    $Script:SCCMHardwareObj = $SCCMHardware.where(
                        { $_.ComputerName -eq $ComputerName }, 'First', 1)
                    $Script:Users = $SCCMDeviceUsers.where(
                        { $_.Name -eq $ComputerName }, 'First', 1)
                    $Script:SCCMandDNS = $DNSandSCCMdetails.where(
                        { $_.ComputerName -eq $ComputerName }, 'First', 1)

                    $_.Name
                }
            },
            @{N = 'DeviceType'; E = {
                    $SCCMHardwareObj.DeviceType
                }
            },
            @{N = 'SoftwareCount'; E = {
                    $($SCCMsoftware.where( { $_.Name -eq $ComputerName }, 'First', 1)).Count
                }
            },
            @{N = 'Manufacturer'; E = {
                    $SCCMHardwareObj.Manufacturer
                }
            },
            @{N = 'Model'; E = {
                    $SCCMHardwareObj.Model
                }
            },
            @{N = 'PrimaryUser'; E = {
                    $Users.Group.SamAccountName -join ", `r`n"
                }
            },
            @{N = 'DisplayName'; E = {
                    $Users.Group.DisplayName -join ", `r`n"
                }
            },
            @{N = 'SCCM IP'; E = {
                    $SCCMandDNS.SCCM -join ", `r`n"
                }
            },
            @{N = 'DNS IP'; E = {
                    $SCCMandDNS.DNS -join ", `r`n"
                }
            },
            @{N = 'Subnet'; E = {
                    $SCCMandDNS.Subnet
                }
            },
            @{N = 'Location'; E = {
                    $SCCMandDNS.Location
                }
            },
            @{N = 'AD Description'; E = {
                    $_.Description
                }
            },
            @{N = 'AD OS'; E = {
                    $_.OS
                }
            },
            @{N = 'OS'; E = {
                    $SCCMHardwareObj.OperatingSystem
                }
            },
            @{N = 'OS Version'; E = {
                    $SCCMHardwareObj.OperatingSystemVersion
                }
            },
            @{N = 'AD Created'; E = {
                    $_.Created
                }
            },
            @{N = 'AD Last logon'; E = {
                    $_.'Last logon'
                }
            },
            @{N = 'AD OU'; E = {
                    $_.OU
                }
            },
            @{N = 'ChassisTypes'; E = {
                    $SCCMHardwareObj.ChassisTypes
                }
            }

            Write-EventLog @EventVerboseParams -Message "Export all AD Computers $($ADComputers.Count) to one Excel file"
            #endregion

            #region Export AD Computers to Excel sheet
            $ExcelParams = @{
                Path               = "$LogFile - SCCM AD computers overview.xlsx"
                AutoSize           = $true
                BoldTopRow         = $true
                FreezeTopRow       = $true
                WorkSheetname      = 'AD Computers'
                TableName          = 'Computers'
                NoNumberConversion = 'SCCM IP', 'DNS IP', 'Subnet'
                ErrorAction        = 'Stop'
            }
            $Excel = $ADComputers | Sort-Object Name | 
            Export-Excel @ExcelParams -PassThru

            $sheet = $Excel.Workbook.Worksheets | Select-Object -First 1

            $ComputerNameColumn = 1
            $CountColumn = 3

            foreach (
                $row in 
                (($sheet.Dimension.Start.Row + 1) .. $sheet.Dimension.End.Row)
            ) {
                if (
                    $Link = $MachineFiles.(
                        $sheet.Cells[$row, $ComputerNameColumn].Value)
                ) {
                    $Value = $sheet.Cells[$row, $CountColumn].Value
                    $sheet.cells[$row, $CountColumn].Hyperlink = $Link
                    $sheet.cells[$row, $CountColumn].Value = $Value
                    $sheet.cells[$row, $CountColumn] | 
                    Set-ExcelRange -Underline -FontColor Blue
                }
            }

            $sheet.Column($ComputerNameColumn) | 
            Set-ExcelRange -HorizontalAlignment Center
            $sheet.Column($CountColumn) | 
            Set-ExcelRange -HorizontalAlignment Center
            $sheet.Column(2) | Set-ExcelRange -HorizontalAlignment Center
            $sheet.Column(4) | Set-ExcelRange -HorizontalAlignment Center
            $sheet.Column(5) | Set-ExcelRange -HorizontalAlignment Center

            Close-ExcelPackage $Excel
            #endregion

            #region Export SCCM Primary device user
            if (
                $MultipleMachineUsers = $SCCMDeviceUsers.Group | 
                Group-Object SamAccountName | Where-Object { $_.Count -ge 2 }
            ) {
                $ExcelParams = @{
                    Path          = "$LogFile - SCCM AD computers overview.xlsx"
                    FreezeTopRow  = $true
                    AutoSize      = $true
                    TableName     = 'MultiMachineUser'
                    WorksheetName = 'MultiMachineUser'
                }
                $MultipleMachineUsers.Group |
                Sort-Object DisplayName, ComputerName |
                Select-Object DisplayName, SamAccountName, ComputerName | 
                Export-Excel @ExcelParams
            }
            #endregion

            $MailParams.Attachments += $ExcelParams.Path
        }

        $MailParams.Subject = "$(@($ProductNamesAll).count) software packages"

        $MailParams.Message = "<p>Within SCCM we found <b>$(@($ProductNamesAll).count) unique software packages</b> installed on <b>$(@($Computers).count) enabled AD computers</b>, regardless of the product version.</p>"

        $MailParams.Message += $ADComputers | 
        Group-Object 'DeviceType', 'AD OS' |
        Select-Object @{N = 'AD OS'; E = { $_.Group[0].'AD OS' } },
        @{N = 'DeviceType'; E = { $_.Group[0].DeviceType } },
        @{N = 'Total'; E = { $_.Count } } |
        Sort-Object 'AD OS', 'DeviceType' | 
        ConvertTo-Html -As Table -Fragment

        $MailParams.Message += "<p><i>* Check the attachment for details</i></p>"
        $MailParams.Message += $OUs | ConvertTo-OuNameHC -OU | Sort-Object |
        ConvertTo-HtmlListHC -Header 'Organizational units:'

        Get-ScriptRuntimeHC -Stop
        Remove-EmptyParamsHC $MailParams
        Send-MailHC @MailParams
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_  -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}