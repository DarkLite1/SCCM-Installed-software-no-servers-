#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $MailAdminParams = {
        ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and ($Subject -eq 'FAILURE')
    }

    $testImportFile = @"
MailTo: Brecht.Gijbels@heidelbergcement.com
OU=XXX,OU=EU,DC=contoso,DC=net
"@

    $testADComputers = @(
        [PSCustomObject]@{
            Name    = 'PC1'
            Enabled = $true
        }
        [PSCustomObject]@{
            Name    = 'PC2'
            Enabled = $true
        }
    )

    $testInstalledSoftware = @(
        [PSCustomObject]@{
            ComputerName   = 'PC1'
            ProductName    = 'Office'
            ProductVersion = 1
        }
        [PSCustomObject]@{
            ComputerName   = 'PC1'
            ProductName    = 'McAffee'
            ProductVersion = 2
        }
        [PSCustomObject]@{
            ComputerName   = 'PC2'
            ProductName    = 'Office'
            ProductVersion = 1
        }
        [PSCustomObject]@{
            ComputerName   = 'PC2'
            ProductName    = 'McAffee'
            ProductVersion = 2
        }
    )

    $SCCMPrimaryDeviceUsersHC = @(
        [PSCustomObject]@{
            ComputerName   = 'PC1'
            SamAccountName = 'Bob'
            DisplayName    = 'Bob Lee swagger'
        }
        [PSCustomObject]@{
            ComputerName   = 'PC1'
            SamAccountName = 'Mike'
            DisplayName    = 'Mike and the mechanics'
        }
        [PSCustomObject]@{
            ComputerName   = 'PC2'
            SamAccountName = 'Jake'
            DisplayName    = 'Jake Sully'
        }
    )

    $testOutParams = @{
        FilePath = (New-Item 'TestDrive:/Test.txt' -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')

    $testParams = @{
        ScriptName = 'Test (Brecht)'
        ImportFile = $testOutParams.FilePath
        LogFolder  = New-Item 'TestDrive:/log' -ItemType Directory
    }

    Mock Get-ADComputerHC
    Mock Get-SCCMHardwareHC
    Mock Get-SCCMPrimaryDeviceUsersHC
    Mock Get-SCCMandDNSdetailsHC
    Mock Send-MailHC
    Mock Write-EventLog
}

Describe 'Prerequisites' {
    Context 'ImportFile' {
        It 'skip comments' {
            @"
MailTo: Brecht.Gijbels@heidelbergcement.com
# comment
# comment
OU=XXX,OU=EU,DC=contoso,DC=net
# Comment
"@ | Out-File @testOutParams
            
            .$testScript @testParams

            $Expected = @(
                'MailTo: Brecht.Gijbels@heidelbergcement.com'
                'OU=XXX,OU=EU,DC=contoso,DC=net'
            )

            Assert-Equivalent -Actual $File -Expected $Expected
        } 
        It 'mandatory parameter' {
            (Get-Command $testScript).Parameters['ImportFile'].Attributes.Mandatory | Should -Be $true
        } 
        It 'file not found' {
            .$testScript -ScriptName $testParams.ScriptName -LogFolder $testParams.LogFolder -ImportFile 'NotExisting.txt'

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "Cannot find path*")
            }
        } 
        It 'OU missing' {
            @"
MailTo: Brecht.Gijbels@heidelbergcement.com
"@  | Out-File @testOutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and
                ($Message -like "*No organizational units found*")
            }
        } 
        It 'MailTo missing' {
            @"
OU=XXX,OU=EU,DC=contoso,DC=net
"@ | Out-File @testOutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*No 'MailTo' found*")
            }
        } 
    }
    Context 'LogFolder' {
        It 'folder not found' {
            $testImportFile | Out-File @testOutParams

            .$testScript -ScriptName $testParams.ScriptName -LogFolder 'NonExisting' -ImportFile $testParams.ImportFile

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*Path*not found")
            }
        } 
    }
}
Describe 'export Excel files' {
    BeforeAll {
        $testImportFile | Out-File @testOutParams
    }
    BeforeEach {
        Remove-Item "$($testParams.LogFolder)\*" -Recurse -Force
    }
    It 'one file for each computer with its software in the machines folder' {
        Mock Get-ADComputerHC {
            $testADComputers
        }
        Mock Get-SCCMInstalledSoftwareHC {
            $testInstalledSoftware
        }

        .$testScript @testParams

        $testMachines = @($testInstalledSoftware.ComputerName | 
            Sort-Object -Unique).Count

        $testMachines | Should -Not -BeExactly 0

        Get-ChildItem $testParams.LogFolder -Recurse -Directory |
        Where-Object { $_.Name -like '*Machines' } | Get-ChildItem -File |
        Where-Object { 
            $testInstalledSoftware.ComputerName -contains $_.BaseName } |
        Should -HaveCount $testMachines
    } 
    It 'one overview file for all SCCM software installed' {
        Mock Get-ADComputerHC {
            $testADComputers
        }
        Mock Get-SCCMInstalledSoftwareHC {
            $testInstalledSoftware
        }

        .$testScript @testParams

        @($testInstalledSoftware.ProductName).Count | Should -Not -BeExactly 0

        Get-ChildItem $testParams.LogFolder -Recurse |
        Where-Object { $_.Name -like "*SCCM installed software.xlsx" } | Should -HaveCount 1
    } 
    It 'one file for all AD computers' {
        Mock Get-ADComputerHC {
            $testADComputers
        }
        Mock Get-SCCMInstalledSoftwareHC {
            $testInstalledSoftware
        }

        .$testScript @testParams

        @($testInstalledSoftware.ProductName).Count | Should -Not -BeExactly 0

        Get-ChildItem $testParams.LogFolder -Recurse |
        Where-Object { $_.Name -like "*SCCM AD computers overview.xlsx" } | Should -HaveCount 1
    } 
}
Describe 'send mail' {
    BeforeAll {
        $testImportFile | Out-File @testOutParams
    }
    BeforeEach {
        Remove-Item "$($testParams.LogFolder)\*" -Recurse -Force
    }
    It "with the 'AD Computers' and 'All installed software' in attachment" {
        Mock Get-ADComputerHC {
            $testADComputers
        }
        Mock Get-SCCMInstalledSoftwareHC {
            $testInstalledSoftware
        }

        Mock Get-SCCMPrimaryDeviceUsersHC {
            $SCCMPrimaryDeviceUsersHC
        }

        .$testScript @testParams

        @($testInstalledSoftware.ProductName).Count | Should -Not -BeExactly 0

        Should -Invoke Send-MailHC -Times 1 -Exactly -ParameterFilter {
            (@($Attachments).Count -eq 2)
        }
    } 
}

<# 
Invoke-Pester 'T:\Prod\SCCM Reports\SCCM Installed software (No servers)\SCCM Installed software (No servers).Tests.ps1' -Output Detailed 
#>