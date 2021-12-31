<#
.SYNOPSIS
    This script will export user licenses from M365. 
.DESCRIPTION
    This script will export list of users with their assigned license. You have option to provide list of users from CSV file `
    or use All switch to export license of all users in the tenant. If you use CSV file, the file must contain column heading called 'UserPrincipalName' with no space and quotes. 
.EXAMPLE
    PS C:\> Get-M365UserLicenses.ps1 -All
    This will export one or more license(s) of all users in the Microsoft 365 tenant. 
.Example
    PS C:\> Get-M365UserLicenses.ps1 -FileName Users.csv
    This will export one or more license(s) of users listed in the CSV file. The file must contain column hader named UserPrincipalName with list of UPN of users. 
.Example
    PS C:\> Get-M365UserLicenses.ps1 -All -ExportCsv UserLicenseFile.csv
    This will export list of users and their license of all users in the tenant to CSV file. 
.Example
    PS C:\> Get-M365UserLicenses.ps1 -FileName listofusers.csv -ExportCsv UserLicenseFile.csv
    This will export license of only users present in Listofusers.csv file. Note that the file must have column header 'UserPrincipalName' with no space and quotes. 
.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    General notes
#>
[CmdletBinding()]
param (
    [Parameter(ValueFromPipeline=$true, Mandatory=$false, Position=1)][Switch]$All,
    [Parameter(ValueFromPipeline=$false, Mandatory=$false)][String]$FileName,
    [Parameter(ValueFromPipeline=$false, Mandatory=$false)][String]$ExportCsv  
)
<# If statement to check paramenters.     
If All switch is used then all users will be retrieved. 
If user supply CSV file then users in CSV file will be used. 
#>

#function block
function Get-License {
    #Declaring strings as null
    $LicArray = ""
    $FinalReport = ""
    $GroupObject = ""

    #Declaring variables as Array. 
    $FinalReport = @()
    $GroupObject = @()
    $AllLicenses = @()

    #running for each statement
    foreach ($User in $Users) {
        $UPN = $User.UserPrincipalName
        try {
            $AllLicenses = (Get-MsolUser -UserPrincipalName $UPN).Licenses	
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            Write-Host "Error: $ErrorMessage" -ForegroundColor Red
        }
        for($i = 0; $i -lt $AllLicenses.Count; $i++)
            {
                $LicArray += $AllLicenses[$i].AccountSkuId + "; "
            }
        if ($LicArray -eq "") {
            Write-Host "$($UPN) is unlicensed:" -ForegroundColor Cyan
            $LicArray = "User Unlicensed"
        }
        else {
            Write-Host "$($UPN) has following licenses:"
            Write-Host "$LicArray" -ForegroundColor Green
        }
        #Write-Host "`n"
        $ObjectProperties = [Ordered]@{
                "UserPrincipalName" = $UPN
                "License" = $LicArray
                }
        $GroupObject = New-Object -TypeName PSObject -Property $ObjectProperties
        $FinalReport += $GroupObject
        #declaring variable $licArray as null for next loop. 
        $LicArray = ""
    }   
    if ($ExportCsv) {
        $FinalReport | Export-Csv "$ExportCsv"  -NoTypeInformation
    }
    else {
        $FinalReport
    }
}

if ($filename) {
    try {
        $users = Import-Csv $FileName
    }
    catch {
        $ErrorMessage = $_.Exception.Message
        Write-Host "Error: $ErrorMessage" -ForegroundColor Red
        exit
    }
    $Headers = ($Users | Get-Member -MemberType NoteProperty).Name
    $UPNHeading = "UserPrincipalName"
    if ($UPNHeading -in $Headers) {
        #calling function
        Get-License
    }
    else {
        Write-Host "Error: CSV file doesn't contain header named UserPrincipalName" -ForegroundColor Red
        exit
    }
}
else {
    $Users = Get-MsolUser -All
    #calling function
    Get-License
}