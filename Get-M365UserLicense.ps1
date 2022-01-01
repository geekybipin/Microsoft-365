<#
.SYNOPSIS
    This script will export user licenses from M365. 
.DESCRIPTION
    This script uses MSOL service. So make sure your PowerShell session is connected to MSOL service. `
    This script will export list of users with their assigned M365 license. You have option to provide list of users from CSV file `
    or use All switch to export license of all users in the tenant or use UserPrincipleName parameter to specify one or more users. `
    If you use CSV file, the file must contain column heading as UserPrincipalName with no space and quotes. 
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
.Example
    PS C:\> Get-M365UserLicenses.ps1 -UserPrincipalName username@domain.com 
    This will show license of one user. 
.Example
    PS C:\> Get-M365UserLicenses.ps1 -UserPrincipalName user1@domain.com, user2@domain.com
    This will show license of more than one user. 
.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    General notes
#>
[CmdletBinding()]
param (
    [Parameter(ValueFromPipeline=$false, Mandatory=$false)][Switch]$All,
    [Parameter(ValueFromPipeline=$false, Mandatory=$false)][String]$FileName,
    [Parameter(ValueFromPipeline=$false, Mandatory=$false)][String]$ExportCsv,  
    [Parameter(ValueFromPipeline=$false, Mandatory=$false)][String[]]$UserPrincipalName
)

#Parameter Validation
if (($FileName) -and ($All) )  {
    Write-Host "Error: You cannot use -FileName and -All parameter at the same time. `nnPlease choose correct parameter name. `nSee script examples using help for more information." -ForegroundColor Red
    break
}
elseif (($FileName) -and ($UserPrincipalName)) {
    Write-Host "Error: You cannot use -FileName and -UserPrincipalName parameters at the same time. Please choose correct parameter name. See script examples using help for more information." -ForegroundColor Red
    break   
}
elseif (($All) -and ($UserPrincipalName) )  {
    Write-Host "Error: You cannot use -All and -UserPrincipalName parameters at the same time. `nPlease choose correct parameter name. `nSee script examples using help for more information." -ForegroundColor Red
    break
}

#function block
function Get-License {
#Declaring variables as Array. 
$FinalReport = @()
$GroupObject = @()
$GetLicense = @()
    #running ForEach statement
    foreach ($User in $Users) {
        $UPN = $User.UserPrincipalName
        $UserAccount = Get-MsolUser -UserPrincipalName $UPN -ErrorAction SilentlyContinue
        if ($UserAccount) {
            $GetLicense = ($UserAccount).Licenses
            for($i = 0; $i -lt $GetLicense.Count; $i++){
                $LicArray += $GetLicense[$i].AccountSkuId + "; "
            }
            if ($LicArray -eq "") {
                Write-Host "$($UPN) is unlicensed:" -ForegroundColor Cyan
                $LicArray = "User Unlicensed"
            }
            else {
                $LicArray = ($LicArray).TrimEnd('; ')
                Write-Host "$($UPN) has following licenses:" -ForegroundColor Green
                Write-Host "$LicArray"
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
        else {
            Write-Host "Error: User $UPN not found." -ForegroundColor Red
        }
    }
    if ($ExportCsv) {
        Write-Host "Exporting to csv..." -ForegroundColor Magenta
        $FinalReport | Export-Csv "$ExportCsv"  -NoTypeInformation
    }
}

<# If statement to check paramenter inputs.     
If FileName is used then user is processed based on list of users available in CSV file. 
If All switch is used then all users will be retrieved. 
If CSV file is used with FileName parameter then users in CSV file will be used. 
#>
if ($FileName) {
    try {
        $Users = Import-Csv $FileName
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
elseif ($All) {
    #Retrieve all users in the tenant
    $Users = Get-MsolUser -All
    #calling function
    Get-License
} 
elseif ($UserPrincipalName) {
    $Users = @()
    foreach ($item in $UserPrincipalName) {
        #Creating hashtable
        $newUsersHash = @{
        'UserPrincipalName' = "$item"
        }
    #Creating custom object
    $UserObj = [PSCustomObject]$newUsersHash
    $Users += $UserObj
    }
    #calling function
    Get-License
}