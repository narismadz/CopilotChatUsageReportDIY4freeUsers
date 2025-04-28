$choices = @("Yes", "No")
$caption = "Choose an option"
$message = "Do you have meeting room displayname prefixes `n (Ex Conf Room Pattaya or Conf Room BKK) to be excluded from the export results? 
`n1. Yes, please type prefix following a wildcard (Ex. Conf Room*)`n2. No"
$choice = $host.ui.PromptForChoice($caption, $message, $choices, 0)

if ($choices[$choice] -eq "Yes") {
    $RoomPrefix = Read-Host "Please enter the room account prefix"
    Write-Host "You entered:  $RoomPrefix"
} else {
    $RoomPrefix = ""
    Write-Host "No room account prefix entered."
}

## Required PS module check
if (Get-Module -ListAvailable -Name Microsoft.Graph) {
    Write-Host "Microsoft Graph module is installed." -ForegroundColor Green
} else {
    Write-Host "Microsoft Graph module is not installed. Installing ..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Force
}

$error.clear()
try {

## Connect to MS Graph and sign in
Connect-MgGraph -Scopes "User.Read.All" -NoWelcome


## Export user as CSV to C:\ exclude Guest
$Users = Get-MgUser -All -Property 'UserPrincipalName','DisplayName','Department','JobTitle','UserType' 

$FilteredUsers = $Users | Where-Object { ($_.UserType -ne 'Guest') -and -not ($_.DisplayName -like $RoomPrefix) `
-and -not ($_.DisplayName -like "Microsoft Service Account")  `
-and -not ($_.DisplayName -like "On-Premises Directory Synchronization Service Account") `
-and -not ($_.DisplayName -like "package_*") `
} | Select-Object `
    @{Name='Username'; Expression= {$_.UserPrincipalName}},
    @{Name='DisplayName'; Expression= {$_.DisplayName}},
    @{Name='UserPrincipalName'; Expression= {$_.UserPrincipalName}},
    @{Name='Department'; Expression= {$_.Department}},
    @{Name='JobTitle'; Expression= {$_.JobTitle}}
$FilteredUsers | Export-Csv -Path "C:\EntraUserDetails.csv" -NoTypeInformation -Encoding UTF8
}

catch { "Error occured" }
if (!$error) { Write-Host "Exported CSV saved in C:\EntraUserDetails.csv." -ForegroundColor Green }






