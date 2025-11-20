#========================================================================================#
# INPUTS (replace the values below)
#========================================================================================#
$UserName     = "Mohamed Mohamed Ahmed Moahmed Elmesary"
$Pager        = "10150"
$Department   = "DevOps Sector"
$Phone        = "01558901769"
$Title        = "Senior DevOps"
$Company      = "HD Bank"
$ManagerEmail = "ahmed.elhgawy@hdbank.local"
$Path         = "hdbank.local/Departments/IT/DevOps"
$PasswordText = "qwer1010@"
$Email        = "hdbank.local"
#========================================================================================#


#===============================
# Function: Write-Info / Error
#===============================
function Write-Info($Message)  { Write-Host "[INFO]  $Message" -ForegroundColor Cyan }
function Write-Warn($Message)  { Write-Host "[WARN]  $Message" -ForegroundColor Yellow }
function Write-Err($Message)   { Write-Host "[ERROR] $Message" -ForegroundColor Red }


#===============================
# 1) Validate Required Fields
#===============================
if (-not $UserName)     { Write-Err "UserName is required"; exit 1 }
if ($Company -eq "HD Bank" -and -not $Pager) { Write-Err "Pager is required"; exit 1 }
if (-not $ManagerEmail) { Write-Err "ManagerEmail is required"; exit 1 }
if (-not $Path)         { Write-Err "OU Path is required"; exit 1 }


#===============================
# 2) Parse Name (Safe)
#===============================
$Clean = $UserName.Trim()
$Parts = $Clean -split '\s+'

switch ($Parts.Count) {
    1 {
        Write-Err "Full name must contain at least first & last name"
        exit 1
    }
    2 {
        $First = $Parts[0]
        $Last  = $Parts[1]
        $Mid   = ""
    }
    default {
        $First = $Parts[0]
        $Last  = $Parts[-1]
        $Mid   = ($Parts[1..($Parts.Count-2)] -join " ")
    }
}

$SurName = ($Mid + " " +$Last)
$LoginName = ($First + "." + $Last).ToLower()
$UserEmail = ($LoginName + "@" + $Email)
Write-Info "Generated login name: $LoginName"


#===============================
# 3) Expiration Date Policy
#===============================
$ExpirationDate = if ($Company -eq "HD Bank") {
    (Get-Date).AddMonths(12)
}
else {
    (Get-Date).AddMonths(3)
}
Write-Info "User Expiration Date: $ExpirationDate"


#===============================
# 4) Resolve Manager DN
#===============================
$ManagerName = $ManagerEmail -replace "@.*", ""
$Manager = Get-ADUser -Filter "SamAccountName -eq '$ManagerName'"
if (-not $Manager) {
    Write-Err "Manager not found in AD: $ManagerEmail"
    exit 1
}

$ManagerDN = $Manager.DistinguishedName
Write-Info "Manager found: $($Manager.SamAccountName)"


#===============================
# 5) Convert OU Path → DN
#===============================
try {
    $Domain, $OUPath = $Path.Split("/", 2)
    $DC = $Domain.Split(".") | ForEach-Object { "DC=$_" }
    $OU = $OUPath.Split("/")  | ForEach-Object { "OU=$_" }
    $Count = $OU.Count
    $OU = $OU[($Count-1)..0] 
    $DN = ($OU + $DC) -join ","
}
catch {
    Write-Err "Failed to parse OU path: $Path"
    exit 1
}

Write-Info "Target OU DN: $DN"


#===============================
# 6) Check for Conflicts
#===============================
if ($Company -eq "HD Bank" -and (Get-ADUser -Filter "pager -eq '$Pager'" -Properties pager)) {
    Write-Err "A user with pager=$Pager already exists."
    exit 1
}

if (Get-ADUser -Filter "SamAccountName -eq '$LoginName'") {
    Write-Err "A user with login name '$LoginName' already exists."
    exit 1
}


#===============================
# 7) Convert Password Securely
#===============================
$Password = ConvertTo-SecureString $PasswordText -AsPlainText -Force


#===============================
# 8) Create User
#===============================
try {
    New-ADUser `
        -Name $UserName `
        -GivenName $First `
        -Surname $SurName `
        -DisplayName $UserName `
        -SamAccountName $LoginName `
        -UserPrincipalName $UserEmail `
        -Path $DN `
        -Department $Department `
        -MobilePhone $Phone `
        -Company $Company `
        -Title $Title `
        -Manager $ManagerDN `
        -AccountPassword $Password `
        -Enabled $true `
        -ChangePasswordAtLogon $true `
        -AccountExpirationDate $ExpirationDate `
        -OtherAttributes @{
            pager = $Pager
        }

    Write-Host ""
    Write-Host "==========================================" -ForegroundColor Green
    Write-Host " USER CREATED SUCCESSFULLY " -ForegroundColor Green
    Write-Host "------------------------------------------"
    Write-Host " Login Name : HDBank\$LoginName"
    Write-Host " Password   : $PasswordText"
    Write-Host " Expires    : $($ExpirationDate.ToShortDateString())"
    Write-Host "==========================================" -ForegroundColor Green
}
catch {
    Write-Err "Failed to create user: $($_.Exception.Message)"
    exit 1
}
