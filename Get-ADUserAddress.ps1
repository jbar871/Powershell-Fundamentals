# Requires the ActiveDirectory module (RSAT or AD DS role)
# Run on a domain-joined machine or provide -Server to target a DC

param(
    [Parameter(Mandatory)]
    [string]$Identity   # samAccountName, UPN, DN, or display name
)

$properties = @(
    'StreetAddress',
    'City',
    'State',
    'PostalCode',
    'Country',
    'co'          # country display name (friendly text)
)

try {
    $user = Get-ADUser -Identity $Identity -Properties $properties -ErrorAction Stop
} catch {
    Write-Error "User '$Identity' not found or AD module unavailable: $_"
    exit 1
}

[PSCustomObject]@{
    Name        = $user.Name
    SamAccount  = $user.SamAccountName
    Street      = $user.StreetAddress
    City        = $user.City
    State       = $user.State
    PostalCode  = $user.PostalCode
    Country     = $user.co          # human-readable country name
}
