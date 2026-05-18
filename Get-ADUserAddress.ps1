param(
    [Parameter(Mandatory)]
    [string]$Identity
)

$searcher = [System.DirectoryServices.DirectorySearcher]::new()
$searcher.Filter = "(samAccountName=$Identity)"
$searcher.PropertiesToLoad.AddRange(@('cn','samAccountName','streetAddress','l','st','postalCode','co','userAccountControl'))

$result = $searcher.FindOne()

if (-not $result) {
    Write-Error "User '$Identity' not found in Active Directory."
    exit 1
}

$p = $result.Properties
$disabled = ($p['userAccountControl'][0] -band 2) -ne 0

[PSCustomObject]@{
    Name       = "$($p['cn'][0])"
    SamAccount = "$($p['samAccountName'][0])"
    Disabled   = $disabled
    Street     = "$($p['streetAddress'][0])"
    City       = "$($p['l'][0])"
    State      = "$($p['st'][0])"
    PostalCode = "$($p['postalCode'][0])"
    Country    = "$($p['co'][0])"
}
