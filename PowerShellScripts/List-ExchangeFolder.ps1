param (
    $MailAddress,
    $OutputFolder
)

function Create-OutputFolder($OutputFolder) {
    if (-not (Test-Path -Path $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder
    }
}

Create-OutputFolder -OutputFolder $OutputFolder
Connect-ExchangeOnline -UserPrincipalName $MailAddress
$x = Get-MailboxFolder -Recurse
$json = $x | ConvertTo-Json
$json | Out-File "${OutputFolder}\${UserPrincipalName}.json"
