<#
=============================================================================================
Name:           Userwise Shared Mailbox Information
Description:    This script exports Office 365 user Shared Mailbox Access to CSV
Version:        1.0
Website:        gwltd.co.nz

============================================================================================
#>




Connect-ExchangeOnline
Write-Host "Enter The UPN" -ForegroundColor Green
$copyfrom = Read-Host "Enter the user id, for example 'ricoh.test@shorehire.com.au': "

# Fetch shared mailbox permissions for the user
$Result = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Get-MailboxPermission -User $copyfrom

if ($Result.Count -eq 0) {
    Write-Host "$copyfrom does not have access to any shared mailbox" -ForegroundColor Yellow
} else {
    # Export the result to CSV and display it
    $Result | Export-Csv "C:\TEMP\Sharedmailbox.csv" -NoTypeInformation
    $Result | Out-GridView
    Write-Host "$copyfrom has access to shared mailboxes. You can find the details in the CSV file." -ForegroundColor Green
}

# Signature
Write-Host "`n~~ Script prepared by Ausaf Ahmad ~~`n" -ForegroundColor Green
