# email-expiring-users

Python project built to take a spread sheet of users and email users with domain password expiring in 15 days.

Powershell was used to export users, email and expiring date.

openpyxl library used to read spread sheet.
smtplib library used to send emails.

Command to extract AD users and expiry date from via Powershell

** Get-ADUser -Filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} -Properties DisplayName, EmailAddress, "msDS-UserPasswordExpiryTimeComputed" | Select-Object -Property "Displayname","EmailAddress",@{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}} | Sort-Object EXPIRYDATE  | Export-CSV C:\Script **
