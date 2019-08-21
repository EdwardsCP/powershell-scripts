#Sample Code.  This is not designed to be run as a script.


#Prep Step 1 - Use Powershell to Generate a 256-Bit Key and store it in a given path
#            - For this step, you need to identify a secure location to store your Key.  It is critical that you limit access to this location using ACLs.
$KeyStoragePath = "c:\YourKeyStorage   OR   \\FileServer01\KeyStorageShare"
$KeyFileName = "Username@YourDomainDotCom.AES.Key"
$CreateKey = New-Object Byte[] 32
[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($CreateKey)
$CreateKey | out-file "$KeyStoragePath\$KeyFileName"

#Prep Step 2 - Capture the password for your User Account as a Secure String, AES Encrypt it using the Key generated in Step 1, and save it to a file.
#            - For this step, you need to identify a secure location to store your encrypted Password.  Although this password is encrypted, you should still limit access to this location using ACLs.
#            - Complete this step immediately after Prep Step 1.  It relies on variables that were defined in Prep Step 1.
#            - After the 4th line below is executed, you will need to type the password for your user account (Username@YourDomainDotCom) at Powershell's Read-Host prompt.
$GetKey = Get-Content "$KeyStoragePath\$KeyFileName"
$CredentialsStoragePath = "C:\YourEncryptedCredentialsStorage   OR   \\FileServer02\CredentialsStorageShare"
$CredentialsFileName = "Username@YourDomainDotCom.securestring"
$PasswordSecureString = Read-Host -AsSecureString
$PasswordSecureString | ConvertFrom-SecureString -key $GetKey | Out-File -FilePath "$CredentialsStoragePath\$CredentialsFileName"


#Use your AES Encrypted password file to authenticate with a Mail Server. Define mail server and user, decrypt the encrypted Credentials file, using the Key File, and load it into PSCredential so it can be passed to Send-MailMessage, compose email, and send.
#Define Mail Server Details
$PSEmailServer = "Mail.YourDomainDotCom"
$SMTPPort = 587
$SMTPUsername = "Username"
#Define Key File Details
$KeyStoragePath = "c:\YourKeyStorage\"
$KeyFileName = "Username@YourDomainDotCom.AES.Key"
$GetKey = Get-Content "$KeyStoragePath\$KeyFileName"
#Define Encrypted Password File Detaisl
$CredentialsStoragePath = "C:\YourEncryptedCredentialsStorage"
$CredentialsFileName = "Username@YourDomainDotCom.securestring"
$EncryptedPasswordFile = "$CredentialsStoragePath\$CredentialsFileName"
#Use the Key to decrypt the password and load it into memory as a SecureString
$SecureStringPassword = Get-Content -Path $EncryptedPasswordFile | ConvertTo-SecureString -Key $GetKey
$EmailCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SMTPUsername,$SecureStringPassword
#Define Email Message Options
$MailTo = "RecipientAddress@DomainDotCom"
$MailFrom = "Username@YourDomainDotCom"
$MailSubject = "Hello world"
$MailBody = "Here's the a test email that was sent using Powershell Send-MailMessage to send an email.  The Powershell Script authenticated with the sending mail server using credentials that were stored in an encrypted file and decrypted on the fly during script execution to be passed as a System.Management.Automation.PSCredential."
#Send Email
Send-MailMessage -From $MailFrom -To $MailTo -Subject $MailSubject -Body $MailBody -Port $SMTPPort -Credential $EmailCredential -UseSsl

#Decrypting a password file to reveal the plaintext password.  (re-using some variables that were used in previous code above)
$SMTPUsername = "Username"
$KeyStoragePath = "c:\YourKeyStorage\"
$KeyFileName = "Username@YourDomainDotCom.AES.Key"
$GetKey = Get-Content "$KeyStoragePath\$KeyFileName"
$CredentialsStoragePath = "C:\YourEncryptedCredentialsStorage"
$CredentialsFileName = "Username@YourDomainDotCom.securestring"
$EncryptedPasswordFile = "$CredentialsStoragePath\$CredentialsFileName"
$SecureStringPassword = Get-Content -Path $EncryptedPasswordFile | ConvertTo-SecureString -Key $GetKey
$EmailCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SMTPUsername,$SecureStringPassword
$EmailCredential.GetNetworkCredential().Password
