# Script Framework
# Name    : StringEncryption.psm1
# Version : 0.1
# Date    : 2017-02-19
# Author  : Roeland van den Bosch
# Website : http://www.roelandvdbosch.nl

[string]$strSaltPassword = "D1t 1s 33n w1llek3ur1g w@chtw00rd"
[string]$strInitPassword = "N0g 33n w@chtw00rd!"

Function New-EncryptedString
{
<#
  .Synopsis
    Create an encrypted string.
  .Description
    The New-EncryptedString cmdlet encrypts a string.
  .Example
    New-EncryptedString -InputString String -Passphrase Passphrase -SaltPassword Salt -InitPassword InitPassword -Log @($true,$objLogValue)
    Returns the encrypted value of the input
    string and log messages
  .Example
    New-EncryptedString -InputString String -Passphrase Passphrase -SaltPassword Salt -InitPassword InitPassword
    Returns the encrypted value of the input
    string
  .Parameter InputString
    String.
  .Parameter Passphrase
    Custom passphrase.
  .Parameter SaltPassword
    First passphrase.
  .Parameter InitPassword
    Second passphrase.
  .Inputs
    [string]
    [string]
    [string]
    [string]
    [array]
  .Output
    [string]
  .Notes
    NAME: New-EncryptedString
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170219
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param (
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Input String")]
    [ValidateNotNullOrEmpty()]
    [string]$InputString,
    [Parameter(Mandatory=$true, Position=1, HelpMessage="Passphrase")]
    [ValidateNotNullOrEmpty()]
    [string]$Passphrase,
    [Parameter(Mandatory=$false, Position=2, HelpMessage="Salt")]
    [ValidateNotNullOrEmpty()]
    [string]$SaltPassword = $strSaltPassword,
    [Parameter(Mandatory=$false, Position=3, HelpMessage="Init Password")]
    [ValidateNotNullOrEmpty()]
    [string]$InitPassword = $strInitPassword,
    [Parameter(Mandatory=$false, Position=4, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t[New-EncryptedString]"}

  [object]$objRijndaelManaged = new-Object System.Security.Cryptography.RijndaelManaged
  [array]$arrPassphrase = [Text.Encoding]::UTF8.GetBytes($Passphrase)
  [array]$arrSalt = [Text.Encoding]::UTF8.GetBytes($SaltPassword)

  $objRijndaelManaged.Key = (new-Object Security.Cryptography.PasswordDeriveBytes $arrPassphrase, $arrSalt, "SHA256", 5).GetBytes(32) #256/8
  $objRijndaelManaged.IV = (new-Object Security.Cryptography.SHA256Managed).ComputeHash( [Text.Encoding]::UTF8.GetBytes($InitPassword) )[0..15]
   
  [object]$objEncryptor = $objRijndaelManaged.CreateEncryptor()
  [object]$objMemoryStream = new-Object IO.MemoryStream
  [object]$objCryptoStream = new-Object Security.Cryptography.CryptoStream $objMemoryStream,$objEncryptor,"Write"
  [object]$objStreamWriter = new-Object IO.StreamWriter $objCryptoStream
  $objStreamWriter.Write($InputString)
  $objStreamWriter.Close()
  $objCryptoStream.Close()
  $objMemoryStream.Close()
  $objRijndaelManaged.Clear()
  [array]$arrResult = $objMemoryStream.ToArray()

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t[New-EncryptedString]"}

  Return [Convert]::ToBase64String($arrResult)
}

Function New-DecryptedString
{
<#
  .Synopsis
    Create a decrypted string.
  .Description
    The New-DecryptedString cmdlet decrypts an 
    encrypted string.
  .Example
    New-DecryptedString -InputString EncrytedString -Passphrase Passphrase -SaltPassword Salt -InitPassword InitPassword -Log @($true,$objLogValue)
    Returns the decrypted value of the input
    encrypted string and log messages
  .Example
    New-DecryptedString -InputString EncrytedString -Passphrase Passphrase -SaltPassword Salt -InitPassword InitPassword -Log @($true,$objLogValue)
    Returns the decrypted value of the 
    input encrypted string
  .Parameter InputString
    Encrypted string.
  .Parameter Passphrase
    Custom passphrase.
  .Parameter SaltPassword
    First passphrase.
  .Parameter InitPassword
    Second passphrase.
  .Parameter Log
    Array of constant log values.
  .Inputs
    [string]
    [string]
    [string]
    [string]
    [array]
  .Output
    [string]
  .Notes
    NAME: New-DecryptedString
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170117
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param (
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Input String")]
    [ValidateNotNullOrEmpty()]
    [string]$InputString,
    [Parameter(Mandatory=$true, Position=1, HelpMessage="Passphrase")]
    [ValidateNotNullOrEmpty()]
    [string]$Passphrase,
    [Parameter(Mandatory=$false, Position=2, HelpMessage="Salt")]
    [ValidateNotNullOrEmpty()]
    [string]$SaltPassword = $strSaltPassword,
    [Parameter(Mandatory=$false, Position=3, HelpMessage="Init Password")]
    [ValidateNotNullOrEmpty()]
    [string]$InitPassword = $strInitPassword,
    [Parameter(Mandatory=$false, Position=4, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t[New-DecryptedString]"}

  [array]$arrInput = [Convert]::FromBase64String($InputString)
  [object]$objRijndaelManaged = new-Object System.Security.Cryptography.RijndaelManaged
  [array]$arrPassphrase = [System.Text.Encoding]::UTF8.GetBytes($Passphrase)
  [array]$arrSalt = [System.Text.Encoding]::UTF8.GetBytes($SaltPassword)

  $objRijndaelManaged.Key = (new-Object Security.Cryptography.PasswordDeriveBytes $arrPassphrase, $arrSalt, "SHA256", 5).GetBytes(32) #256/8
  $objRijndaelManaged.IV = (new-Object Security.Cryptography.SHA256Managed).ComputeHash( [Text.Encoding]::UTF8.GetBytes($InitPassword) )[0..15]

  [object]$objDecryptor = $objRijndaelManaged.CreateDecryptor()
  [object]$objMemoryStream = new-Object IO.MemoryStream @(,$arrInput)
  [object]$objCryptoStream = new-Object Security.Cryptography.CryptoStream $objMemoryStream,$objDecryptor,"Read"
  [object]$objStreamWriter = new-Object IO.StreamReader $objCryptoStream
  [string]$strDecryptedString = $objStreamWriter.ReadToEnd()
  $objStreamWriter.Close()
  $objCryptoStream.Close()
  $objMemoryStream.Close()
  $objRijndaelManaged.Clear()

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t[New-DecryptedString]"}

  Return $strDecryptedString
}
