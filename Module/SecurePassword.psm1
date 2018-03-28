# Script Framework
# Name    : SecurePassword.psm1
# Version : 0.1
# Date    : 2018-03-28
# Author  : Roeland van den Bosch
# Website : http://www.roelandvdbosch.nl

Function New-EncryptedPassPhrase
{
<#
  .Synopsis
    Create an encrypted passphrase.
  .Description
    The New-EncryptedPassPhrase cmdlet encrypts a string.
  .Example
    New-EncryptedPassPhrase -InputString String -Passphrase Passphrase -Log @($true,$objLogValue)
    Returns the encrypted value of the input
    string and log messages
  .Example
    New-EncryptedPassPhrase -InputString String -Passphrase Passphrase
    Returns the encrypted value of the input
    string
  .Parameter InputString
    String.
  .Parameter Passphrase
    Optional passphrase.
  .Inputs
    [string]
    [string]
    [array]
  .Output
    [string]
  .Notes
    NAME: New-EncryptedPassPhrase
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20180328
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param (
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Input String")]
    [ValidateNotNullOrEmpty()]
    [string]$InputString,
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Passphrase")]
    [ValidateNotNullOrEmpty()]
    [string]$Passphrases,
    [Parameter(Mandatory=$false, Position=2, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )
  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t[New-EncryptedPassPhrase]"}
  Try {[string]$strIncludeDirectory = Split-Path $script:MyInvocation.MyCommand.Path} Catch {}
  [void][system.reflection.Assembly]::UnsafeLoadFrom("$strIncludeDirectory\SecurePassword\SecurePassword.dll")
  If ([string]::IsNullOrEmpty($Passphrases)) {[string]$strReturnValue = [SecurePassword.Encode_Decode]::EncryptString($InputString)}
  Else {[string]$strReturnValue = [SecurePassword.Encode_Decode]::EncryptString($InputString, $Passphrases)}
  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t[New-EncryptedPassPhrase]"}
  Return $strReturnValue
}

Function New-DecryptedPassPhrase
{
<#
  .Synopsis
    Create a decrypted passphrase.
  .Description
    The New-DecryptedPassPhrase cmdlet decrypts an 
    encrypted string.
  .Example
    New-DecryptedPassPhrase -InputString EncrytedString -Passphrase Passphrase -Log @($true,$objLogValue)
    Returns the decrypted value of the input
    encrypted string and log messages
  .Example
    New-DecryptedPassPhrase -InputString EncrytedString -Passphrase Passphrase -Log @($true,$objLogValue)
    Returns the decrypted value of the 
    input encrypted string
  .Parameter InputString
    Encrypted string.
  .Parameter Passphrase
    Optional passphrase.
  .Parameter Log
    Array of constant log values.
  .Inputs
    [string]
    [string]
    [array]
  .Output
    [string]
  .Notes
    NAME: New-DecryptedPassPhrase
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20180328
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param (
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Input String")]
    [ValidateNotNullOrEmpty()]
    [string]$InputString,
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Passphrase")]
    [ValidateNotNullOrEmpty()]
    [string]$Passphrase,
    [Parameter(Mandatory=$false, Position=2, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )
  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t[New-DecryptedPassPhrase]"}
  Try {[string]$strIncludeDirectory = Split-Path $script:MyInvocation.MyCommand.Path} Catch {}
  [void][system.reflection.Assembly]::UnsafeLoadFrom("$strIncludeDirectory\SecurePassword\SecurePassword.dll")
  If ([string]::IsNullOrEmpty($Passphrases)) {[string]$strReturnValue = [SecurePassword.Encode_Decode]::DecryptString($InputString)}
  Else {[string]$strReturnValue = [SecurePassword.Encode_Decode]::DecryptString($InputString, $Passphrases)}
  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t[New-DecryptedPassPhrase]"}
  Return $strReturnValue
}
