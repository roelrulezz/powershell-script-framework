# Script Framework
# Name    : Log.psm1
# Version : 0.1
# Date    : 2017-02-19
# Author  : Roeland van den Bosch
# Website : http://www.roelandvdbosch.nl

Function New-ExchangeServiceObject
{
<#
  .Synopsis
    Create an Exchange service object.
  .Description
    The New-ExchangeServiceObject cmdlet creates an Exchange service object.
  .Example
    New-ExchangeServiceObject -ExchangeUsername emailaddress -ExchangePassword password -Log @($true,$objLogValue)
    Returns True or False, an Exchange Service datababase object and log messages
  .Example
    New-ExchangeServiceObject -ExchangeUsername emailaddress -ExchangePassword password
    Returns True or False and an Exchange Service datababase object
  .Parameter ExchangeUsername
    Exchange username / email adres.
  .Parameter ExchangePassword
    Exchange passphrase.
  .Parameter Log
    Array of constant log values.
  .Inputs
    [string]
    [string]
    [array]
  .Output
    [array]
  .Notes
    NAME: New-ExchangeServiceObject
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170303
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
  #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Exchange Username")]
    [ValidateNotNullOrEmpty()]
    [string]$ExchangeUsername,
    [Parameter(Mandatory=$true, Position=1, HelpMessage="Exchange Password")]
    [ValidateNotNullOrEmpty()]
    [string]$ExchangePassword,
    [Parameter(Mandatory=$false, Position=2, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t[New-ExchangeServiceObject]"}

  [array]$arrReturnValue = @($false)

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Exchange username:`t[$ExchangeUsername]"}
  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Exchange password:`t[********]"}

  Try {[string]$strIncludeDirectory = Split-Path $script:MyInvocation.MyCommand.Path} Catch {}
  
  Try
  {
    [void][system.reflection.Assembly]::UnsafeLoadFrom("$strIncludeDirectory\MsExchange\Microsoft.Exchange.WebServices.dll")
    [object]$objExchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList Exchange2013
    $objExchangeService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList  $ExchangeUsername, $ExchangePassword
    $objExchangeService.AutodiscoverUrl($ExchangeUsername, {$true})
    $arrReturnValue = @($true, $objExchangeService)
  }
  Catch
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage "$_"}
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t[New-ExchangeServiceObject]"}

  Return $arrReturnValue
}

Function New-EmailMessageObject
{
<#
  .Synopsis
    Create an email message object.
  .Description
    The New-EmailMessageObject cmdlet creates an email message object.
  .Example
    New-EmailMessageObject -ExchangeServiceObject ExchangeServiceObject -From from -ToRecipients to -Subject subject -Log $true
    Create mail message object and log messages
  .Example
    New-EmailMessageObject -ExchangeServiceObject ExchangeServiceObject -From from -ToRecipients to -Subject subject
    Create mail message object
  .Parameter ExchangeServiceObject
    Exchange service object.
  .Parameter From
    Mail message from.
  .Parameter ToRecipients
    Mail message to recipients.
  .Parameter CcRecipients
    Mail message cc recipients.
  .Parameter BccRecipients
    Mail message bcc recipients.
  .Parameter Subject
    Mail message subject.
  .Parameter HtmlBody
    Mail message HTML body.
  .Parameter Body
    Mail message body.
  .Parameter Attachment
    Mail message attachment full file path
  .Parameter Log
    Array of constant log values.
  .Inputs
    [object]
    [string]
    [array]
    [array]
    [array]
    [string]
    [boolean]
    [string]
    [string]
    [array]
  .Output
    [array]
  .Notes
    NAME: New-EmailMessageObject
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170201
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Exchange service object")]
    [ValidateNotNullOrEmpty()]
    [object]$ExchangeServiceObject,
    [Parameter(Mandatory=$true, Position=1, HelpMessage="From")]
    [ValidateNotNullOrEmpty()]
    [string]$From,
    [Parameter(Mandatory=$true, Position=2, HelpMessage="To recipients")]
    [ValidateNotNullOrEmpty()]
    [array]$ToRecipients,
    [Parameter(Mandatory=$false, Position=3, HelpMessage="Cc recipients")]
    [ValidateNotNullOrEmpty()]
    [array]$CcRecipients,
    [Parameter(Mandatory=$false, Position=4, HelpMessage="Bcc recipients")]
    [ValidateNotNullOrEmpty()]
    [array]$BccRecipients,
    [Parameter(Mandatory=$true, Position=5, HelpMessage="Subject")]
    [ValidateNotNullOrEmpty()]
    [string]$Subject,
    [Parameter(Mandatory=$false, Position=6, HelpMessage="HTML body")]
    [ValidateNotNullOrEmpty()]
    [boolean]$HtmlBody = $false,
    [Parameter(Mandatory=$true, Position=7, HelpMessage="Body")]
    [ValidateNotNullOrEmpty()]
    [string]$Body = $false,
    [Parameter(Mandatory=$false, Position=8, HelpMessage="Attachment full file path")]
    [ValidateNotNullOrEmpty()]
    [string]$Attachment,
    [Parameter(Mandatory=$false, Position=9, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t[New-EmailMessageObject]"}

  [array]$arrReturnValue = @($false)

  Try {[string]$strIncludeDirectory = Split-Path $script:MyInvocation.MyCommand.Path} Catch {}
  
  Try
  {
    [void][system.reflection.Assembly]::UnsafeLoadFrom("$strIncludeDirectory\MsExchange\Microsoft.Exchange.WebServices.dll")
    [object]$objEmailMessage = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $ExchangeServiceObject
    [object]$objEmailBody = New-Object Microsoft.Exchange.WebServices.Data.MessageBody

    $objEmailMessage.From = $From
    $ToRecipients | Foreach-Object {$objEmailMessage.ToRecipients.Add($_) | Out-Null}
    If ($CcRecipients.Count) {$CcRecipients | Foreach-Object {$objEmailMessage.CcRecipients.Add($_) | Out-Null}}
    If ($BccRecipients.Count) {$BccRecipients | Foreach-Object {$objEmailMessage.BccRecipients.Add($_) | Out-Null}}
    $objEmailMessage.Subject = $Subject

    If ($HtmlBody) {$objEmailBody.BodyType = "HTML"}
    Else {$objEmailBody.BodyType = "Text"}
    $objEmailBody.Text = $Body
    $objEmailMessage.Body = $objEmailBody
    If ($Attachment) {$objEmailMessage.Attachments.AddFileAttachment($Attachment.Split("\")[-1],$Attachment) | Out-Null}
    $arrReturnValue = @($true, $objEmailMessage)
  }
  Catch
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage "$_"}
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t[New-EmailMessageObject]"}
  
  Return $arrReturnValue
}