# Script Framework
# Name    : Zip.psm1
# Version : 0.1
# Date    : 2017-03-21
# Author  : Roeland van den Bosch
# Website : http://www.roelandvdbosch.nl

Function Start-Zip
{
<#
  .Synopsis
    Zip a folder.
  .Description
    The Start-Zip cmdlet creates a zip file
    from a folder.
  .Example
    Start-Zip -Source C:\source -Destination C:\destination.zip -Log @($true,$objLogValue)
    Creates a zip file C:\destination.zip from
    source folder C:\source
  .Parameter Source
    Source folder.
  .Parameter Destination
    Destination file.
  .Parameter Log
    Array of constant log values.
  .Inputs
    [string]
    [string]
    [array]
  .OutPuts
    [boolean]
  .Notes
    NAME: Start-Zip
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170321
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Source")]
    [ValidateNotNullOrEmpty()]
    [string]$Source,
    [Parameter(Mandatory=$true, Position=1, HelpMessage="Destination")]
    [ValidateNotNullOrEmpty()]
    [string]$Destination,
    [Parameter(Mandatory=$false, Position=2, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
    )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t[Start-Zip]"}

  [boolean]$blnReturnValue = $false
  Try
  {
    Add-Type -AssemblyName 'system.io.compression.filesystem'
    [io.compression.zipfile]::CreateFromDirectory($Source, $Destination)
    $blnReturnValue = $true
  }
  Catch
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage $_}
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t[Start-Zip]"}

  Return [boolean]$blnReturnValue
}

Function Start-Unzip
{
<#
  .Synopsis
    Unzip a folder.
  .Description
    The Start-Unzip cmdlet unzips a folder
    from a zip file.
  .Example
    Start-Unzip -Source C:\source.zip -Destination C:\destination -Log @($true,$objLogValue)
    Unzips a file C:\source.zip to
    destination folder C:\destination
  .Parameter Source
    Source file.
  .Parameter Destination
    Destination folder.
  .Parameter Log
    Array of constant log values.
  .Inputs
    [string]
    [string]
    [array]
  .OutPuts
    [boolean]
  .Notes
    NAME: Start-Zip
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170321
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Source")]
    [ValidateNotNullOrEmpty()]
    [string]$Source,
    [Parameter(Mandatory=$true, Position=1, HelpMessage="Destination")]
    [ValidateNotNullOrEmpty()]
    [string]$Destination,
    [Parameter(Mandatory=$false, Position=2, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
    )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t[Start-Unzip]"}

  [boolean]$blnReturnValue = $false
  Try
  {
    Add-Type -AssemblyName 'system.io.compression.filesystem'
    [io.compression.zipfile]::ExtractToDirectory($Source, $Destination)
    $blnReturnValue = $true
  }
  Catch
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage $_}
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t[Start-Unzip]"}

  Return [boolean]$blnReturnValue
}
