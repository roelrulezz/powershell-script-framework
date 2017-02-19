# Script Framework
# Name    : ConsoleLog.psm1
# Version : 0.1
# Date    : 2017-02-19
# Author  : Roeland van den Bosch
# Website : http://www.roelandvdbosch.nl

Function Start-ConsoleLog
{
<#
  .Synopsis
    Create first entry in the console log.
  .Description
    The Start-ConsoleLog cmdlet creates the first
    entry in the console log.
  .Example
    Start-ConsoleLog -Loglevel "ERROR"
    Creates the first entry of the console log.
  .Parameter Loglevel
    Level of logging.
  .Inputs
    [string]
  .Notes
    NAME: Start-ConsoleLog
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170219
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Loglevel")]
    [ValidateNotNullOrEmpty()]
    [string]$Loglevel
    )
  Write-Host "$(Get-Date -format yyyyMMddhhmmss) ############################################################" -ForegroundColor White
  Write-Host "$(Get-Date -format yyyyMMddhhmmss) INFO    - Log Level [$Loglevel]" -ForegroundColor White
}

Function Write-ConsoleLog
{
<#
  .Synopsis
    Write entry in console log.
  .Description
    The Write-ConsoleLog cmdlet creates an
    entry in the console log.
  .Example
    Write-ConsoleLog -Loglevel "ERROR" -LogMessage "This is going totally wrong!!!"
    Creates a log entry with sevirety "ERROR" and message "This is going totally wrong!!!".
  .Parameter Loglevel
    Level of logging.
  .Parameter LogMessage
    Log message.
  .Inputs
    [string]
    [string]
  .Notes
    NAME: Write-ConsoleLog
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170219
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Log Level")]
    [ValidateNotNullOrEmpty()]
    [string]$Loglevel,
    [Parameter(Mandatory=$true, Position=1, HelpMessage="Log Message")]
    [ValidateNotNullOrEmpty()]
    [string]$LogMessage
  )  
  Switch ($LogLevel)
  {
    "DEBUG" {Write-Host "$(Get-Date -format yyyyMMddhhmmss) $Loglevel   - $LogMessage" -ForegroundColor Magenta}
    "INFO" {Write-Host "$(Get-Date -format yyyyMMddhhmmss) $Loglevel    - $LogMessage" -ForegroundColor White}
    "WARNING" {Write-Host "$(Get-Date -format yyyyMMddhhmmss) $Loglevel - $LogMessage" -ForegroundColor Yellow}
    "ERROR" {Write-Host "$(Get-Date -format yyyyMMddhhmmss) $Loglevel   - $LogMessage" -ForegroundColor Red}
    Default {Write-Host "$(Get-Date -format yyyyMMddhhmmss) UNKNOWN   - $LogMessage" -ForegroundColor DarkGray}
  }
}

Function Stop-ConsoleLog
{
<#
  .Synopsis
    Create last entry in the console log.
  .Description
    The Stop-ConsoleLog cmdlet creates the last
    entry in the console log.
  .Example
    Stop-ConsoleLog
    creates the last entry in the console log.
  .Notes
    NAME: Stop-ConsoleLog
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170219
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Write-Host "$(Get-Date -format yyyyMMddhhmmss) ############################################################" -ForegroundColor White
  Write-Host ""
}