# Script Framework
# Name    : FileLog.psm1
# Version : 0.1
# Date    : 2017-02-19
# Author  : Roeland van den Bosch
# Website : http://www.roelandvdbosch.nl

Function Start-FileLog
{
<#
  .Synopsis
    Create start entry in the logfile.
  .Description
    The Start-FileLog cmdlet creates the first
    entry in the logfile.
  .Example
    Start-FileLog -Loglevel "ERROR" -Logfile "C:\logFile.txt"
    creates the first entry in the logfile.
  .Parameter Loglevel
    Level of logging.
  .Parameter Logfile
    Fullpath for the logfile.
  .Inputs
    [string]
    [string]
  .Notes
    NAME: Start-FileLog
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
    [string]$Loglevel,
    [Parameter(Mandatory=$true, Position=1, HelpMessage="Logfile")]
    [ValidateNotNullOrEmpty()]
    [string]$Logfile
    )
  "$(Get-Date -format yyyyMMddhhmmss) ############################################################" | Out-File -File $Logfile -Append
  "$(Get-Date -format yyyyMMddhhmmss) INFO    - Log Level [$Loglevel]" | Out-File -File $Logfile -Append
}

Function Write-FileLog
{
<#
  .Synopsis
    Write entry in logfile.
  .Description
    The Write-FileLog cmdlet creates an
    entry in the logfile.
  .Example
    Write-FileLog -LogFile "C:\logFile.txt" -LogLevel "ERROR" -LogMessage "This is going totally wrong!!!"
    creates a log entry with sevirety "ERROR" and message "This is going totally wrong!!!".
  .Parameter Logfile
    Fullpath for the logfile.
  .Parameter Loglevel
    Level of logging.
  .Parameter LogMessage
    Log message.
  .Inputs
    [string]
    [string]
    [string]
  .Notes
    NAME: Write-FileLog
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
    [string]$Loglevel,
    [Parameter(Mandatory=$true, Position=1, HelpMessage="Logfile")]
    [ValidateNotNullOrEmpty()]
    [string]$Logfile,
    [Parameter(Mandatory=$true, Position=2, HelpMessage="Log Message")]
    [ValidateNotNullOrEmpty()]
    [string]$LogMessage
  )  
  Switch ($LogLevel)
  {
    "DEBUG" {"$(Get-Date -format yyyyMMddhhmmss) $Loglevel   - $LogMessage" | Out-File -File $Logfile -Append}
    "INFO" {"$(Get-Date -format yyyyMMddhhmmss) $Loglevel    - $LogMessage" | Out-File -File $Logfile -Append}
    "WARNING" {"$(Get-Date -format yyyyMMddhhmmss) $Loglevel - $LogMessage" | Out-File -File $Logfile -Append}
    "ERROR" {"$(Get-Date -format yyyyMMddhhmmss) $Loglevel   - $LogMessage" | Out-File -File $Logfile -Append}
    Default {"$(Get-Date -format yyyyMMddhhmmss) UNKNOWN - $LogMessage" | Out-File -File $Logfile -Append}
  }
}

Function Stop-FileLog
{
<#
  .Synopsis
    Create last entry in logfile.
  .Description
    The Stop-Filelog cmdlet creates the last
    entry in the logfile.
  .Example
    Stop-Filelog -Logfile "C:\logFile.txt"
    creates the last entry in the logfile.
  .Parameter Logfile
    Fullpath for the logfile.
  .Inputs
    [string]
  .Notes
    NAME: Stop-FileLog
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170219
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
  #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Logfile")]
    [ValidateNotNullOrEmpty()]
    [string]$Logfile
    )  
    "$(Get-Date -format yyyyMMddhhmmss) ############################################################" | Out-File -File $Logfile -Append
}