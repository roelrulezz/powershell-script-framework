# Script Framework
# Name    : Log.psm1
# Version : 0.1
# Date    : 2017-03-21
# Author  : Roeland van den Bosch
# Website : http://www.roelandvdbosch.nl

Function Start-Log
{
<#
  .Synopsis
    First entry in the log.
  .Description
    The Start-Log cmdlet creates the first
    entry in the log.
  .Example
    Start-Log -LogValue @{'Loglevel' = "ERROR";'Logtype' = 1;'Logfile' = ''}
    Creates the first entry in the console log.
  .Example
    Start-Log -LogValue @{'Loglevel' = "ERROR";'Logtype' = 2;'Logfile' = $strLogfile}
    Creates the first entry in the logfile.
  .Example
    Start-Log -LogValue @{'Loglevel' = "ERROR";'Logtype' = 3;'Logfile' = $strLogfile}
    Creates the first entry in the console log and logfile.
  .Parameter LogValue
    Object of constant log values.
  .Inputs
    [object]
  .Notes
    NAME: Start-Log
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170321
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
  #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="LogValue")]
    [ValidateNotNullOrEmpty()]
    [object]$LogValue
  )

  [string]$strLogtypeConsole = [convert]::ToString(1,2)
  [string]$strLogtypeFile = [convert]::ToString(2,2)
  [string]$strLogtypeAltiris = [convert]::ToString(4,2)

  Switch ([convert]::ToString([int32]$LogValue.Logtype,2))
  {
    {$_ -band $strLogtypeConsole} {Start-ConsoleLog -LogLevel $LogValue.Loglevel}
    {$_ -band $strLogtypeFile} {Start-FileLog -LogLevel $LogValue.Loglevel -LogFile $LogValue.Logfile}
    {$_ -band $strLogtypeAltiris} {Start-AltirisLog -LogLevel $LogValue.Loglevel}
  }
}

Function Write-Log
{
<#
  .Synopsis
    Write entry in log.
  .Description
    The Write-Log cmdlet creates an
    entry in the log.
  .Example
    Write-FileLog -LogValue @{'Loglevel' = "ERROR";'Logtype' = 1;'Logfile' = ''} -LogMessageLevel "ERROR" -LogMessage "This is going totally wrong!!!"
    Creates a console log entry with sevirety "ERROR" and message "This is going totally wrong!!!".
  .Example
    Write-FileLog -LogValue @{'Loglevel' = "ERROR";'Logtype' = 2;'Logfile' = $strLogfile} -LogMessageLevel "ERROR" -LogMessage "This is going totally wrong!!!"
    Creates a logfile entry with sevirety "ERROR" and message "This is going totally wrong!!!".
  .Example
    Write-FileLog -LogValue @{'Loglevel' = "ERROR";'Logtype' = 3;'Logfile' = $strLogfile} -LogMessageLevel "ERROR" -LogMessage "This is going totally wrong!!!"
    Creates a console log and filelog entry with sevirety 
    "ERROR" and message "This is going totally wrong!!!".
  .Parameter LogValue
    Object of constant log values.
  .Parameter LogMessageLevel
    Log message level.
  .Parameter LogMessage
    Log message.
  .Inputs
    [object]
    [int]
    [string]
  .Notes
    NAME: Write-Log
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170321
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
  #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="LogValue")]
    [ValidateNotNullOrEmpty()]
    [object]$LogValue,
    [Parameter(Mandatory=$true, Position=1, HelpMessage="Log message level")]
    [ValidateNotNullOrEmpty()]
    [string]$LogMessageLevel,
    [Parameter(Mandatory=$true, Position=2, HelpMessage="Log message")]
    [ValidateNotNullOrEmpty()]
    [string]$LogMessage
  )  

  [object]$objLogLevel = @{"ERROR" = 1;"WARNING" = 2;"INFO" = 3;"DEBUG" = 4}

  If ($objLogLevel.$LogMessageLevel -le $objLogLevel.$($LogValue.Loglevel))
  {
    [string]$strLogtypeConsole = [convert]::ToString(1,2)
    [string]$strLogtypeFile = [convert]::ToString(2,2)
    [string]$strLogtypeAltiris = [convert]::ToString(4,2)

    Switch ([convert]::ToString([int32]$LogValue.Logtype,2))
    {
      {$_ -band $strLogtypeConsole} {Write-ConsoleLog -LogLevel $LogMessageLevel -LogMessage $LogMessage}
      {$_ -band $strLogtypeFile} {Write-FileLog -LogLevel $LogMessageLevel -LogFile $LogValue.Logfile -LogMessage $LogMessage}
      {$_ -band $strLogtypeAltiris} {Write-AltirisLog -LogLevel $LogMessageLevel -LogMessage $LogMessage}
    }
  }
}

Function Stop-Log
{
<#
  .Synopsis
    Last entry in the log.
  .Description
    The Stop-Log cmdlet creates the last
    entry in the log.
  .Example
    Stop-Log -LogValue @{'Loglevel' = "ERROR";'Logtype' = 1;'Logfile' = $strLogfile}
    Creates the last entry in the console log.
  .Example
    Stop-Log -LogValue @{'Loglevel' = "ERROR";'Logtype' = 2;'Logfile' = $strLogfile}
    Creates the last entry in the logfile.
  .Example
    Stop-Log -LogValue @{'Loglevel' = "ERROR";'Logtype' = 3;'Logfile' = $strLogfile}
    Creates the last entry in the console log and logfile.
  .Parameter LogValue
    Object of constant log values.
  .Inputs
    [object]
  .Notes
    NAME: Stop-Log
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170321
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
  #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="LogValue")]
    [ValidateNotNullOrEmpty()]
    [object]$LogValue
  )
  
  [string]$strLogtypeConsole = [convert]::ToString(1,2)
  [string]$strLogtypeFile = [convert]::ToString(2,2)
  [string]$strLogtypeAltiris = [convert]::ToString(4,2)

  Switch ([convert]::ToString([int32]$LogValue.Logtype,2))
  {
    {$_ -band $strLogtypeConsole} {Stop-ConsoleLog}
    {$_ -band $strLogtypeFile} {Stop-FileLog -LogFile $LogValue.Logfile}
    {$_ -band $strLogtypeAltiris} {Stop-AltirisLog}
  }
}