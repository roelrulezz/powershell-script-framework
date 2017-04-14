# Script Framework
# Name    : AltirisLog.psm1
# Version : 0.1
# Date    : 2017-03-21
# Author  : Roeland van den Bosch
# Website : http://www.roelandvdbosch.nl

Function Start-AltirisLog
{
<#
  .Synopsis
    Create start entry in the Altiris log.
  .Description
    The Start-AltirisLog cmdlet creates the first
    entry in the Altiris log.
  .Example
    Start-AltirisLog -Loglevel "ERROR"
    creates the first entry in the Altiris log.
  .Parameter Loglevel
    Level of logging.
  .Inputs
    [string]
  .Notes
    NAME: Start-AltirisLog
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170321
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

  # Check if the Altiris Deployment Agent service is running
  # The service name can be "AClient" or "Altiris Deployment Agent"
  If ((!(Get-Service -Include "AClient")) -and (!(Get-Service -Include "*Deployment Agent*")))
  {
    Throw "Could not locate Deployment Agent service (AClient or DAgent)"
  }

  Else
  {
    # In WinPE the Altiris logging utility is only present in the root of the system drive
    # Therefore add the system drive to the path environment variable, and test for wlogev*.exe,
    # which includes wlogevnt.exe and wlogevent.exe
    If (Test-Path env:systemdrive) {$env:path += ";" + $env:systemdrive}

    # Check if either wlogevnt.exe (old name) or wlogevent can be located via the path
    $global:wlogEventExe = $null
    If ((Get-Command "wlogev*nt.exe" -ErrorAction SilentlyContinue) -ne $null) {$global:wlogEventExe = (Get-Command "wlogev*nt.exe").Name}
    If (!($global:wlogEventExe)) {throw ("Could not locate Altiris logging utility: " + ($wlogEventExeTests -join " or "))}
  }
    
  cmd.exe /c "$global:wlogEventExe -l:1 -ss:`"############################################################`""
  cmd.exe /c "$global:wlogEventExe -l:1 -ss:`"Log Level [$Loglevel]`""
}

Function Write-AltirisLog
{
<#
  .Synopsis
    Write entry in Altiris log.
  .Description
    The Write-AltirisLog cmdlet creates an
    entry in the Altiris log.
  .Example
    Write-AltirisLog -LogLevel "ERROR" -LogMessage "This is going totally wrong!!!"
    creates a log entry with sevirety "ERROR" and message "This is going totally wrong!!!".
  .Parameter Loglevel
    Level of logging.
  .Parameter LogMessage
    Log message.
  .Inputs
    [string]
    [string]
  .Notes
    NAME: Write-AltirisLog
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170321
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
  #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Loglevel")]
    [ValidateNotNullOrEmpty()]
    [string]$Loglevel,
    [Parameter(Mandatory=$true, Position=1, HelpMessage="Log Message")]
    [ValidateNotNullOrEmpty()]
    [string]$LogMessage
  )  

  # Replace all SQL keywords with the keyword between underscores
  Foreach ($sqlKeyword in "drop", "truncate", "create", "revoke", "grant", "deny", `
                          "shutdown", "kill", "select", "update", "delete", "insert", `
                          "execute", "exec", "call", "alter", "dbcc", "while" )
  {
    # Use negative lookbehind and negative lookahead for a word character around the SQL keyword
    # - ^(.*)  : anything before the keyword
    # - ((?<!\w)" + $sqlKeyword + "(?!\w)) : the keyword, without a word character or underscore
    #                                        before or after the keyword
    # - (.*)$  : anything after the keyword
    # See, e.g., http://www.regular-expressions.info/refadv.html
    While ( $LogMessage -match ( "^(.*)((?<!\w)" + $sqlKeyword + "(?!\w))(.*)$" ) )
    {
      # Put underscores around the keyword
      $LogMessage = $matches[ 1 ] + "_" + $matches[ 2 ] + "_" + $matches[ 3 ]
    }
  }

  # Replace two dashes with two underscores
  $LogMessage = $LogMessage.Replace( "--", "__" )
  # Replace a semicolon with "|"
  $LogMessage = $LogMessage.Replace( ";", "|" )
  # Replace a double quote with a single quote
  $LogMessage = $LogMessage.Replace( "`"", "'" )

  # The Altiris console only shows the first 250 characters, so limit the output for safety
  If ($LogMessage.Length -gt 250) {$LogMessage = $LogMessage.Substring(0, 250)}
  
  Switch ($LogLevel)
  {
    "DEBUG" {cmd.exe /c "$global:wlogEventExe -l:0 -ss:`"$LogMessage`""}
    "INFO" {cmd.exe /c "$global:wlogEventExe -l:1 -ss:`"$LogMessage`""}
    "WARNING" {cmd.exe /c "$global:wlogEventExe -l:2 -ss:`"$LogMessage`""}
    "ERROR" {cmd.exe /c "$global:wlogEventExe -l:3 -ss:`"$LogMessage`" -c:1"}
    Default {cmd.exe /c "$global:wlogEventExe -l:0 -ss:`"$LogMessage`""}
  }
}

Function Stop-AltirisLog
{
<#
  .Synopsis
    Create last entry in Altiris log.
  .Description
    The Stop-AltirisLog cmdlet creates the last
    entry in the Altiris log.
  .Example
    Stop-AltirisLog
    creates the last entry in the Altiris log.
  .Notes
    NAME: Stop-AltirisLog
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170321
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
  #Requires -Version 2.0
#> 
  cmd.exe /c "$global:wlogEventExe -l:1 -ss:`"############################################################`""
  Start-Sleep -Seconds 1
}