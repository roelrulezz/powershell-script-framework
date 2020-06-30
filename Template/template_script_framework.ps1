##############################################################
# Get parameters                                             #
##############################################################

#region Get parameters

Param(
  [Parameter(Mandatory=$false, Position=0, HelpMessage="Parameter")]
  [ValidateNotNullOrEmpty()]
  [string]$Parameter = "Default parameter value")

#endregion Get parameters

##############################################################
# Initialize script variable                                 #
##############################################################

#region Initialize script variable

Clear-Host

Try {[string]$strScriptDirectory = Split-Path $script:MyInvocation.MyCommand.Path} Catch {$_}
[string]$strScriptName = $MyInvocation.MyCommand.Name
[string]$strCurrentDate = Get-Date -format yyyyMMddhhmm

#endregion Initialize script variable

##############################################################
# Set credit variable                                        #
##############################################################

#region Set credit variable 

[boolean]$blnCreditShow = $true
[string]$strCreditName = $strScriptName
[string]$strCreditAuthor = 'Roeland van den Bosch'
[string]$strCreditDate = '2020-06-30'
[string]$strCreditVersion = '0.1'
[string]$strTemplateVersion = '1.0'

#endregion Set credit variable 

##############################################################
# Set script variable                                        #
##############################################################

#region Set script variable

[string]$strLoglevel = "DEBUG" # ERROR / WARNING / INFO / DEBUG / NONE
[boolean]$blnLogToConsole = $true
[boolean]$blnLogToFile = $true
[string]$strLogDirectory = "$strScriptDirectory\Log" 
[string]$strLogfile = "$strLogDirectory\$($strScriptName.Replace('.ps1','.log'))"

[string]$strModulePath = "$strScriptDirectory\Module"

#endregion Set script variable

##############################################################
# Initialize log                                             #
##############################################################

#region Initialize log

[int]$intLogtype = 0

If ($blnLogToConsole) {$intLogtype += 1}
If ($blnLogToFile) {$intLogtype += 2}

[boolean]$blnLog = $intLogtype -gt 0

[object]$objLogValue = @{'Loglevel' = $strLoglevel;
                         'Logtype' = $intLogtype;
                         'Logfile' = $strLogfile}

If (!(Test-Path -Path $strLogDirectory))
{
  New-Item -Path $strLogDirectory -ItemType Directory | Out-Null
}

#endregion Initialize log

##############################################################
# Import module                                              #
##############################################################

#region Import module

Foreach ($Module in $(Get-ChildItem -Path $strModulePath -Name "*.psm1"))
{
  Import-Module -Name "$strModulePath\$Module" -Force
}

#endregion Import module

##############################################################
# Add snapin                                                 #
##############################################################

#region Add snapin

#endregion Add snapin

##############################################################
# Function                                                   #
##############################################################

#region Function

#region Default function

Function Show-Credit
{
  Write-Host "###################### Credits ############################" -ForegroundColor Yellow
  Write-Host "Name:`t`t$strCreditName" -ForegroundColor Yellow
  Write-Host "Author:`t`t$strCreditAuthor" -ForegroundColor Yellow
  Write-Host "Date:`t`t$strCreditDate"  -ForegroundColor Yellow
  Write-Host "Version:`t$strCreditVersion" -ForegroundColor Yellow
  Write-Host "###########################################################`r`n" -ForegroundColor Yellow
}

Function Check-ElevatedPowerShell
{
  If (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
  {
    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "INFO" -LogMessage "Running script with elevated privileges"}
    Return $true
  }
  Else
  {
    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "WARNING" -LogMessage "Restart script with elevated privileges"}
    Start-Process PowerShell -verb runas -ArgumentList '-NoProfile', '-ExecutionPolicy', 'ByPass', '-File', "$strScriptDirectory\$strScriptName", '-EnvironmentName', "$EnvironmentName" -Wait
    Return $false
  }
}

Function Stop-Script
{
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start Function:`t[Stop-Script]"}
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "WARNING" -LogMessage "Script has been stopped by an error"}
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End Function:`t`t[Stop-Script]"}
  If ($blnLog) {Stop-Log -LogValue $objLogValue}
  Exit
}

function PressAnyKey()
{
  if ($host.name -notmatch 'ISE')
  {
    Write-Host "########################################################################################" -ForegroundColor Yellow
    Write-Host "Press any key to continue..."
    $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
    $host.UI.RawUI.Flushinputbuffer()
  }
}

#endregion Default function

#endregion Function

##############################################################
# Main                                                       #
##############################################################

If ($blnCreditShow) {Show-Credit}

If ($blnLog) {Start-Log -LogValue $objLogValue}

If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "ERROR"  -LogMessage "Error message"}
If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "WARNING" -LogMessage "Warning message"}
If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "INFO" -LogMessage "Info message"}
If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Debug message"}

#Stop-Script

If ($blnLog) {Stop-Log -LogValue $objLogValue}
