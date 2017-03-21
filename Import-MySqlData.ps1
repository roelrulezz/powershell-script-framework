##############################################################
# Get parameters                                             #
##############################################################

#region Get parameters

Param(
  [Parameter(Mandatory=$false, Position=0, HelpMessage="Ini file")]
  [ValidateNotNullOrEmpty()]
  [string]$IniFile = '',
  [Parameter(Mandatory=$false, Position=1, HelpMessage="CSV file path")]
  [ValidateNotNullOrEmpty()]
  [string]$CsvFilePath = '',
  [Parameter(Mandatory=$false, Position=2, HelpMessage="Database tables")]
  [ValidateNotNullOrEmpty()]
  [string]$DatabaseTables = ''
)

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
[string]$strCreditCompany = 'Host Your IT'
[string]$strCreditAuthor = 'Roeland van den Bosch'
[string]$strCreditDate = '2017-03-06'
[string]$strCreditVersion = '0.1'
[string]$strTemplateVersion = '0.1'

#endregion Set credit variable 

##############################################################
# Set script variable                                        #
##############################################################

#region Set script variable

[string]$strLoglevel = "INFO" # ERROR / WARNING / INFO / DEBUG / NONE
[boolean]$blnLogToConsole = $true
[boolean]$blnLogToFile = $true
[string]$strLogfile = "$strScriptDirectory\$($strScriptName.Replace('.ps1','.log'))"

[string]$strModulePath = "$strScriptDirectory\Module"

#If ($IniFile -eq '') {[string]$strIniFilePath = "$strScriptDirectory\HostYourIT\hostyourit_db.ini"}
If ($IniFile -eq '') {[string]$strIniFilePath = "\\10.14.1.15\data\HostYourIT\Administratie\Facturen\hostyourit_db.ini"}
Else {[string]$strIniFilePath = $IniFile}
#If ($IniFile -eq '') {[string]$strCsvFilePath = "$strScriptDirectory\HostYourIT"}
If ($IniFile -eq '') {[string]$strCsvFilePath = "\\10.14.1.15\data\HostYourIT\Administratie\Facturen"}
Else {[string]$strCsvFilePath = $CsvFilePath}
If ($DatabaseTables -eq '') {[array]$arrDatabaseTables = @("Klanten","Prijzen_2016","Facturen_2016","Prijzen_2017","Facturen_2017")}
Else {[array]$arrDatabaseTables = $DatabaseTables.Split(',')}
#endregion Set script variable

##############################################################
# Initialize log                                             #
##############################################################

#region Initialize log

[int]$intLogtype = 0

If ($blnLogToConsole) {$intLogtype += 1}
If ($blnLogToFile) {$intLogtype += 2}

[boolean]$blnLog = ($blnLogToFile -or $blnLogToConsole) -and $intLogtype -gt 0

[object]$objLogValue = @{'Loglevel' = $strLoglevel;
                         'Logtype' = $intLogtype;
                         'Logfile' = $strLogfile}

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
  Write-Host "Company:`t$strCreditCompany" -ForegroundColor Yellow
  Write-Host "Author:`t`t$strCreditAuthor" -ForegroundColor Yellow
  Write-Host "Date:`t`t$strCreditDate"  -ForegroundColor Yellow
  Write-Host "Version:`t$strCreditVersion" -ForegroundColor Yellow
  Write-Host "###########################################################`r`n" -ForegroundColor Yellow
}

#endregion Default function

#endregion Function

##############################################################
# Main                                                       #
##############################################################

If ($blnCreditShow) {Show-Credit}

If ($blnLog) {Start-Log -LogValue $objLogValue}

[boolean]$blnDatabaseConnection = $false

#region Read ini file
If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Read ini file]"}

[object]$objIniFile = Read-IniFile -FilePath $strIniFilePath
[string]$strDatabaseServer = $($objIniFile.default.server)
[string]$strDatabaseName = $($objIniFile.Database.name)
[string]$strDatabaseUsername = $($objIniFile.User.user)
[string]$strDatabaseUserEncryptedPassword = $($objIniFile.User.password)

If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Read ini file]"}
#endregion Read ini file

#region Connect to database server
If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Connect to database server]"}

[array]$arrReturnValueDatabaseConnection = New-DatabaseConnection `
                                             -DatabaseServer $strDatabaseServer `
                                             -DatabaseName $strDatabaseName `
                                             -Username $strDatabaseUsername `
                                             -Password $(New-DecryptedString -InputString $strDatabaseUserEncryptedPassword -Passphrase $strDatabaseUsername) `
                                             -Log @($true, $objLogValue)

If ($arrReturnValueDatabaseConnection[0])
{
  [object]$objDatabaseConnection = $arrReturnValueDatabaseConnection[1]
  $blnDatabaseConnection = $true
}
Else
{
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "WARNING" -LogMessage "NO database connection"}
}
If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Connect to database server]"}
#endregion Connect to database server

If ($blnDatabaseConnection)
{
  Foreach ($strDatabaseTable in $arrDatabaseTables)
  {
    #region Read CSV file
    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Read CSV file]"}

    [object]$objCsvFile = Import-Csv -Path "$strCsvFilePath\$strDatabaseTable.csv" -Delimiter ";"
    [array]$arrCsvHeader = ($objCsvFile | Get-Member -MemberType NoteProperty).Name

    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Read CSV file]"}
    #endregion Read CSV file

    #region Insert data
    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Insert data]"}

    Foreach ($arrCsvFile in $objCsvFile)
    {
      [string]$strSqlQueryColumns = "INSERT INTO $strDatabaseTable ("
      [string]$strSqlQueryValues = "VALUES ("
      Foreach ($strCsvHeader in $arrCsvHeader)
      {
        $strSqlQueryColumns += "$strCsvHeader,"
        If ($arrCsvFile.$strCsvHeader -eq 0 -or $arrCsvFile.$strCsvHeader -eq 1 -or $arrCsvFile.$strCsvHeader -eq 'NULL')
        {
          $strSqlQueryValues += "$($arrCsvFile.$strCsvHeader),"
        }
        Else
        {
          $strSqlQueryValues += "'$(($arrCsvFile.$strCsvHeader).Replace("'","\'"))',"
        }
      }
      [string]$strSqlQuery = "$($strSqlQueryColumns.Substring(0,$strSqlQueryColumns.Length -1))) $($strSqlQueryValues.Substring(0,$strSqlQueryValues.Length -1)))"
  
      [array]$arrReturnValueDatabaseQuery = New-DatabaseQuery `
                                              -DatabaseConnection $objDatabaseConnection `
                                              -SqlQuery $strSqlQuery `
                                              -Log @($true, $objLogValue)

      If ($arrReturnValueDatabaseQuery[0])
      {
        If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Successfully add record"}
      }
      Else
      {
        If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "WARNING" -LogMessage "FAILED add record"}
      }
    }

    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Insert data]"}
    #endregion Insert data
  }
  #region Disconnect from database server
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Disconnect from database server]"}
  $objDatabaseConnection.Close()
  $blnDatabaseConnection = $false
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Disconnect from database server]"}
  #endregion Disconnect from database server

}

If ($blnLog) {Stop-Log -LogValue $objLogValue}
