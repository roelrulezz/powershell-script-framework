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
[string]$strCreditCompany = 'Host Your IT'
[string]$strCreditAuthor = 'Roeland van den Bosch'
[string]$strCreditDate = '2017-02-19'
[string]$strCreditVersion = '0.1'
[string]$strTemplateVersion = '0.1'

#endregion Set credit variable 

##############################################################
# Set script variable                                        #
##############################################################

#region Set script variable

[string]$strLoglevel = "DEBUG" # ERROR / WARNING / INFO / DEBUG / NONE
[boolean]$blnLogToConsole = $true
[boolean]$blnLogToFile = $true
[string]$strLogfile = "$strScriptDirectory\$($strScriptName.Replace('.ps1','.log'))"

[string]$strModulePath = "$strScriptDirectory\Module"

[string]$strIniFilePath = "$strScriptDirectory\HostYourIT\hostyourit_db.ini"

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
[boolean]$blnContinue = $false

#region Read ini file
If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Read ini file]"}

[object]$objIniFile = Read-IniFile -FilePath $strIniFilePath
[string]$strDatabaseServer = $($objIniFile.default.server)
[string]$strDatabaseUser = $($objIniFile.default.user)
[string]$strDatabaseEncryptedPassword = $($objIniFile.default.password)
[string]$strDatabaseEncryptedPasswordPassphrase = $($objIniFile.Encryption.passphrase)
[string]$strDatabaseName = $($objIniFile.Database.name)
[string]$strDatabaseUsername = $($objIniFile.User.user)
[string]$strDatabaseUserEncryptedPassword = $($objIniFile.User.password)
[array]$arrDatabaseTables = $($objIniFile.Tables).Keys

If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Read ini file]"}
#endregion Read ini file

#region Connect to database server
If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Connect to database server]"}

[array]$arrReturnValueDatabaseConnection = New-DatabaseConnection `
                                             -DatabaseServer $strDatabaseServer `
                                             -Username $strDatabaseUser `
                                             -Password $(New-DecryptedString -InputString $strDatabaseEncryptedPassword -Passphrase $strDatabaseEncryptedPasswordPassphrase) `
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
  #region Create database
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Create database]"}

  [array]$arrReturnValueDatabaseQuery = New-DatabaseQuery `
                                          -DatabaseConnection $objDatabaseConnection `
                                          -SqlQuery "CREATE DATABASE IF NOT EXISTS $strDatabaseName" `
                                          -Log @($true, $objLogValue)
  If ($arrReturnValueDatabaseQuery[0])
  {
    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "INFO" -LogMessage "Database [$strDatabaseName] created]"}
    $blnContinue = $true
  }
  Else
  {
    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "ERROR" -LogMessage "Database [$strDatabaseName] NOT created"}
  }
  
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Create database]"}
  #endregion Create database

  If ($blnContinue)
  {
    #region Create database user
    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Create database user]"}
    $blnContinue = $false
    [array]$arrReturnValueDatabaseQuery = New-DatabaseQuery `
                                            -DatabaseConnection $objDatabaseConnection `
                                            -SqlQuery "CREATE USER IF NOT EXISTS '$strDatabaseUsername' IDENTIFIED BY '$strDatabaseUserPassword'" `
                                            -Log @($true, $objLogValue)
    If ($arrReturnValueDatabaseQuery[0])
    {
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "INFO" -LogMessage "User [$strDatabaseUsername] created"}
      $blnContinue = $true
    }
    Else
    {
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "ERROR" -LogMessage "User [$strDatabaseUsername] NOT created"}
    }
    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Create database user]"}
    #endregion Create database user
  }

  If ($blnContinue)
  {
    #region Set database permissions
    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Set database permissions]"}
    $blnContinue = $false

    [array]$arrReturnValueDatabaseQuery = New-DatabaseQuery `
                                            -DatabaseConnection $objDatabaseConnection `
                                            -SqlQuery "GRANT ALL PRIVILEGES ON $strDatabaseName.* TO '$strDatabaseUsername' WITH GRANT OPTION" `
                                            -Log @($true, $objLogValue)
      
    If ($arrReturnValueDatabaseQuery[0])
    {
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "INFO" -LogMessage "Permissions for user [$strDatabaseUsername] set"}
      $blnDatabaseCreated = $true
      $blnContinue = $true
    }
    Else
    {
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "ERROR" -LogMessage "Permissions for user [$strDatabaseUsername] NOT set"}
    }
    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Set database permissions]"}
    #endregion Create database user
  }

  #region Disconnect from database server
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Disconnect from database server]"}
  $objDatabaseConnection.Close()
  $blnDatabaseConnection = $false
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Disconnect from database server]"}
  #endregion Disconnect from database server
}

If ($blnContinue)
{
  #region Reconnect to database server
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Reconnect to database server]"}

  [array]$arrReturnValueDatabaseConnection = New-DatabaseConnection `
                                               -DatabaseServer $strDatabaseServer `
                                               -DatabaseName $strDatabaseName `
                                               -Username $strDatabaseUsername `
                                               -Password $(New-DecryptedString -InputString $strDatabaseUserEncryptedPassword -Passphrase $strDatabaseEncryptedPasswordPassphrase) `
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
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Reconnect to database server]"}
  #endregion Reconnect to database server
}

If ($blnDatabaseConnection)
{
  #region Create table SQL query
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Create table SQL query]"}

  [array]$arrSqlQuery = @()

  Foreach ($strTable in $arrDatabaseTables)
  {
    If ($($objIniFile."$strTable".Keys | Where-Object {$_ -notlike "Comment_*"}).Count -ne 0)
    {
      [string]$strSqlQuery = "CREATE TABLE IF NOT EXISTS $strTable ("
      [boolean]$blnContainsPrimaryKey = $false
      [string]$strTablePrimaryKey = "PRIMARY KEY ("
      Foreach ($strColumnName in $($objIniFile."$strTable".Keys | Where-Object {$_ -notlike "Comment_*"}))
      {
        [array]$arrColumnFormat = $objIniFile."$strTable"."$strColumnName".Split(";")
        If ($arrColumnFormat[2] -ne 0)
        {
          $blnContainsPrimaryKey = $true
          $strTablePrimaryKey += "$strColumnName, "
        }
        $strSqlQuery += "$strColumnName $($arrColumnFormat[0])$(If ($arrColumnFormat[1] -ne 0) {"($($arrColumnFormat[1]))"})$(If ($arrColumnFormat[3] -ne 0){" NOT NULL"})$(If ($arrColumnFormat[4] -ne 0){" AUTO_INCREMENT"}), "
      }
      If ($blnContainsPrimaryKey)
      {
        $strSqlQuery += "$($strTablePrimaryKey.Substring(0,$strTablePrimaryKey.Length -2))))"
      }
      Else
      {
        $strSqlQuery = "$($strSqlQuery.Substring(0,$strSqlQuery.Length -2)))"
      }
      $arrSqlQuery += @($strSqlQuery)
    }
    Else
    {
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "WARNING" -LogMessage "NO table [$strTable] format information"}
    }
  }
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Create table SQL query]"}
  #endregion Create table SQL query

  #region Create table
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Create table]"}

  If ($arrSqlQuery.Count -gt 0)
  {
    Foreach ($strSqlQuery in $arrSqlQuery)
    {
      [string]$strTableName = $($strSqlQuery.Split(" ")[5])
      [array]$arrReturnValueDatabaseQuery = New-DatabaseQuery `
                                              -DatabaseConnection $objDatabaseConnection `
                                              -SqlQuery $strSqlQuery `
                                              -Log @($true, $objLogValue)

      If ($arrReturnValueDatabaseQuery[0])
      {
        If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "INFO" -LogMessage "Table [$strTableName] created"}
      }
      Else
      {
        If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "WARNING" -LogMessage "Table [$strTableName] NOT created"}
      }
    }
  }

  $objDatabaseConnection.Close()
  $blnDatabaseConnection = $false

  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Create table]"}
  #endregion Create table  

  #region Disconnect from database server
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Disconnect from database server]"}
  $objDatabaseConnection.Close()
  $blnDatabaseConnection = $false
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Disconnect from database server]"}
  #endregion Disconnect from database server
}

If ($blnLog) {Stop-Log -LogValue $objLogValue}
