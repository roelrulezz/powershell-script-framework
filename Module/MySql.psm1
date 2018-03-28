# Script Framework
# Name    : MySql.psm1
# Version : 0.1
# Date    : 2017-02-19
# Author  : Roeland van den Bosch
# Website : http://www.roelandvdbosch.nl

Function New-DatabaseConnection
{
<#
  .Synopsis
    Create MySQL Connection.
  .Description
    The New-DatabaseConnection cmdlet creates a connection
    with a MySQL database.
  .Example
    New-DatabaseConnection -DatabaseServer "databaseserver" -DatabaseName "databasename" -Username "username" -Password "password" -Log @($true,$objLogValue)
    Returns True or False, a MySQL datababase object and log messages
  .Example
    New-DatabaseConnection -DatabaseServer "databaseserver" -DatabaseName "databasename" -Username "username" -Password "password"
    Returns True or False and a MySQL datababase object
  .Parameter DatabaseServer
    Name of MySQL databaserver
  .Parameter DatabaseName
    Name of MySQL database
  .Parameter Username
    Name of user
  .Parameter Password
    Password for user in string format
  .Parameter Log
    Array of constant log values.
  .Inputs
    [string]
    [string]
    [string]
    [string]
    [array]
  .OutPuts
    [array]
  .Notes
    NAME: New-DatabaseConnection
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 201700219
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
  #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Database Server")]
    [ValidateNotNullOrEmpty()]
    [string]$DatabaseServer,
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Database Name")]
    [ValidateNotNullOrEmpty()]
    [string]$DatabaseName,
    [Parameter(Mandatory=$true, Position=2, HelpMessage="Username")]
    [ValidateNotNullOrEmpty()]
    [string]$Username,
    [Parameter(Mandatory=$true, Position=3, HelpMessage="Password")]
    [ValidateNotNullOrEmpty()]
    [string]$Password,
    [Parameter(Mandatory=$false, Position=4, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t[New-DatabaseConnection]"}
  
  [array]$arrReturnValue = @($false)

  Try
  {
    Try {[string]$strIncludeDirectory = Split-Path $script:MyInvocation.MyCommand.Path} Catch {}
    [void][system.reflection.Assembly]::UnsafeLoadFrom("$strIncludeDirectory\MySql\MySql.Data.dll")
    [object]$objSqlConnection = New-Object MySql.Data.MySqlClient.MySqlConnection
    $objSqlConnection.ConnectionString = "Server=$DatabaseServer;Database=$DatabaseName;User ID=$Username;Password=$Password;"
    $objSqlConnection.Open()
  }
  Catch
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage $_}
  }
  If ($objSqlConnection.State -eq 'Open')
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "INFO" -LogMessage "Successfully connected [Server=$DatabaseServer;Database=$DatabaseName;User ID=$Username;Password=*****]"}
    $arrReturnValue += @([object]$objSqlConnection)
    $arrReturnValue[0] = $true
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t[New-DatabaseConnection]"}

  Return [array]$arrReturnValue
}

Function New-DatabaseQuery
{
<#
  .Synopsis
    Send SQL query.
  .Description
    The New-DatabaseQuery cmdlet sends a
    SQL query.
  .Example
    New-DatabaseQuery -DatabaseConnection "DatabaseConnection" -SqlQuery "Select * From Table" -Log @($true,$objLogValue)
    Returns query result and log messages
  .Example
    New-DatabaseQuery -DatabaseConnection "DatabaseConnection" -SqlQuery "Select * From Table"
    Returns query result
  .Parameter DatabaseConnection
    Database connection object
  .Parameter SqlQuery
    SQL query
  .Parameter Log
    Array of constant log values.
  .Inputs
    [object]
    [string]
    [array]
  .OutPuts
    [array]
  .Notes
    NAME: New-DatabaseQuery
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170116
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Database Connection")]
    [ValidateNotNullOrEmpty()]
    [object]$DatabaseConnection,
    [Parameter(Mandatory=$true, Position=1, HelpMessage="SQL Query")]
    [ValidateNotNullOrEmpty()]
    [string]$SqlQuery,
    [Parameter(Mandatory=$false, Position=2, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t[New-DatabaseQuery]"}
  
  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "SQL query [$(If ($SqlQuery -like "*IDENTIFIED BY*") {$SqlQuery.Replace($SqlQuery.Split(" ")[-1],"'*****'")} Else {$SqlQuery})]"}

  [array]$arrReturnValue = @($false)

  Try {[string]$strIncludeDirectory = Split-Path $script:MyInvocation.MyCommand.Path} Catch {}
  [void][system.reflection.Assembly]::UnsafeLoadFrom("$strIncludeDirectory\MySql\MySql.Data.dll")
  [object]$objSqlCmd = New-Object MySql.Data.MySqlClient.MySqlCommand
  $objSqlCmd.CommandText = $SqlQuery
  $objSqlCmd.Connection = $DatabaseConnection

  If ($SqlQuery.ToUpper().StartsWith("SELECT "))
  {
    Try
    {
      [object]$objSqlAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter
      $objSqlAdapter.SelectCommand = $objSqlCmd
      [object]$objDataSet = New-Object System.Data.DataSet
      $objSqlAdapter.Fill($objDataSet) | Out-Null
      $arrReturnValue = @($true,$objDataSet.Tables[0])
    }
    Catch
    {
      If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage $_}
      $arrReturnValue = @($false,$_)
    }
  }
  Else
  {
    Try
    {
      [int]$intRowsAffected = $objSqlCmd.ExecuteNonQuery()
      $arrReturnValue = @($true,$intRowsAffected)
    }
    Catch
    {
      If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage $_}
      $arrReturnValue = @($false,$_)
    }
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t[New-DatabaseQuery]"}

  Return $arrReturnValue
}
