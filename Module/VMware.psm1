# Script Framework
# Name    : IniFile.psm1
# Version : 1.3
# Date    : 2020-06-30
# Author  : Roeland van den Bosch
# Website : http://www.roelandvdbosch.nl

Function Connect-VMware
{
<#
  .Synopsis
    Create VMware Connection.
  .Description
    The Connect-VMware cmdlet creates a connection
    with a VMware environment.
  .Example
    Connect-VMware -Server "vmwareserver" -Username "username" -Password "password" -Log @($true,$objLogValue)
    Creates a connection with a VMware environment and log messages
  .Example
    Connect-VMware -Server "vmwareserver" -Username "username" -Password "password"
    Creates a connection with a VMware environment
  .Parameter Server
    Name of VMware server
  .Parameter Username
    Name of user
  .Parameter Password
    Password for user in string format
  .Parameter Log
    Create log messages
  .Inputs
    [string]
    [string]
    [string]
    [object]
  .OutPuts
    [boolean]
  .Notes
    NAME: Connect-VMware
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20200630
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Server")]
    [ValidateNotNullOrEmpty()]
    [string]$Server,
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Username")]
    [ValidateNotNullOrEmpty()]
    [string]$Username,
    [Parameter(Mandatory=$true, Position=2, HelpMessage="Password")]
    [ValidateNotNullOrEmpty()]
    [string]$Password,
    [Parameter(Mandatory=$false, Position=3, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t`t[Connect-VMware]"}

  [boolean]$blnReturnValue = $false

  Try
  {
    Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$False
    [object]$objConnectionInformation = Connect-VIServer -Server $Server -User $Username -Password $Password -ErrorAction Stop
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "INFO" -LogMessage "Connected to [$($objConnectionInformation.Name):$($objConnectionInformation.Port)]"}
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "INFO" -LogMessage "Username [$($objConnectionInformation.User)]"}
    $blnReturnValue = $objConnectionInformation.IsConnected
  }
  Catch
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage "$_"}
  }
  
  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t`t[Connect-VMware]"}
  Return $blnReturnValue
}

Function Disconnect-VMware
{
<#
  .Synopsis
    End VMware Connection.
  .Description
    The Disconnect-VMware cmdlet ends a connection
    with a VMware environment.
  .Example
    Disconnect-VMware -Server "vmwareserver" -Log @($true,$objLogValue)
    Ends a connection with a VMware environment and log messages
  .Example
    Disconnect-VMware -Server "vmwareserver"
    Ends a connection with a VMware environment
  .Parameter Server
    Name of VMware server
  .Parameter Log
    Create log messages
  .Inputs
    [string]
    [object]
  .OutPuts
    [boolean]
  .Notes
    NAME: Disconnect-VMware
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20200630
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Server")]
    [ValidateNotNullOrEmpty()]
    [string]$Server,
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  [boolean]$blnReturnValue = $false

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t`t[Disconnect-VMware]"}

  Try
  {
     Disconnect-VIServer -Server $Server -Confirm:$false -Force -ErrorAction Stop
     $blnReturnValue = $true
  }
  Catch
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage "$_"}
  }
  
  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t`t[Disconnect-VMware]"}

  Return $blnReturnValue
}

Function Get-VMwareFolder
{
<#
  .Synopsis
    Check if folder exists in VMware.
  .Description
    The Get-VMwareFolder cmdlet checks if
    a folder is present in a VMware environment.
  .Example
    Get-VMwareFolder -Server "vmwareserver" -Folder "folder" -Root "rootfolder" -Log @($true,$objLogValue)
    Checks if a folder is present in  a VMware 
    environment and log messages
  .Example
    Get-VMwareFolder -Server "vmwareserver" -Folder "folder" -Root "rootfolder"
    Checks if a folder is present in  a VMware 
    environment
  .Parameter Server
    Name of VMware server
  .Parameter Folder
    Name of folder
  .Parameter Root
    Root folder
  .Parameter Log
    Create log messages
  .Inputs
    [string]
    [string]
    [string]
    [object]
  .OutPuts
    [object]
  .Notes
    NAME: Get-VMwareFolder
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20200630
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Server")]
    [ValidateNotNullOrEmpty()]
    [string]$Server,
    [Parameter(Mandatory=$true, Position=1, HelpMessage="Folder")]
    [ValidateNotNullOrEmpty()]
    [string]$Folder,
    [Parameter(Mandatory=$false, Position=2, HelpMessage="Root")]
    [ValidateNotNullOrEmpty()]
    [string]$Root = '-',
    [Parameter(Mandatory=$false, Position=3, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t`t[Get-VMwareFolder]"}

  [boolean]$blnFolderFound = $false

  Try
  {
     If ($Root -eq '-')
     {
       [object]$objReturnValue = Get-Folder -Type VM `
                                            -Server $Server `
                                            -Name $Folder `
                                            -ErrorAction Stop
     }
     Else
     {
       [object]$objReturnValue = Get-Folder -Type VM `
                                            -Server $Server `
                                            -Name $Folder `
                                            -Location $Root `
                                            -ErrorAction Stop
     }
     $blnFolderFound = $true
  }
  Catch
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage "$_"}
  }
  
  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t`t[Disconnect-VMware]"}
 
  Return @($blnFolderFound,$objReturnValue)
}

Function Get-VMwareGuestServers
{
<#
  .Synopsis
    List VMware guest servers.
  .Description
    The Get-VMwareGuestServers cmdlet lists the
    guest servers in a VMware environment.
  .Example
    Get-VMwareGuestServers -Host "vmwareserver" -Name "guestserver" -Location "vmwarefolder" -Log @($true,$objLogValue)
    List all VMware guest servers "guestserver" in 
    VMware folder "vmwarefolder" and log messages
  .Example
    Get-VMwareGuestServers -Host "vmwareserver" -Name "guestserver" -Log @($true,$objLogValue)
    List all VMware guest servers "guestserver" 
    and log messages
  .Example
    Get-VMwareGuestServers -Host "vmwareserver" -Location "vmwarefolder" -Log @($true,$objLogValue)
    List all VMware guest servers in VMware 
    folder "vmwarefolder" and log messages
  .Example
    Get-VMwareGuestServers -Host "vmwareserver" -Name "guestserver" -Location "vmwarefolder"
    List all VMware guest servers "guestserver" in 
    VMware folder "vmwarefolder"
  .Parameter Host
    Name of VMware host server
  .Parameter Name
    Name of VMware guest server
  .Parameter Location
    Name of VMware folder
  .Parameter Log
    Create log messages
  .Inputs
    [string]
    [string]
    [string]
    [object]
  .OutPuts
    [array]
  .Notes
    NAME: Get-VMwareGuestServers
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20200630
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Host")]
    [ValidateNotNullOrEmpty()]
    [string]$Host,
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Name")]
    [ValidateNotNullOrEmpty()]
    [string]$Name,
    [Parameter(Mandatory=$false, Position=2, HelpMessage="Location")]
    [ValidateNotNullOrEmpty()]
    [string]$Location,
    [Parameter(Mandatory=$false, Position=3, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t`t[Get-VMwareGuestServers]"}

  [boolean]$blnServerFound = $false

  If ($Name -ne '' -or $Location -ne '')
  {
    Try
    {
       If ($Name -ne '' -and $Location -ne '') {[array]$arrReturnValue = Get-VM -Server $Host -Name $Name -Location $Location -ErrorAction Stop}
       ElseIf ($Name -ne '') {[array]$arrReturnValue = Get-VM -Server $Host -Name $Name -ErrorAction Stop}
       ElseIf ($Location -ne '') {[array]$arrReturnValue = Get-VM -Server $Host -Location $Location -ErrorAction Stop}
       If ($arrReturnValue.Count -ge 1) {$blnServerFound = $true}
    }
    Catch
    {
      If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage "$_"}
    }
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t`t[Get-VMwareGuestServers]"}
 
  Return @($blnServerFound,$arrReturnValue)
}

Function Start-VMwareGuestServers
{
<#
  .Synopsis
    Start VMware guest servers.
  .Description
    The Start-VMwareGuestServers cmdlet starts
    guest servers in a VMware environment.
  .Example
    Start-VMwareGuestServers -Host "vmwareserver" -Servers "guestservers" -Location "vmwarefolder" -Log @($true,$objLogValue)
    Start all VMware guest servers "guestservers" in 
    VMware folder "vmwarefolder" and log messages
  .Example
    Start-VMwareGuestServers -Host "vmwareserver" -Servers "guestservers" -Location "vmwarefolder"
    Start all VMware guest servers "guestservers" in 
    VMware folder "vmwarefolder"
  .Parameter Host
    Name of VMware host server
  .Parameter Servers
    Name of VMware guest servers
  .Parameter Location
    Name of VMware folder
  .Parameter Log
    Create log messages
  .Inputs
    [string]
    [object]
    [string]
    [object]
  .OutPuts
    [boolean]
  .Notes
    NAME: Start-VMwareGuestServers
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20200630
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Host")]
    [ValidateNotNullOrEmpty()]
    [string]$Host,
    [Parameter(Mandatory=$true, Position=1, HelpMessage="Servers")]
    [ValidateNotNullOrEmpty()]
    [object]$Servers,
    [Parameter(Mandatory=$true, Position=2, HelpMessage="Location")]
    [ValidateNotNullOrEmpty()]
    [string]$Location,
    [Parameter(Mandatory=$false, Position=3, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t`t[Start-VMwareGuestServers]"}

  [boolean]$blnServersRunning = $false
  Try
  {
    [System.Collections.ArrayList]$arrTaskList = @()
    Foreach ($Server in $Servers)
    {
      If ($Server.PowerState -eq "PoweredOff")
      {
        If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "INFO" -LogMessage "Start VMware guest server [$($Server.Name)]"}
        If ($Server.Guest.ToolsVersion -ne '')
        {
          If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "INFO" -LogMessage "VMware tools installed [$($Server.Guest.ToolsVersion)]"}
        }
        Else
        {
          If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "INFO" -LogMessage "NO VMware tools installed"}
        }
        Start-VM -VM $Server -Server $Host -RunAsync -Confirm:$false | Out-Null
        $arrTaskList.Add($Server.Name)
      }
      Else
      {
        If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "INFO" -LogMessage "VMware guest server already running [$($Server.Name)]"}
      }
    }
    [int]$intTotalTasks = $arrTaskList.Count

    While ($arrTaskList.Count -ge 1)
    {
      Write-Progress -Activity "Monitoring boot process" -Status "Boot in progress for $($arrTaskList | Foreach-Object {"[$_]"})" -PercentComplete (100 * ($intTotalTasks - $arrTaskList.Count) / $intTotalTasks)
      Foreach ($Server in $Servers)
      {
        $objServer = (Get-VMwareGuestServers -Host $Host -Name $Server -Location $Location)[1]
        If ($objServer.Guest.ToolsVersion -ne '')
        {
          If ($objServer.Guest.State -eq 'Running') {$arrTaskList.Remove($objServer.Name)}
        }
        Else
        {
          If ($objServer.PowerState -eq 'PoweredOn') {$arrTaskList.Remove($objServer.Name)}
        }
      }
      Start-Sleep -Seconds 5
    }

    If ($intTotalTasks -gt 0) {Write-Progress -Activity "Monitoring boot process" -Status "Boot in progress for $($arrTaskList | Foreach-Object {"[$_]"})" -PercentComplete (100 * ($intTotalTasks - $arrTaskList.Count) / $intTotalTasks)}

    $blnServersRunning = $true
  }
  Catch
  {
      If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage "$_"}
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t`t[Start-VMwareGuestServers]"}
 
  Return $blnServersRunning
}

Function Stop-VMwareGuestServers
{
<#
  .Synopsis
    Stop VMware guest servers.
  .Description
    The Stop-VMwareGuestServers cmdlet stops
    guest servers in a VMware environment.
  .Example
    Stop-VMwareGuestServers -Host "vmwareserver" -Servers "guestservers" -Location "vmwarefolder" -Force $true -Log @($true,$objLogValue)
    Stop all VMware guest servers "guestservers" in 
    VMware folder "vmwarefolder" and log messages
  .Example
    Stop-VMwareGuestServers -Host "vmwareserver" -Servers "guestservers" -Location "vmwarefolder" -Log @($true,$objLogValue)
    Shutdown all VMware guest servers "guestservers" in 
    VMware folder "vmwarefolder" and log messages
  .Example
    Stop-VMwareGuestServers -Host "vmwareserver" -Servers "guestservers" -Location "vmwarefolder"
    Shutdown all VMware guest servers "guestservers" in 
    VMware folder "vmwarefolder"
  .Parameter Host
    Name of VMware host server
  .Parameter Servers
    Name of VMware guest servers
  .Parameter Location
    Name of VMware folder
  .Parameter Log
    Create log messages
  .Inputs
    [string]
    [array]
    [string]
    [object]
  .OutPuts
    [boolean]
  .Notes
    NAME: Stop-VMwareGuestServers
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20200630
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Host")]
    [ValidateNotNullOrEmpty()]
    [string]$Host,
    [Parameter(Mandatory=$true, Position=1, HelpMessage="Servers")]
    [ValidateNotNullOrEmpty()]
    [array]$Servers,
    [Parameter(Mandatory=$true, Position=2, HelpMessage="Location")]
    [ValidateNotNullOrEmpty()]
    [string]$Location,
    [Parameter(Mandatory=$false, Position=3, HelpMessage="Force shutdown")]
    [ValidateNotNullOrEmpty()]
    [boolean]$Force = $false,
    [Parameter(Mandatory=$false, Position=4, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t`t[Stop-VMwareGuestServers]"}

  [boolean]$blnServersStopped = $false

  Try
  {
    [System.Collections.ArrayList]$arrTaskList = @()
    Foreach ($Server in $Servers)
    {
      If ($Server.PowerState -eq "PoweredOn")
      {
        If ($Server.Guest.ToolsVersion -ne '' -and -not $Force)
        {
          If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "INFO" -LogMessage "Shutdown VMware guest server [$($Server.Name)]"}
          If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "INFO" -LogMessage "VMware tools installed [$($Server.Guest.ToolsVersion)]"}
          Stop-VMGuest -VM $Server -Server $Host -Confirm:$false | Out-Null
          $arrTaskList.Add($Server.Name)
        }
        Else
        {
          If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "INFO" -LogMessage "Stop VMware guest server [$($Server.Name)]"}
          If ($Force)
          {
            If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "INFO" -LogMessage "Forced stop"}
          }
          Else
          {
            If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "INFO" -LogMessage "NO VMware tools installed"}
          }
          Stop-VM -VM $Server -Server $Host -Confirm:$false -RunAsync | Out-Null
          $arrTaskList.Add($Server.Name)
        }
      }
      Else
      {
        If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "INFO" -LogMessage "VMware guest server already stopped [$($Server.Name)]"}
      }
    }
    [int]$intTotalTasks = $arrTaskList.Count

    While ($arrTaskList.Count -ge 1)
    {
      Write-Progress -Activity "Monitoring shutdown process" -Status "Shutdown in progress for $($arrTaskList | Foreach-Object {"[$_]"})" -PercentComplete (100 * ($intTotalTasks - $arrTaskList.Count) / $intTotalTasks)
      Foreach ($Server in $Servers)
      {
        If ((Get-VMwareGuestServers -Host $Host -Name $Server -Location $Location)[1].PowerState -eq 'PoweredOff') {$arrTaskList.Remove($Server.Name)}
      }
      Start-Sleep -Seconds 5
    }

    If ($intTotalTasks -gt 0) {Write-Progress -Activity "Monitoring shutdown process" -Status "Shutdown in progress for $($arrTaskList | Foreach-Object {"[$_]"})" -PercentComplete (100 * ($intTotalTasks - $arrTaskList.Count) / $intTotalTasks)}

    $blnServersStopped = $true
  }
  Catch
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage "$_"}
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t`t[Stop-VMwareGuestServers]"}
 
  Return $blnServersStopped
}

Function Remove-VMwareGuestServers
{
<#
  .Synopsis
    Remove VMware guest servers.
  .Description
    The Remove-VMwareGuestServers cmdlet removes
    guest servers in a VMware environment.
  .Example
    Remove-VMwareGuestServers -Host "vmwareserver" -Servers "guestservers" -Location "vmwarefolder" -Log @($true,$objLogValue)
    Remove all VMware guest servers "guestservers" in 
    VMware folder "vmwarefolder" and log messages
  .Example
    Remove-VMwareGuestServers -Host "vmwareserver" -Servers "guestservers" -Location "vmwarefolder"
    Remove all VMware guest servers "guestservers" in 
    VMware folder "vmwarefolder"
  .Parameter Host
    Name of VMware host server
  .Parameter Servers
    Name of VMware guest servers
  .Parameter Location
    Name of VMware folder
  .Parameter Log
    Create log messages
  .Inputs
    [string]
    [array]
    [string]
    [object]
  .OutPuts
    [boolean]
  .Notes
    NAME: Remove-VMwareGuestServers
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20200630
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Host")]
    [ValidateNotNullOrEmpty()]
    [string]$Host,
    [Parameter(Mandatory=$true, Position=1, HelpMessage="Servers")]
    [ValidateNotNullOrEmpty()]
    [array]$Servers,
    [Parameter(Mandatory=$true, Position=2, HelpMessage="Location")]
    [ValidateNotNullOrEmpty()]
    [string]$Location,
    [Parameter(Mandatory=$false, Position=3, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t`t[Remove-VMwareGuestServers]"}
  
  Try
  {
    [boolean]$blnServersStopped = $false
    If ($Servers.PowerState -contains 'PoweredOn')
    {
      If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "INFO" -LogMessage "VMware guest servers still running [$($Servers | Foreach-Object {If ($_.PowerState -eq 'PoweredOn') {$_.Name}})]"}
      $blnServersStopped = Stop-VMwareGuestServers -Host $Host `
                                                   -Servers $Servers `
                                                   -Location $Location `
                                                   -Force $true `
                                                   -Log $Log
    }
    Else
    {
      $blnServersStopped = $true
    }

    If ($blnServersStopped)
    {
      [object]$objTaskList = @{}
      Foreach ($Server in $Servers)
      {
        If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "INFO" -LogMessage "Remove VMware guest server [$($Server.Name)]"}
        $objTaskList[(Remove-VM -VM $Server -Server $Host -Confirm:$false -DeletePermanently -RunAsync -ErrorAction Stop).Id] = $Server.Name
      }
      [int]$intTotalTasks = $objTaskList.Count

      While ($objTaskList.Count -ge 1)
      {
        Write-Progress -Activity "Monitoring remove process" -Status "Remove in progress for $($objTaskList.Values | Foreach-Object {"[$_]"})" -PercentComplete (100 * ($intTotalTasks - $objTaskList.Count) / $intTotalTasks)
        Get-Task -Server $Host -Status Success | Foreach-Object {If ($objTaskList.Keys -contains $_.Id) {$objTaskList.Remove($_.Id)}}
        Start-Sleep -Seconds 5
      }

      Write-Progress -Activity "Monitoring remove process" -Status "Remove in progress for $($objTaskList.Values | Foreach-Object {"[$_]"})" -PercentComplete (100 * ($intTotalTasks - $objTaskList.Count) / $intTotalTasks)

      $blnServersRemoved = $true
    }
  }
  Catch
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage "$_"}
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t`t[Remove-VMwareGuestServers]"}
 
  Return $blnServersRemoved
}

Function Set-VMwareGuestServerMacAddress
{
<#
  .Synopsis
    Set VMware guest server MAC address.
  .Description
    The Set-VMwareGuestServerMacAddress cmdlet sets
    the MAC address for a guest servers in a VMware
    environment.
  .Example
    Set-VMwareGuestServerMacAddress -Host "vmwareserver" -Nic "nicobject" -Type "generated" -Log @($true,$objLogValue)
    Set MAC address for VMware guest server nic object "nicobject" with 
    type "generated" and log messages
  .Example
    Set-VMwareGuestServerMacAddress -Host "vmwareserver" -Nic "nicobject" -Type "manual" -Address "00:50:56:ab:00:10" -Log @($true,$objLogValue)
    Set MAC address for VMware guest server nic object "nicobject" with 
    type "manual", MAC address "00:50:56:ab:00:10" and log messages
  .Parameter Host
    Name of VMware host server
  .Parameter Nic
    Network adapter object
  .Parameter Type
    Type of MAC address
  .Parameter Address
    MAC address
  .Parameter Log
    Create log messages
  .Inputs
    [string]
    [object]
    [string]
    [string]
    [object]
  .OutPuts
    [array]
  .Notes
    NAME: Set-VMwareGuestServerMacAddress
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20200630
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Host")]
    [ValidateNotNullOrEmpty()]
    [string]$Host,
    [Parameter(Mandatory=$true, Position=1, HelpMessage="Network Adapter")]
    [ValidateNotNullOrEmpty()]
    [object]$Nic,
    [Parameter(Mandatory=$true, Position=2, HelpMessage="MAC Type (manual, generated or assigned)")]
    [ValidateNotNullOrEmpty()]
    [string]$Type,
    [Parameter(Mandatory=$false, Position=3, HelpMessage="MAC Address")]
    [ValidateNotNullOrEmpty()]
    [string]$Address,
    [Parameter(Mandatory=$false, Position=3, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t`t[Set-VMwareGuestServerMacAddress]"}

  Try
  {
    [boolean]$blnVMwareGuestServerMacAddressSet = $false

    [object]$objVMwareGuestServer = $Nic.Parent
    $objVirtualMachineConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
    $objVirtualDeviceConfigSpec = New-Object VMware.Vim.VirtualDeviceConfigSpec

    $objVirtualDeviceConfigSpec.Operation = 'edit'
    $objVMwareGuestServer.ExtensionData.Config.Hardware.Device |
      Where-Object {$_.DeviceInfo.Label -eq $Nic.Name} |
      ForEach-Object {$objVirtualDeviceConfigSpec.Device = $_
                      $objVirtualDeviceConfigSpec.Device.AddressType = $Type
                      $objVirtualDeviceConfigSpec.Device.MacAddress = $(If ($Address -ne '') {$Address} Else {$null})}
    $objVirtualMachineConfigSpec.DeviceChange += $objVirtualDeviceConfigSpec
    $objVMwareGuestServer.ExtensionData.ReconfigVM($objVirtualMachineConfigSpec)
    [object]$objNic = $objVMwareGuestServer |
                        Get-NetworkAdapter -Server $Host |
                        Where-Object {$_.Name -eq $Nic.Name}
    $blnVMwareGuestServerMacAddressSet = $true
  }
  Catch
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage "$_"}
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t`t[Set-VMwareGuestServerMacAddress]"}
 
  Return @($blnVMwareGuestServerMacAddressSet, $objNic)
}