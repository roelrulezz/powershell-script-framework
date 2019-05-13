# Script Framework
# Name    : IniFile.psm1
# Version : 0.2
# Date    : 2018-11-05
# Author  : Roeland van den Bosch
# Website : http://www.roelandvdbosch.nl

Function Read-IniFile 
{
<#
  .Synopsis
    Read ini file.
  .Description
    The Read-IniFile cmdlet reads and
    returns the content of an ini file.
  .Example
    Read-IniFile -FilePath 'C:\Windows\win.ini' -Log @($true,$objLogValue)
    Gets the content of 'C:\Windows\win.ini'
    and log messages.
  .Example
    Read-IniFile -FilePath 'C:\Windows\win.ini'
    Gets the content of 'C:\Windows\win.ini'.
  .Parameter FilePath
    File path of the ini file.
  .Parameter Log
    Array of constant log values.
  .Inputs
    [string]
    [object]
  .Outputs
    [object]
  .Notes
    NAME: Read-IniFile
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20181105
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
  #Requires -Version 2.0
#> 
  Param (
	[parameter(Mandatory=$true, Position=0, HelpMessage="File path")]
    [String]$FilePath,
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t`t[Read-IniFile]"}

  [object]$objIni = New-Object System.Collections.Specialized.OrderedDictionary
  [object]$objCurrentSection = New-Object System.Collections.Specialized.OrderedDictionary
  [string]$strCurrentSectionName = "default"
  
  [int]$intCommentCounter = 1

  If (Test-Path $FilePath)
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "INFO" -LogMessage "Ini file found:`t`t[$FilePath]"}
    
    Foreach ($Line in $(Get-Content -Path $FilePath))
    {
      Switch -Regex ($Line)
      {
	    "^\[(?<section>.*)\]$" # Section
	    {
          $objIni.Add($strCurrentSectionName, $objCurrentSection)
	      $strCurrentSectionName = $Matches['Section']
	      $objCurrentSection = New-Object System.Collections.Specialized.OrderedDictionary
          $intCommentCounter = 1
          Break
	    }
	    "^\s*\;\s*(?<value>.*)" # Comment
	    {
          $objCurrentSection.Add("Comment_$intCommentCounter", $Matches['Value'])
          $intCommentCounter++
          Break  
	    }
        "(?<key>.+?)\s*=\s*(?<value>.*)" # Key  
	    {
	      # Add to current section Hash Set
	      $objCurrentSection.Add($Matches['Key'], $Matches['Value'])
          Break
	    }
	    "^$" # Blank line
        {
          # Ignore blank line
          Break
	    }
	    Default
        {
	      Throw "Unidentified: $_" # Should not happen
	    }
      }
    }
  }
  Else
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "WARNING" -LogMessage "Ini file NOT found:`t`t[$FilePath]"}
  }

  If ($objIni.Keys -notcontains $strCurrentSectionName)
  {
    $objIni.Add($strCurrentSectionName, $objCurrentSection)
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t`t[Read-IniFile]"}

  Return $objIni
}

Function Write-IniFile
{
<#
  .Synopsis
    Write ini file.
  .Description
    The Write-IniFile cmdlet writes an ini file.
  .Example
    Write-IniFile -Ini "IniContent" -FilePath "C:\Windows\win.ini" -Log @($true,$objLogValue)
    Sets the content from "IniContent" to "C:\Windows\win.ini"
    and log messages.
  .Example
    Write-IniFile -Ini "IniContent" -FilePath "C:\Windows\win.ini"
    Sets the content from "IniContent" to "C:\Windows\win.ini".
  .Parameter Ini
    Content of the ini file.
  .Parameter FilePath
    File path of the ini file.
  .Parameter LogValue
    Object of constant log values.
  .Inputs
    [object]
    [string]
    [object]
  .Outputs
    [string]
  .Notes
    NAME: Write-IniFile
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20181105
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
  #Requires -Version 2.0
#> 
  Param (
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Ini content")]
    [System.Collections.Specialized.OrderedDictionary]$Ini,
	[parameter(Mandatory=$false, Position=1, HelpMessage="File path")]
    [String]$FilePath,
    [Parameter(Mandatory=$false, Position=2, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t`t[Write-IniFile]"}
	
  [string]$strOutput = ""
  ForEach ($objSection in $Ini.GetEnumerator())
  {
    If ($objSection.Name -ne "default") 
    { 
	  # Insert a blank line after a section
	  [string]$strPrefix = @{$true="";$false="`r`n"}[[String]::IsNullOrWhiteSpace($strOutput)]
	  $strOutput += "$strPrefix[$($objSection.Name)]`r`n" 
	}
	ForEach ($objEntry in $objSection.Value.GetEnumerator())
	{
	  [string]$strPrefix = @{$true=";";$false="$($objEntry.Name)="}[$objEntry.Name -like 'Comment_*']
	  $strOutput += "$strPrefix$($objEntry.Value)`r`n"
	}
  }

  $strOutput = $strOutput.TrimEnd("`r`n")
  If ([String]::IsNullOrEmpty($FilePath))
  {
	If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t`t[Write-IniFile]"}
    Return $strOutput
  }
  Else
  {
	$strOutput | Out-File -FilePath $FilePath -Encoding:ASCII
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t`t[Write-IniFile]"}
}
