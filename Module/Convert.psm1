# Script Framework
# Name    : Convert.psm1
# Version : 0.1
# Date    : 2017-02-25
# Author  : Roeland van den Bosch
# Website : http://www.roelandvdbosch.nl


Function Convert-HexStrToUInt32
{
<#
  .Synopsis
    Convert HexStr > UInt32.
  .Description
    Convert a hexadecimal string
    to a 32 bits unsigned integer.
  .Example
    Convert-HexStrToUInt32 -HexString "1FA6"
    Returns an unsigned integer
    value ‭"8102‬".
  .Parameter HexString
    Hexadecimal string to convert.
  .Inputs
    [string]
  .Outputs
    [uInt32]
  .Notes
    NAME: Convert-HexStrToUInt32
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170225
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
  #Requires -Version 2.0
#> 
  Param(
	[parameter(Mandatory=$true, Position=0, HelpMessage="Hexadecimal string")]
    [String]$HexString
  )  
  Return [convert]::ToUInt32($HexString,16)
}

Function Convert-StrToHexStr
{
<#
  .Synopsis
    Convert Str > HexStr.
  .Description
    Convert a string to a
    hexadecimal string.
  .Example
    Convert-StrToHexStr -String "A sentence"
    Returns a string with the hexadecimal
    value "412073656E74656E6365".
  .Parameter String
    String to convert.
  .Inputs
    [string]
  .Outputs
    [string]
  .Notes
    NAME: Convert-StrToHexStr
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170225
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
  #Requires -Version 2.0
#> 
  Param (
	[parameter(Mandatory=$true, Position=0, HelpMessage="String")]
    [String]$String
  )
  Return [System.BitConverter]::ToString(([system.Text.Encoding]::UTF8).GetBytes($String)).Replace('-','')
}

Function Convert-UInt32ToBinStr
{
<#
  .Synopsis
    Convert UInt32 > BinStr.
  .Description
    Convert an unsigned integer
    to a binary string.
  .Example
    Convert-UInt32ToBinStr -Integer 8102
    Returns a string with the binary
    value "1111110100110".
  .Parameter Integer
    Integer to convert.
  .Inputs
    [uInt32]
  .Outputs
    [string]
  .Notes
    NAME: Convert-UInt32ToBinStr
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170225
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
  #Requires -Version 2.0
#> 
  Param (
	[parameter(Mandatory=$true, Position=0, HelpMessage="Integer")]
    [uInt32]$Integer
  )
  Return [convert]::ToString($Integer,2)
}
