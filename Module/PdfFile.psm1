# Script Framework
# Name    : Log.psm1
# Version : 0.1
# Date    : 2017-03-05
# Author  : Roeland van den Bosch
# Website : http://www.roelandvdbosch.nl

Function New-Pdf
{
<#
  .Synopsis
    Create PDF document.
  .Description
    The New-Pdf cmdlet creates a PDF document.
  .Example
    New-Pdf -File "FullFilePath" -TopMargin 20 -BottomMargin 20 -LeftMargin 20 -RightMargin 20 -Author "Roeland van den Bosch" -Log @($true,$objLogValue)
    Returns a PDF ([iTextSharp.text.Document]) object and log messages
  .Example
    New-Pdf -File "FullFilePath" -TopMargin 20 -BottomMargin 20 -LeftMargin 20 -RightMargin 20 -Author "Roeland van den Bosch"
    Returns a PDF ([iTextSharp.text.Document]) object
  .Parameter File
    PDF file name.
  .Parameter TopMargin
    Top margin PDF document.
  .Parameter BottomMargin
    Bottom margin PDF document.
  .Parameter LeftMargin
    Left margin PDF document.
  .Parameter RightMargin
    Right margin PDF document.
  .Parameter Author
    Author PDF document.
  .Parameter Log
    Array of constant log values.
  .Inputs
    [string]
    [int32]
    [int32]
    [int32]
    [int32]
    [string]
    [array]
  .OutPuts
    [iTextSharp.text.Document]
  .Notes
    NAME: New-Pdf
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170304
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Full Filepath")]
    [ValidateNotNullOrEmpty()]
    [string]$File,
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Top Margin")]
    [ValidateNotNullOrEmpty()]
    [int32]$TopMargin = 20,
    [Parameter(Mandatory=$false, Position=2, HelpMessage="Bottom Margin")]
    [ValidateNotNullOrEmpty()]
    [int32]$BottomMargin = 20,
    [Parameter(Mandatory=$false, Position=3, HelpMessage="Left Margin")]
    [ValidateNotNullOrEmpty()]
    [int32]$LeftMargin = 20,
    [Parameter(Mandatory=$false, Position=4, HelpMessage="Right Margin")]
    [ValidateNotNullOrEmpty()]
    [int32]$RightMargin = 20,
    [Parameter(Mandatory=$false, Position=5, HelpMessage="Author")]
    [ValidateNotNullOrEmpty()]
    [string]$Author = "Unkown",
    [Parameter(Mandatory=$false, Position=6, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t[New-Pdf]"}

  [array]$arrReturnValue = @($false)

  Try {[string]$strIncludeDirectory = Split-Path $script:MyInvocation.MyCommand.Path} Catch {}
  [void][system.reflection.Assembly]::UnsafeLoadFrom("$strIncludeDirectory\iTextSharp\itextsharp.dll")

  Try
  {
    [iTextSharp.text.Document]$objPdfDocument = New-Object iTextSharp.text.Document
    $objPdfDocument.SetPageSize([iTextSharp.text.PageSize]::A4) | Out-Null
    $objPdfDocument.SetMargins($LeftMargin, $RightMargin, $TopMargin, $BottomMargin) | Out-Null
    [void][iTextSharp.text.pdf.PdfWriter]::GetInstance($objPdfDocument, [System.IO.File]::Create($File))
    $objPdfDocument.AddAuthor($Author) | Out-Null
    $arrReturnValue = @($true,$objPdfDocument)
  }
  Catch
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage $_}
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t[New-Pdf]"}

  Return [array]$arrReturnValue
}

Function New-PdfTitle
{
<#
  .Synopsis
    Creates a PDF title.
  .Description
    The New-PdfTitle cmdlet creates a title for a PDF document.
  .Example
    New-PdfTitle -Text "Dit is een test titel" -Centered -FontName "Arial" -FontSize 12 -FontColor "Black" -Log @($true,$objLogValue)
    Adds a title to a PDF document and log messages
  .Example
    New-PdfTitle -Text "Dit is een test titel" -Centered -FontName "Arial" -FontSize 12 -FontColor "Black"
    Adds a title to a PDF document
  .Parameter Text
    Title to add.
  .Parameter Centered
    Title is centered
  .Parameter FontName
    Font name of the text.
  .Parameter FontSize
    Font size of the text.
  .Parameter FontColor
    Color of the text.
  .Parameter Log
    Array of constant log values.
  .Inputs
    [string]
    [switch]
    [string]
    [int32]
    [string]
    [array]
  .Outputs
    [object]
  .Notes
    NAME: New-PdfTitle
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170304
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Text")]
    [ValidateNotNullOrEmpty()]
    [string]$Text,
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Centered")]
    [ValidateNotNullOrEmpty()]
    [switch]$Centered,
    [Parameter(Mandatory=$false, Position=2, HelpMessage="Font Name")]
    [ValidateNotNullOrEmpty()]
    [string]$FontName = "Arial",
    [Parameter(Mandatory=$false, Position=3, HelpMessage="Font Size")]
    [ValidateNotNullOrEmpty()]
    [int32]$FontSize = 12,
    [Parameter(Mandatory=$false, Position=4, HelpMessage="Font Color")]
    [ValidateNotNullOrEmpty()]
    [string]$FontColor = "Black",
    [Parameter(Mandatory=$false, Position=5, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t[New-PdfTitle]"}

  Try {[string]$strIncludeDirectory = Split-Path $script:MyInvocation.MyCommand.Path} Catch {}
  [void][system.reflection.Assembly]::UnsafeLoadFrom("$strIncludeDirectory\iTextSharp\itextsharp.dll")

  Try
  {
    [object]$objPdfParagraph = New-Object iTextSharp.text.paragraph
    $objPdfParagraph.Font = [iTextSharp.text.FontFactory]::GetFont($FontName, $FontSize, [iTextSharp.text.Font]::BOLD, [iTextSharp.text.BaseColor]::$FontColor)
    if($Centered) { $objPdfParagraph.Alignment = [iTextSharp.text.Element]::ALIGN_CENTER }
    $objPdfParagraph.SpacingBefore = 5
    $objPdfParagraph.SpacingAfter = 5
    $objPdfParagraph.Add($Text) | Out-Null
  }
  Catch
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage $_}
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t[New-PdfTitle]"}

  Return [object]$objPdfParagraph
}

Function New-PdfText
{
<#
  .Synopsis
    Creates a PDF text.
  .Description
    The New-PdfText cmdlet creates a PDF text.
  .Example
    New-PdfText -Text "Dit is een test tekstje" -FontName "Arial" -FontSize 12 -FontColor "Black" -Log @($true,$objLogValue)
    Adds text to a PDF file and log messages
  .Example
    New-PdfText -Text "Dit is een test tekstje" -FontName "Arial" -FontSize 12 -FontColor "Black"
    Adds text to a PDF file
  .Parameter Text
    Text to add.
  .Parameter FontName
    Font name of the text.
  .Parameter FontSize
    Font size of the text.
  .Parameter FontColor
    Color of the text.
  .Parameter Log
    Array of constant log values.
  .Inputs
    [string]
    [string]
    [int32]
    [string]
    [array]
  .Outputs
    [object]
  .Notes
    NAME: New-PdfText
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170304
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Text")]
    [ValidateNotNullOrEmpty()]
    [string]$Text,
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Font Name")]
    [ValidateNotNullOrEmpty()]
    [string]$FontName = "Arial",
    [Parameter(Mandatory=$false, Position=2, HelpMessage="Font Size")]
    [ValidateNotNullOrEmpty()]
    [int32]$FontSize = 12,
    [Parameter(Mandatory=$false, Position=3, HelpMessage="Font Color")]
    [ValidateNotNullOrEmpty()]
    [string]$FontColor = "Black",
    [Parameter(Mandatory=$false, Position=4, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
    )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t[New-PdfText]"}

  Try {[string]$strIncludeDirectory = Split-Path $script:MyInvocation.MyCommand.Path} Catch {}
  [void][system.reflection.Assembly]::UnsafeLoadFrom("$strIncludeDirectory\iTextSharp\itextsharp.dll")

  Try
  {
    [object]$objPdfParagraph = New-Object iTextSharp.text.Paragraph
    $objPdfParagraph.Font = [iTextSharp.text.FontFactory]::GetFont($FontName, $FontSize, [iTextSharp.text.Font]::NORMAL, [iTextSharp.text.BaseColor]::$FontColor)
    $objPdfParagraph.SpacingBefore = 2
    $objPdfParagraph.SpacingAfter = 2
    $objPdfParagraph.Add($Text) | Out-Null
  }
  Catch
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage $_}
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t[New-PdfTitle]"}

  Return [object]$objPdfParagraph
}

Function New-PdfPhrase
{
<#
  .Synopsis
    Creates a PDF phrase.
  .Description
    The New-PdfPhrase cmdlet creates a PDF phrase.
  .Example
    New-PdfPhrase -Text "Dit is een test tekstje" -FontName "Arial" -FontSize 12 -FontColor "Black" -Log @($true,$objLogValue)
    Adds text to a PDF file and log messages
  .Example
    New-PdfPhrase -Text "Dit is een test tekstje" -FontName "Arial" -FontSize 12 -FontColor "Black"
    Adds text to a PDF file
  .Parameter Text
    Text to add.
  .Parameter FontName
    Font name of the text.
  .Parameter FontSize
    Font size of the text.
  .Parameter FontColor
    Color of the text.
  .Parameter FontDecoration
    Decoration of the text.
  .Parameter Log
    Array of constant log values.
  .Inputs
    [string]
    [string]
    [int32]
    [string]
    [string]
    [array]
  .Outputs
    [object]
  .Notes
    NAME: New-PdfPhrase
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170305
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Text")]
    [ValidateNotNullOrEmpty()]
    [string]$Text,
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Font Name")]
    [ValidateNotNullOrEmpty()]
    [string]$FontName = "Arial",
    [Parameter(Mandatory=$false, Position=2, HelpMessage="Font Size")]
    [ValidateNotNullOrEmpty()]
    [int32]$FontSize = 10,
    [Parameter(Mandatory=$false, Position=3, HelpMessage="Font Color")]
    [ValidateNotNullOrEmpty()]
    [string]$FontColor = "Black",
    [Parameter(Mandatory=$false, Position=4, HelpMessage="Font Decoration")]
    [ValidateNotNullOrEmpty()]
    [string]$FontDecoration = "None",
    [Parameter(Mandatory=$false, Position=5, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t[New-PdfPhrase]"}

  Try {[string]$strIncludeDirectory = Split-Path $script:MyInvocation.MyCommand.Path} Catch {}
  [void][system.reflection.Assembly]::UnsafeLoadFrom("$strIncludeDirectory\iTextSharp\itextsharp.dll")

  Try
  {
    $objPhrase = New-Object iTextSharp.text.Phrase
    $objPhrase.Font = [iTextSharp.text.FontFactory]::GetFont($FontName, $FontSize, [iTextSharp.text.Font]::$FontDecoration, [iTextSharp.text.BaseColor]::$FontColor)
    $objPhrase.Add($Text) | Out-Null
  }
  Catch
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage $_}
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t[New-PdfPhrase]"}

  Return [object]$objPhrase
}

Function New-PdfImage
{
<#
  .Synopsis
    Creates a PDF image.
  .Description
    The New-PdfImage cmdlet creates a PDF image.
  .Example
    New-PdfImage -Document "objPdf" -File "Image file fullpath" -Scale 100 -Log @($true,$objLogValue)
    Creates a PDF image and log messages
  .Example
    New-PdfImage -Document "objPdf" -File "Image file fullpath" -Scale 100
    Creates a PDF image
  .Parameter File
    Image to add.
  .Parameter Scale
    Image scale.
  .Parameter Alignment
    Image alignment.
  .Parameter Log
    Array of constant log values.
  .Inputs
    [string]
    [int32]
    [int32]
    [array]
  .Outputs
    [object]
  .Notes
    NAME: New-PdfImage
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170305
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="File")]
    [ValidateNotNullOrEmpty()]
    [string]$File,
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Scale")]
    [ValidateNotNullOrEmpty()]
    [int]$Scale = 100,
    [Parameter(Mandatory=$false, Position=2, HelpMessage="Alignment")]
    [ValidateNotNullOrEmpty()]
    [int]$Alignment = 0, #Left
    [Parameter(Mandatory=$false, Position=3, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t[New-PdfImage]"}

  Try {[string]$strIncludeDirectory = Split-Path $script:MyInvocation.MyCommand.Path} Catch {}
  [void][system.reflection.Assembly]::UnsafeLoadFrom("$strIncludeDirectory\iTextSharp\itextsharp.dll")

  Try
  {
    [iTextSharp.text.Image]$objPdfImage = [iTextSharp.text.Image]::GetInstance($File)
    $objPdfImage.ScalePercent($Scale) | Out-Null
    $objPdfImage.Alignment = $Alignment
  }
  Catch
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage $_}
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t[New-PdfImage]"}
  
  Return [object]$objPdfImage
}

Function New-PdfTable
{
<#
  .Synopsis
    Creates a PDF table.
  .Description
    The New-PdfTable creates a PDF table.
  .Example
    New-PdfTable -NumberOfColumns 2 -Log @($true,$objLogValue)
    Creates a PDF table and log messages
  .Example
    New-PdfTable -NumberOfColumns 2
    Creates a PDF table
  .Parameter NumberOfColumns
    Number of columns.
  .Parameter TotalWidth
    Total width of the table.
  .Parameter LockedWidth
    Lock table width.
  .Parameter ColumnWidths
    Set widths of the columns.
  .Parameter HorizontalAlignment
    Horizontal alignment of the table.
  .Parameter Log
    Array of constant log values.
  .Inputs
    [int32]
    [int32]
    [boolean]
    [aray]
    [int32]
    [array]
  .Outputs
    [object]
  .Notes
    NAME: New-PdfTable
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170305
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Number Of Columns")]
    [ValidateNotNullOrEmpty()]
    [int]$NumberOfColumns,
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Total Width")]
    [ValidateNotNullOrEmpty()]
    [int]$TotalWidth = 0,
    [Parameter(Mandatory=$false, Position=2, HelpMessage="LockedWidth")]
    [ValidateNotNullOrEmpty()]
    [boolean]$LockedWidth = $false,
    [Parameter(Mandatory=$false, Position=3, HelpMessage="Column Widths")]
    [ValidateNotNullOrEmpty()]
    [array]$ColumnWidths,
    [Parameter(Mandatory=$false, Position=4, HelpMessage="Horizontal Alignment")]
    [ValidateNotNullOrEmpty()]
    [int]$HorizontalAlignment = 0, #left
    [Parameter(Mandatory=$false, Position=5, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )
 
  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t[New-PdfTable]"}

  Try {[string]$strIncludeDirectory = Split-Path $script:MyInvocation.MyCommand.Path} Catch {}
  [void][system.reflection.Assembly]::UnsafeLoadFrom("$strIncludeDirectory\iTextSharp\itextsharp.dll")

  Try
  {
    $objPdfTable = New-Object iTextSharp.text.pdf.PDFPTable($NumberOfColumns)
    $objPdfTable.TotalWidth = $TotalWidth
    $objPdfTable.LockedWidth = $LockedWidth
    If ($ColumnWidths -ne $null) {$objPdfTable.SetWidths($ColumnWidths)}
    $objPdfTable.HorizontalAlignment = $HorizontalAlignment
  }
  Catch
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage $_}
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t[New-PdfTable]"}
  
  Return [object]$objPdfTable
}

Function New-PdfCell
{
<#
  .Synopsis
    Creates a PDF cell.
  .Description
    The New-PdfCell creates a PDF cell.
  .Example
    New-PdfCell -Log @($true,$objLogValue)
    Creates a PDF cell and log messages
  .Example
    New-PdfCell
    Creates a PDF cell
  .Parameter ColSpan
    Column span.
  .Parameter RowSpan
    Row span.
  .Parameter HorizontalAlignment
    Horizontal alignment.
  .Parameter VerticalAlignment
    Vertical alignment.
  .Parameter BorderColor
    Border color.
  .Parameter BorderWidth
    Border width.
  .Parameter Padding
    Cell padding.
  .Parameter Log
    Array of constant log values.
  .Inputs
    [int32]
    [int32]
    [int32]
    [int32]
    [string]
    [int32]
    [int32]
    [array]
  .Outputs
    [object]
  .Notes
    NAME: New-PdfCell
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170305
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$false, Position=0, HelpMessage="Col Span")]
    [ValidateNotNullOrEmpty()]
    [int]$ColSpan = 1,
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Row Span")]
    [ValidateNotNullOrEmpty()]
    [int]$RowSpan = 1,
    [Parameter(Mandatory=$false, Position=2, HelpMessage="Horizontal Alignment")]
    [ValidateNotNullOrEmpty()]
    [int]$HorizontalAlignment = 0, #Left
    [Parameter(Mandatory=$false, Position=3, HelpMessage="Vertical Alignment")]
    [ValidateNotNullOrEmpty()]
    [int]$VerticalAlignment = 4, #Center
    [Parameter(Mandatory=$false, Position=4, HelpMessage="Border Color")]
    [ValidateNotNullOrEmpty()]
    [string]$BorderColor = "Black",
    [Parameter(Mandatory=$false, Position=5, HelpMessage="Border Width")]
    [ValidateNotNullOrEmpty()]
    [int]$BorderWidth = 0,
    [Parameter(Mandatory=$false, Position=6, HelpMessage="Padding")]
    [ValidateNotNullOrEmpty()]
    [int]$Padding = 3,
    [Parameter(Mandatory=$false, Position=7, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t[New-PdfCell]"}

  Try {[string]$strIncludeDirectory = Split-Path $script:MyInvocation.MyCommand.Path} Catch {}
  [void][system.reflection.Assembly]::UnsafeLoadFrom("$strIncludeDirectory\iTextSharp\itextsharp.dll")

  Try
  {
    $objPdfCell = New-Object iTextSharp.text.pdf.PdfPCell
    $objPdfCell.Colspan = $ColSpan
    $objPdfCell.Rowspan = $RowSpan
    $objPdfCell.HorizontalAlignment = $HorizontalAlignment
    $objPdfCell.VerticalAlignment = $VerticalAlignment
    $objPdfCell.BorderColor = [iTextSharp.text.BaseColor]::$BorderColor
    $objPdfCell.BorderWidth = $BorderWidth
    $objPdfCell.Padding = $Padding
  }
  Catch
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage $_}
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t[New-PdfCell]"}
  
  Return [object]$objPdfCell
}

Function New-PdfWhiteCell
{
<#
  .Synopsis
    Creates a white PDF cell.
  .Description
    The New-PdfWhiteCell creates a white PDF cell.
  .Example
    New-PdfWhiteCell -Log @($true,$objLogValue)
    Creates a white PDF cell and log messages
  .Example
    New-PdfWhiteCell
    Creates a white PDF cell
  .Parameter ColSpan
    Column span.
  .Parameter RowSpan
    Row span.
  .Parameter BorderWidthTop
    Border width top.
  .Parameter BorderWidthLeft
    Border width left.
  .Parameter BorderWidthRight
    Border width right.
  .Parameter BorderWidthBottom
    Border width bottom.
  .Parameter BorderColorTop
    Border color top.
  .Parameter BorderColorLeft
    Border color left.
  .Parameter BorderColorRight
    Border color right.
  .Parameter BorderColorBottom
    Border color bottom.
  .Parameter Height
    Row Height.
  .Parameter Log
    Create log messages
  .Inputs
    [int32]
    [int32]
    [int32]
    [int32]
    [int32]
    [int32]
    [string]
    [string]
    [string]
    [string]
    [int32]
    [array]
  .Outputs
    [object]
  .Notes
    NAME: New-PdfWhiteCell
    AUTHOR: Roeland van den Bosch
    LASTEDIT: 20170305
    KEYWORDS:
  .Link
     http://www.roelandvdbosch.nl
 #Requires -Version 2.0
#> 
  Param(
    [Parameter(Mandatory=$false, Position=0, HelpMessage="Col Span")]
    [ValidateNotNullOrEmpty()]
    [int]$ColSpan = 1,
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Row Span")]
    [ValidateNotNullOrEmpty()]
    [int]$RowSpan = 1,
    [Parameter(Mandatory=$false, Position=2, HelpMessage="BorderWidthTop")]
    [ValidateNotNullOrEmpty()]
    [int]$BorderWidthTop = 0,
    [Parameter(Mandatory=$false, Position=3, HelpMessage="BorderWidthLeft")]
    [ValidateNotNullOrEmpty()]
    [int]$BorderWidthLeft = 0,
    [Parameter(Mandatory=$false, Position=4, HelpMessage="BorderWidthRight")]
    [ValidateNotNullOrEmpty()]
    [int]$BorderWidthRight = 0,
    [Parameter(Mandatory=$false, Position=5, HelpMessage="BorderWidthBottom")]
    [ValidateNotNullOrEmpty()]
    [int]$BorderWidthBottom = 0,
    [Parameter(Mandatory=$false, Position=6, HelpMessage="BorderColorTop")]
    [ValidateNotNullOrEmpty()]
    [string]$BorderColorTop = "Black",
    [Parameter(Mandatory=$false, Position=7, HelpMessage="BorderColorLeft")]
    [ValidateNotNullOrEmpty()]
    [string]$BorderColorLeft = "Black",
    [Parameter(Mandatory=$false, Position=8, HelpMessage="BorderColorRight")]
    [ValidateNotNullOrEmpty()]
    [string]$BorderColorRight = "Black",
    [Parameter(Mandatory=$false, Position=9, HelpMessage="BorderColorBottom")]
    [ValidateNotNullOrEmpty()]
    [string]$BorderColorBottom = "Black",
    [Parameter(Mandatory=$false, Position=10, HelpMessage="Height")]
    [ValidateNotNullOrEmpty()]
    [int]$Height = 12,
    [Parameter(Mandatory=$false, Position=11, HelpMessage="Create log")]
    [ValidateNotNullOrEmpty()]
    [array]$Log = @($false)
  )

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "Start function:`t[New-PdfWhiteCell]"}

  Try {[string]$strIncludeDirectory = Split-Path $script:MyInvocation.MyCommand.Path} Catch {}
  [void][system.reflection.Assembly]::UnsafeLoadFrom("$strIncludeDirectory\iTextSharp\itextsharp.dll")

  Try
  {
    [object]$objPhrase = New-PdfPhrase -Text ' '
    [object]$objPdfCell = New-PdfCell -ColSpan $ColSpan -RowSpan $RowSpan
    $objPdfCell.Phrase = $objPhrase
    $objPdfCell.BorderWidthTop = $BorderWidthTop
    $objPdfCell.BorderWidthLeft = $BorderWidthLeft
    $objPdfCell.BorderWidthRight = $BorderWidthRight
    $objPdfCell.BorderWidthBottom = $BorderWidthBottom
    $objPdfCell.BorderColorTop = [iTextSharp.text.BaseColor]::$BorderColorTop
    $objPdfCell.BorderColorLeft = [iTextSharp.text.BaseColor]::$BorderColorLeft
    $objPdfCell.BorderColorRight = [iTextSharp.text.BaseColor]::$BorderColorRight
    $objPdfCell.BorderColorBottom = [iTextSharp.text.BaseColor]::$BorderColorBottom
    $objPdfCell.FixedHeight = $Height
  }
  Catch
  {
    If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "ERROR" -LogMessage $_}
  }

  If ($Log[0]) {Write-Log -LogValue $Log[1] -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t[New-PdfWhiteCell]"}
  
  Return [object]$objPdfCell
}
