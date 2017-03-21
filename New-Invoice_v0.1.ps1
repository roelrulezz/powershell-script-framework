##############################################################
# Get parameters                                             #
##############################################################

#region Get parameters

Param(
  [Parameter(Mandatory=$false, Position=0, HelpMessage="Database ini file")]
  [ValidateNotNullOrEmpty()]
  [string]$DbIniFile = '',
  [Parameter(Mandatory=$false, Position=1, HelpMessage="Exchange ini file")]
  [ValidateNotNullOrEmpty()]
  [string]$ExchangeIniFile = ''
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

If ($DbIniFile -eq '') {[string]$strDbIniFilePath = '\\10.14.1.15\data\HostYourIT\Administratie\Facturen\hostyourit_db.ini'}
Else {[string]$strDbIniFilePath = $DbIniFile}
If ($ExchangeIniFile -eq '') {[string]$strExchangeIniFilePath = '\\10.14.1.15\data\HostYourIT\Administratie\Facturen\hostyourit_exchange.ini'}
Else {[string]$strExchangeIniFilePath = $ExchangeIniFile}

[string]$strYear = '2017'

[string]$strDatabaseTableInvoices = "Facturen_$strYear"
[string]$strDatabaseTableCustomers = 'Klanten'
[string]$strDatabaseTablePrices = "Prijzen_$strYear"

[string]$strPdfFilePath = "\\10.14.1.15\data\HostYourIT\Administratie\$strYear\Factuur\In"
[string]$strPdfAuthor = 'Roeland van den Bosch'

[array]$arrFrameworkColumnWidths = @(10,3,2,16,12,8,15,22) # Percentage

[string]$strImageCompanyLogo = '\\10.14.1.15\data\HostYourIT\Administratie\Facturen\host your it logo.jpg'

[int]$intTextSize = 9

[int]$intHorizontalAlignmentLeft = 0
[int]$intHorizontalAlignmentCenter = 1
[int]$intHorizontalAlignmentRight = 2
[int]$intVerticalAlignmentTop = 3
[int]$intVerticalAlignmentCenter = 4
[int]$intVerticalAlignmentBottom = 5

#endregion Set script variable

##############################################################
# Text                                                       #
##############################################################

#region text

[string]$strLabel_CompanyAddress = "Adres"
[string]$strLabel_ChamberOfCommerceNumber = "KvK nr"
[string]$strLabel_VatNumber = "BTW nr"
[string]$strLabel_BankNumber = "Bank"
[string]$strLabel_PhoneNumber = "Tel"
[string]$strLabel_Subject = "Factuur"
[array]$arrLabel_InvoiceInformation = @("Factuurdatum",
                                        "Uw kenmerk",
                                        "Factuurnummer")
[string]$strLabel_CustomerCompanyName = "Bedrijfsnaam"
[string]$strLabel_HeaderDate = "Datum"
[string]$strLabel_HeaderDescription = "Omschrijving"
[string]$strLabel_HeaderNumber = "Aantal"
[string]$strLabel_HeaderPrice = "Prijs EUR"
[string]$strLabel_HeaderSubTotal = "Bedrag EUR"
[string]$strLabel_Discount = "Korting"
[string]$strLabel_Subtotal = "Subtotaal"
[string]$strLabel_VatPercentage = "BTW 21%"
[string]$strLabel_Total = "Totaal"

[string]$strLabel_Currency = "€"

[array]$arrCompanyAddress = @("Host Your IT",
                              "Jan Tooropstraat 14",
                              "3443 GB Woerden")
[string]$strChamberOfCommerceNumber = "50532820"
[string]$strVatNumber = "NL176065684B01"
[string]$strBankNumber = "NL85 INGB 0004 2134 31"
[string]$strPhoneNumber = "+31 (0)6-16486270"

[string]$strFooter = "Gelieve het factuurbedrag binnen 30 dagen over te maken op rekeningnr. $strBankNumber, t.n.v. Host Your IT, te Woerden, onder vermelding van het factuurnummer."

[string]$strExchangeUsername = "roelandvandenbosch@hostyourit.nl"
[string]$strExchangePassword = "1FXTo72qH1QdG4clFGR8GQ=="

[string]$strMailMessageSignature = "
Met vriendelijke groet,<br>
Roeland van den Bosch<br>
<br>
$($arrCompanyAddress[0])<br>
$strLabel_ChamberOfCommerceNumber`: $strChamberOfCommerceNumber<br>
$strLabel_PhoneNumber`: $strPhoneNumber<br>
E-mail: <a href=""mailto:roelandvandenbosch@hostyourit.nl"">roelandvandenbosch@hostyourit.nl</a><br>
Website: <a href=""http://www.hostyourit.nl/"">http://www.hostyourit.nl/</a><br>
<br>
<img width=""208"" height=""73"" src=""http://www.hostyourit.nl/data/HYI_Signature.png"" border=""0"" type=""image/png""><br>"

#endregion text

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

Function New-InvoiceFile
{
  Param (
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Invoice number")]
    [ValidateNotNullOrEmpty()]
    [string]$InvoiceNumber,
    [Parameter(Mandatory=$true, Position=1, HelpMessage="Invoice line")]
    [ValidateNotNullOrEmpty()]
    [array]$InvoiceLine
  )

  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start function:`t[New-InvoiceFile]"}

  [boolean]$blnInvoiceFile = $false

  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "INFO" -LogMessage "Invoice Number:`t[$InvoiceNumber]"}
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "INFO" -LogMessage "Company Name:`t`t[$($InvoiceLine[0].Bedrijfsnaam)]"}

  #region Get customer information
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Get customer information]"}

  [boolean]$blnCustomerInformation = $false

  [string]$strSqlQuery = "SELECT * FROM $strDatabaseTableCustomers WHERE Bedrijfsnaam = '$($arrInvoiceLine[0].Bedrijfsnaam.Replace("'","''"))'"
  [array]$arrReturnValueDatabaseQuery = New-DatabaseQuery `
                                          -DatabaseConnection $objDatabaseConnection `
                                          -SqlQuery $strSqlQuery
  
  If ($arrReturnValueDatabaseQuery[0] -and $arrReturnValueDatabaseQuery[1].Rows.Count -eq 1)
  {
    [object]$objCustomerInformation = $arrReturnValueDatabaseQuery[1]
    $blnCustomerInformation = $true
  }
  Else
  {
    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "ERROR" -LogMessage "Customer information NOT found [$($InvoiceLine[0].Bedrijfsnaam)]"}
  }

  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Get customer information]"}
  #endregion Get customer information

  If ($blnCustomerInformation)
  {
    #region Create PDF file
    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Create PDF file]"}

    [boolean]$blnPdfFile = $false

    [string]$strPdfFileFullPath = "$strPdfFilePath\$InvoiceNumber`_$($InvoiceLine[0].Bedrijfsnaam.Replace(' ','_').Replace("'",'').Replace('.','')).pdf"
    [array]$arrPdfReport = New-Pdf -File $strPdfFileFullPath -LeftMargin 30 -RightMargin 30 -TopMargin 50 -Author $strPdfAuthor
    If ($arrPdfReport[0])
    {
      [object]$objPdfReport = $arrPdfReport[1]
      $blnPdfFile = $true
    }
    Else
    {
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "ERROR" -LogMessage "Error creating PDF file [$strPdfFileFullPath]"}
    }
    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Create PDF file]"}
    #endregion Create PDF file

    If ($blnPdfFile)
    {
      #region Create framework table
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Create framework table]"}

      [array]$arrColumnWidths = @()
      [int]$intPageWidth = $objPdfReport.PageSize.Width - $objPdfReport.LeftMargin - $objPdfReport.RightMargin
      $arrFrameworkColumnWidths | Foreach-Object {$arrColumnWidths += [int](($intPageWidth / 100) * $_)}
      [object]$objPdfTable_Framework = New-PdfTable -NumberOfColumns $arrFrameworkColumnWidths.Count `
                                                    -TotalWidth $intPageWidth `
                                                    -ColumnWidths $arrColumnWidths `
                                                    -LockedWidth $true

      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Create framework table]"}
      #endregion Create framework table

      #region Add company logo
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Add company logo]"}

      [object]$objPdfCell = New-PdfCell -Colspan 6 `
                                        -RowSpan 11
      $objPdfCell.AddElement($(New-PdfImage -File $strImageCompanyLogo -Scale 30))
      $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Add company logo]"}
      #endregion Add company logo

      #region Add company information
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Add company information]"}

      #region Add company address
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Add company address]"}

      [object]$objPhrase = New-PdfPhrase -Text "$strLabel_CompanyAddress`:" `
                                         -FontDecoration "Bold" `
                                         -FontSize $intTextSize
      [object]$objPdfCell = New-PdfCell -RowSpan 3 `
                                        -HorizontalAlignment $intHorizontalAlignmentRight
      $objPdfCell.Phrase = $objPhrase
      $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

      Foreach ($strCompanyAddress in $arrCompanyAddress)
      {
        [object]$objPhrase = New-PdfPhrase -Text $strCompanyAddress `
                                           -FontSize $intTextSize
        [object]$objPdfCell = New-PdfCell
        $objPdfCell.Phrase = $objPhrase
        $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null
      }

      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Add company address]"}
      #endregion Add company address

      $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 2)) | Out-Null

      #region Add Chamber of Commerce number
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Add Chamber of Commerce number]"}

      [object]$objPhrase = New-PdfPhrase -Text "$strLabel_ChamberOfCommerceNumber`:" `
                                         -FontDecoration "Bold" `
                                         -FontSize $intTextSize
      [object]$objPdfCell = New-PdfCell -HorizontalAlignment $intHorizontalAlignmentRight
      $objPdfCell.Phrase = $objPhrase
      $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

      [object]$objPhrase = New-PdfPhrase -Text $strChamberOfCommerceNumber `
                                         -FontSize $intTextSize
      [object]$objPdfCell = New-PdfCell
      $objPdfCell.Phrase = $objPhrase
      $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Add Chamber of Commerce number]"}
      #endregion Add Chamber of Commerce number

      #region Add VAT number
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Add VAT number]"}

      [object]$objPhrase = New-PdfPhrase -Text "$strLabel_VatNumber`:" `
                                         -FontDecoration "Bold" `
                                         -FontSize $intTextSize
      [object]$objPdfCell = New-PdfCell -HorizontalAlignment $intHorizontalAlignmentRight
      $objPdfCell.Phrase = $objPhrase
      $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

      [object]$objPhrase = New-PdfPhrase -Text $strVatNumber `
                                         -FontSize $intTextSize
      [object]$objPdfCell = New-PdfCell
      $objPdfCell.Phrase = $objPhrase
      $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null
  
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Add VAT number]"}
      #endregion Add VAT number

      $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 2)) | Out-Null

      #region Add bank number
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Add bank number]"}
      
      [object]$objPhrase = New-PdfPhrase -Text "$strLabel_BankNumber`:" `
                                         -FontDecoration "Bold" `
                                         -FontSize $intTextSize
      [object]$objPdfCell = New-PdfCell -HorizontalAlignment $intHorizontalAlignmentRight
      $objPdfCell.Phrase = $objPhrase
      $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

      [object]$objPhrase = New-PdfPhrase -Text $strBankNumber `
                                         -FontSize $intTextSize
      [object]$objPdfCell = New-PdfCell
      $objPdfCell.Phrase = $objPhrase
      $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null
      
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Add bank number]"}
      #endregion Add bank number

      $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 2)) | Out-Null

      #region Add phone number
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Add phone number]"}
  
      [object]$objPhrase = New-PdfPhrase -Text "$strLabel_PhoneNumber`:" `
                                         -FontDecoration "Bold" `
                                         -FontSize $intTextSize
      [object]$objPdfCell = New-PdfCell -HorizontalAlignment $intHorizontalAlignmentRight
      $objPdfCell.Phrase = $objPhrase
      $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

      [object]$objPhrase = New-PdfPhrase -Text $strPhoneNumber `
                                         -FontSize $intTextSize
      [object]$objPdfCell = New-PdfCell
      $objPdfCell.Phrase = $objPhrase
      $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Add phone number]"}
      #endregion Add phone number

      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Add company information]"}
      #endregion Add company information

      $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 2 -BorderWidthBottom 1)) | Out-Null

      #region Add customer information
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Add customer information]"}

      #region Add customer name
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Add customer name]"}
  
      [object]$objPhrase = New-PdfPhrase -Text $objCustomerInformation.Bedrijfsnaam `
                                         -FontSize $intTextSize
      [object]$objPdfCell = New-PdfCell -ColSpan 4
      $objPdfCell.Phrase = $objPhrase
      $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null
      
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Add customer name]"}
      #endregion Add customer name

      $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 4)) | Out-Null

      #region Add customer contact
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Add customer contact]"}
  
      [object]$objPhrase = New-PdfPhrase -Text "T.a.v.: $($objCustomerInformation.ContactPersoon)" `
                                         -FontDecoration "Bold" `
                                         -FontSize $intTextSize
      [object]$objPdfCell = New-PdfCell -ColSpan 4
      $objPdfCell.Phrase = $objPhrase
      $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null
      
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Add customer contact]"}
      #endregion Add customer name

      $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 4)) | Out-Null

      #region Add customer address
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Add customer address]"}
  
      Foreach ($strCustomerAddress in @($objCustomerInformation.Adres, $objCustomerInformation.Postcode, $objCustomerInformation.Plaats))
      {
        [object]$objPhrase = New-PdfPhrase -Text $strCustomerAddress `
                                           -FontSize $intTextSize
        [object]$objPdfCell = New-PdfCell -ColSpan 4
        $objPdfCell.Phrase = $objPhrase
        $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null
        $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 4)) | Out-Null
      }
      
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Add customer address]"}
      #endregion Add customer address

      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Add customer information]"}
      #endregion Add customer information

      $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 8)) | Out-Null

      #region Add subject
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Add subject]"}

      [object]$objPhrase = New-PdfPhrase -Text $strLabel_Subject `
                                         -FontSize 20 `
                                         -FontDecoration "Bold"
      [object]$objPdfCell = New-PdfCell -ColSpan 2 -RowSpan 2
      $objPdfCell.Phrase = $objPhrase
      $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null
      $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 6 -RowSpan 2)) | Out-Null
  
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Add subject]"}
      #endregion Add subject

      $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 8)) | Out-Null

      #region Add invoice information
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Add invoice information]"}

      [array]$arrInvoiceInformation = @($(Get-Date -f "dd-MM-yyyy"), " ", $InvoiceNumber)
      Foreach ($strLabel_InvoiceInformation in $arrLabel_InvoiceInformation)
      {
        [object]$objPhrase = New-PdfPhrase -Text $strLabel_InvoiceInformation `
                                       -FontSize $intTextSize
        [object]$objPdfCell = New-PdfCell -ColSpan 2
        $objPdfCell.Phrase = $objPhrase
        $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

        [object]$objPhrase = New-PdfPhrase -Text ":" `
                                           -FontSize $intTextSize
        [object]$objPdfCell = New-PdfCell -HorizontalAlignment $intHorizontalAlignmentRight
        $objPdfCell.Phrase = $objPhrase
        $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

        [object]$objPhrase = New-PdfPhrase -Text $arrInvoiceInformation[$([array]::indexof($arrLabel_InvoiceInformation,$strLabel_InvoiceInformation))] `
                                           -FontSize $intTextSize
        [object]$objPdfCell = New-PdfCell -ColSpan 2
        $objPdfCell.Phrase = $objPhrase
        $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null
  
        $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 3)) | Out-Null
      }

      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Add invoice information]"}
      #endregion Add invoice information

      $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 8 -BorderWidthBottom 1)) | Out-Null
      $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 8)) | Out-Null

      #region Add header
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Add header]"}
  
      [object]$objPhrase = New-PdfPhrase -Text $strLabel_HeaderDate `
                                         -FontDecoration "Bold" `
                                         -FontSize $intTextSize
      [object]$objPdfCell = New-PdfCell
      $objPdfCell.Phrase = $objPhrase
      $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

      [object]$objPhrase = New-PdfPhrase -Text $strLabel_HeaderDescription `
                                         -FontDecoration "Bold" `
                                         -FontSize $intTextSize
      If ($InvoiceLine.VertoonAantal) {[object]$objPdfCell = New-PdfCell -ColSpan 4}
      Else {[object]$objPdfCell = New-PdfCell -ColSpan 6}
      $objPdfCell.Phrase = $objPhrase
      $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

      If ($InvoiceLine.VertoonAantal)
      {
        [object]$objPhrase = New-PdfPhrase -Text $strLabel_HeaderNumber `
                                           -FontDecoration "Bold" `
                                           -FontSize $intTextSize
        [object]$objPdfCell = New-PdfCell
        $objPdfCell.Phrase = $objPhrase
        $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

        [object]$objPhrase = New-PdfPhrase -Text $strLabel_HeaderPrice `
                                           -FontDecoration "Bold" `
                                           -FontSize $intTextSize
        [object]$objPdfCell = New-PdfCell -HorizontalAlignment $intHorizontalAlignmentRight
        $objPdfCell.Phrase = $objPhrase
        $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null
      }

      [object]$objPhrase = New-PdfPhrase -Text $strLabel_HeaderSubTotal `
                                         -FontDecoration "Bold" `
                                         -FontSize $intTextSize
      [object]$objPdfCell = New-PdfCell -HorizontalAlignment $intHorizontalAlignmentRight
      $objPdfCell.Phrase = $objPhrase
      $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null
      
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Add header]"}
      #endregion Add header

      $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 8 -BorderWidthBottom 1)) | Out-Null
      $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 8)) | Out-Null

      #region Add products
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Add products]"}
      [decimal]$decTotal = 0
      [decimal]$decDiscount = 0
      [boolean]$blnPriceInformation = $true
      For ($i = 0;$i -le 12;$i++)
      {
        If ($i -lt $InvoiceLine.Count)
        {
          [object]$objPhrase = New-PdfPhrase -Text $("{0:dd-MM-yyyy}" -f $InvoiceLine[$i].Datum) `
                                             -FontSize $intTextSize
          [object]$objPdfCell = New-PdfCell
          $objPdfCell.Phrase = $objPhrase
          $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

          [object]$objPhrase = New-PdfPhrase -Text $InvoiceLine[$i].Omschrijving `
                                             -FontSize $intTextSize
          If ($InvoiceLine.VertoonAantal) {[object]$objPdfCell = New-PdfCell -Colspan 4}
          Else {[object]$objPdfCell = New-PdfCell -ColSpan 6}
          $objPdfCell.Phrase = $objPhrase
          $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

          #region Get price information
          If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Get price information]"}

          [string]$strSqlQuery = "SELECT Prijs FROM $strDatabaseTablePrices WHERE HoofdCategorie = '$($InvoiceLine[$i].HoofdCategorie)' AND SubCategorie = '$($InvoiceLine[$i].SubCategorie)'"
          [array]$arrReturnValueDatabaseQuery = New-DatabaseQuery `
                                                  -DatabaseConnection $objDatabaseConnection `
                                                  -SqlQuery $strSqlQuery
  
          If ($arrReturnValueDatabaseQuery[0] -and $arrReturnValueDatabaseQuery[1].Rows.Count -eq 1)
          {
            [decimal]$decPrice = ($arrReturnValueDatabaseQuery[1].Prijs)
            $blnPriceInformation = $blnPriceInformation -and $true
          }
          Else
          {
            $blnPriceInformation = $blnPriceInformation -and $false
            If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "ERROR" -LogMessage "Price information NOT found [$($InvoiceLine[$i].HoofdCategorie) \ $($InvoiceLine[$i].SubCategorie)]"}
          }
          If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Get price information]"}
          #endregion Get price information

          If ($blnPriceInformation)
          {
            If ($InvoiceLine.VertoonAantal)
            {
              [object]$objPhrase = New-PdfPhrase -Text $InvoiceLine[$i].Aantal `
                                                 -FontSize $intTextSize
              [object]$objPdfCell = New-PdfCell -HorizontalAlignment $intHorizontalAlignmentCenter
              $objPdfCell.Phrase = $objPhrase
              $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

              [object]$objPhrase = New-PdfPhrase -Text "$strLabel_Currency $($decPrice.ToString().Replace('.',','))" `
                                                 -FontSize $intTextSize
              [object]$objPdfCell = New-PdfCell -HorizontalAlignment $intHorizontalAlignmentRight
              $objPdfCell.Phrase = $objPhrase
              $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null
            }

            [decimal]$decSubTotal = $decPrice * $InvoiceLine[$i].Aantal
            $decDiscount += $InvoiceLine[$i].Korting * $InvoiceLine[$i].Aantal
            $decTotal += $decSubTotal
            [object]$objPhrase = New-PdfPhrase -Text "$strLabel_Currency $($decSubTotal.ToString().Replace('.',','))" `
                                             -FontSize $intTextSize
            [object]$objPdfCell = New-PdfCell -HorizontalAlignment $intHorizontalAlignmentRight
            $objPdfCell.Phrase = $objPhrase
            $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

            #region Update invoice line in database
            If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Update invoice line in database]"}
            [string]$strSqlQuery = "UPDATE $strDatabaseTableInvoices SET FactuurNummer = '$InvoiceNumber' WHERE Id = '$($InvoiceLine[$i].Id)'"
            [array]$arrReturnValueDatabaseQuery = New-DatabaseQuery `
                                                    -DatabaseConnection $objDatabaseConnection `
                                                    -SqlQuery $strSqlQuery 

            If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Update invoice line in database]"}
            #endregion Update invoice line in database
          }
        }
        Else
        {
          $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 8 -Height 15)) | Out-Null
        }
      }
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Add products]"}
      #endregion Add products

      If ($blnPriceInformation)
      {
        #region Add discount
        If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Add discount]"}
        
        If ($decDiscount)
        {
          $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -Height 15)) | Out-Null

          [object]$objPhrase = New-PdfPhrase -Text $strLabel_Discount `
                                             -FontSize $intTextSize
          [object]$objPdfCell = New-PdfCell -ColSpan 4 `
                                            -HorizontalAlignment $intHorizontalAlignmentLeft
          $objPdfCell.Phrase = $objPhrase
          $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

          $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 2 -Height 15)) | Out-Null

          [object]$objPhrase = New-PdfPhrase -Text "- $strLabel_Currency $($decDiscount.ToString().Replace('.',','))" `
                                             -FontSize $intTextSize
          [object]$objPdfCell = New-PdfCell -HorizontalAlignment $intHorizontalAlignmentRight
          $objPdfCell.Phrase = $objPhrase
          $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null
        }
        Else
        {
          $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 8 -Height 15)) | Out-Null
        }
        
        If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Add discount]"}
        #endregion Add discount 

        $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 8 -BorderWidthBottom 1)) | Out-Null
        $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 8)) | Out-Null

        #region Add total price
        If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Add total price]"}

        $decTotal -= $decDiscount

        $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 6)) | Out-Null
        [object]$objPhrase = New-PdfPhrase -Text "$strLabel_Subtotal`:" `
                                           -FontSize $intTextSize
        [object]$objPdfCell = New-PdfCell -HorizontalAlignment $intHorizontalAlignmentRight
        $objPdfCell.Phrase = $objPhrase
        $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

        [object]$objPhrase = New-PdfPhrase -Text "$strLabel_Currency $($decTotal.ToString().Replace('.',','))" `
                                           -FontSize $intTextSize
        [object]$objPdfCell = New-PdfCell -HorizontalAlignment $intHorizontalAlignmentRight
        $objPdfCell.Phrase = $objPhrase
        $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

        $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 6)) | Out-Null
        [object]$objPhrase = New-PdfPhrase -Text "$strLabel_VatPercentage`:" `
                                           -FontSize $intTextSize
        [object]$objPdfCell = New-PdfCell -HorizontalAlignment $intHorizontalAlignmentRight
        $objPdfCell.Phrase = $objPhrase
        $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

        [object]$objPhrase = New-PdfPhrase -Text "$strLabel_Currency $(([math]::Round($decTotal * 0.21, 2)).ToString().Replace('.',','))" `
                                           -FontSize $intTextSize
        [object]$objPdfCell = New-PdfCell -HorizontalAlignment $intHorizontalAlignmentRight
        $objPdfCell.Phrase = $objPhrase
        $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

        $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 6)) | Out-Null
        [object]$objPhrase = New-PdfPhrase -Text "$strLabel_Total`:" `
                                           -FontDecoration "Bold" `
                                           -FontSize $intTextSize
        [object]$objPdfCell = New-PdfCell -HorizontalAlignment $intHorizontalAlignmentRight
        $objPdfCell.Phrase = $objPhrase
        $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

        [object]$objPhrase = New-PdfPhrase -Text "$strLabel_Currency $(([math]::Round($decTotal * 1.21, 2)).ToString().Replace('.',','))" `
                                           -FontDecoration "Bold" `
                                           -FontSize $intTextSize
        [object]$objPdfCell = New-PdfCell -HorizontalAlignment $intHorizontalAlignmentRight
        $objPdfCell.Phrase = $objPhrase
        $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null

        If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Add total price]"}
        #endregion Add total price

        $objPdfTable_Framework.AddCell($(New-PdfWhiteCell -ColSpan 8 -Height 60)) | Out-Null

        #region Add footer
        If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Add footer]"}
        
        [object]$objPhrase = New-PdfPhrase -Text $strFooter `
                                           -FontSize $intTextSize
        [object]$objPdfCell = New-PdfCell -ColSpan 8 `
                                          -HorizontalAlignment $intHorizontalAlignmentCenter
        $objPdfCell.Phrase = $objPhrase
        $objPdfCell.PaddingLeft = 30
        $objPdfCell.PaddingRight = 30
        $objPdfTable_Framework.AddCell($objPdfCell) | Out-Null
        
        If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Add footer]"}
        #endregion Add footer

        #region Open, save and close PDF file
        If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Open, save and close PDF file]"}
        
        $objPdfReport.Open()
        $objPdfReport.Add($objPdfTable_Framework) | Out-Null
        $objPdfReport.Close()

        If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Open, save and close PDF file]"}
        #endregion Open, save and close PDF file

        $blnInvoiceFile = $true
      }
    }
  }

  If ($blnConsoleLog -and $blnLogLevel4) {Write-ConsoleLog -LogLevel "DEBUG" -LogMessage "End function:`t`t[New-InvoiceFile]"}
  If ($blnFileLog -and $blnLogLevel4) {Write-FileLog -LogFile $strLogFile -LogLevel "DEBUG" -LogMessage "End function:`t`t[New-InvoiceFile]"}

  Return @($blnInvoiceFile,$strPdfFileFullPath)
}

Function Send-InvoiceMail
{
  Param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Exchange Service Object")]
    [ValidateNotNullOrEmpty()]
    [object]$ExchangeServiceObject,
    [Parameter(Mandatory=$true, Position=1, HelpMessage="Invoice Number")]
    [ValidateNotNullOrEmpty()]
    [string]$InvoiceNumber,
    [Parameter(Mandatory=$true, Position=2, HelpMessage="Invoice Information")]
    [ValidateNotNullOrEmpty()]
    [array]$InvoiceInformation,
    [Parameter(Mandatory=$true, Position=3, HelpMessage="Invoice Full File Path")]
    [ValidateNotNullOrEmpty()]
    [string]$InvoiceFullFilePath
  )

  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start function:`t[Send-InvoiceMail]"}

  #region Get customer information
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Get customer information]"}

  [string]$strSqlQuery = "SELECT * FROM $strDatabaseTableCustomers WHERE Bedrijfsnaam = '$($InvoiceInformation[0].Bedrijfsnaam.Replace("'","''"))'"
  [array]$arrReturnValueDatabaseQuery = New-DatabaseQuery `
                                          -DatabaseConnection $objDatabaseConnection `
                                          -SqlQuery $strSqlQuery
  
  If ($arrReturnValueDatabaseQuery[0])
  {
    [boolean]$blnCustomerInformation = $true
    [object]$objCustomerInformation = $arrReturnValueDatabaseQuery[1]
  }
  Else
  {
    [boolean]$blnCustomerInformation = $false
  }

  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Get customer information]"}
  #endregion Get customer information

  If ($blnCustomerInformation)
  {
    #region Set address information
    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Set address information]"}

    [string]$strFrom = "roelandvandenbosch@hostyourit.nl"
    [array]$arrToRecipients = @($objCustomerInformation.EmailFactuur)
    [array]$arrBccRecipients = @("factuur@hostyourit.nl")

    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "INFO" -LogMessage "Send mail to:`t`t[$arrToRecipients]"}

    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Set address information]"}
    #endregion Set address information

    #region Set subject
    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Set subject]"}

    [array]$arrMainSubjects = $InvoiceInformation.HoofdCategorie | Sort-Object -Unique

    If ($arrMainSubjects.Count -eq 1)
    {
      Switch ($arrMainSubjects[0])
      {
        "Hosting" {[string]$strSubject = "Factuur hosting $('{0:MMMM}' -f $($InvoiceInformation[0].Datum)) [$InvoiceNumber]"}
        "Domein" {[string]$strSubject = "Factuur domein registratie [$InvoiceNumber]"}
      }
    }

    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "INFO" -LogMessage "Mail subject:`t`t[$strSubject]"}

    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Set subject]"}
    #endregion Set subject

    #region Create body
    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Create body]"}

    [string]$strMailMessageBody = "Beste $(($objCustomerInformation.ContactPersoon).Split(" ")[0]),</br>`r`n</br>`r`n"
    If ($arrMainSubjects.Count -eq 1)
    {
      Switch ($arrMainSubjects[0])
      {
        "Hosting" {$strMailMessageBody += "Hierbij de factuur van $('{0:MMMM}' -f $($InvoiceInformation[0].Datum)).<br>`r`n"}
        "Domein" {$strMailMessageBody += "Hierbij de factuur voor de domein registratie(s) van:<br>`r`n";
                  $InvoiceInformation | Foreach-Object {$strMailMessageBody += " - $($_.Omschrijving)<br>`r`n"}}
      }
    }

    $strMailMessageBody += "<br>`r`n"

    $strMailMessageBody = ConvertTo-Html -Head "<style>BODY{font-family:Tahoma; font-size:10pt;}</style>" `
                                         -Body $strMailMessageBody `
                                         -PostContent $strMailMessageSignature 
   
    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Create body]"}
    #endregion Create body

    #region Create email object
    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Create email object]"}

    [array]$arrNewEmailMessageReturnValue = New-EmailMessageObject -ExchangeServiceObject $ExchangeServiceObject `
                                                                   -From $strFrom `
                                                                   -ToRecipients $arrToRecipients `
                                                                   -BccRecipients $arrBccRecipients `
                                                                   -Subject $strSubject `                                                                   -HtmlBody $true `
                                                                   -Body $strMailMessageBody `
                                                                   -Attachment $InvoiceFullFilePath

    If ($arrNewEmailMessageReturnValue[0])
    {
      [boolean]$blnNewEmailMessage = $true
      [object]$objEmailMessage = $arrNewEmailMessageReturnValue[1]
    }
    Else
    {
      [boolean]$blnNewEmailMessage = $false
    }

    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Create email object]"}
    #endregion Create email object

    If ($blnNewEmailMessage)
    {
      #region Send email message
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Send email message]"}

      $objEmailMessage.SendAndSaveCopy()
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "INFO" -LogMessage "Mail message send"}

      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Send email message]"}
      #endregion Send email message
    }
  }
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End function:`t`t[Send-InvoiceMail]"}
}

#endregion Function

##############################################################
# Main                                                       #
##############################################################

If ($blnCreditShow) {Show-Credit}

If ($blnLog) {Start-Log -LogValue $objLogValue}

#region Read ini file
If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Read ini file]"}

[object]$objDbIniFile = Read-IniFile -FilePath $strDbIniFilePath
[string]$strDatabaseServer = $($objDbIniFile.default.server)
[string]$strDatabaseName = $($objDbIniFile.Database.name)
[string]$strDatabaseUsername = $($objDbIniFile.User.user)
[string]$strDatabaseUserEncryptedPassword = $($objDbIniFile.User.password)

[object]$objExchangeIniFile = Read-IniFile -FilePath $strExchangeIniFilePath
[string]$strExchangeUsername = $($objExchangeIniFile.default.username)
[string]$strExchangeUserEncryptedPassword = $($objExchangeIniFile.default.password)

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
  #region Get last invoice number
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Get last invoice number]"}

  [boolean]$blnInvoiceNumber = $false

  [string]$strSqlQuery = "SELECT MAX(FactuurNummer) AS FactuurNummer FROM $strDatabaseTableInvoices WHERE FactuurNummer IS NOT NULL"
  [array]$arrReturnValueDatabaseQuery = New-DatabaseQuery `
                                          -DatabaseConnection $objDatabaseConnection `
                                          -SqlQuery $strSqlQuery
  
  If ($arrReturnValueDatabaseQuery[0])
  {
    If ($arrReturnValueDatabaseQuery[1].FactuurNummer -eq [System.DBNull]::Value) {[string]$strInvoiceNumber = "$strYear-001"}
    Else {[string]$strInvoiceNumber = "$($arrReturnValueDatabaseQuery[1].FactuurNummer.Split('-')[0])-$("{0:D3}" -f ([int]($arrReturnValueDatabaseQuery[1].FactuurNummer.Split('-')[1]) + 1))"}

    $blnInvoiceNumber = $true

    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "INFO" -LogMessage "Invoice number [$strInvoiceNumber]"}
  }
  Else
  {
    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "ERROR" -LogMessage "NO invoice number found"}
  }

  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Get last invoice number]"}
  #endregion Get last invoice number

  If ($blnInvoiceNumber)
  {
    #region Get open invoice lines
    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Get open invoice lines]"}

    [boolean]$blnOpenInvoiceLines = $false

    [string]$strSqlQuery = "SELECT * FROM $strDatabaseTableInvoices WHERE FactuurNummer IS NULL ORDER BY Id"
    [array]$arrReturnValueDatabaseQuery = New-DatabaseQuery `
                                            -DatabaseConnection $objDatabaseConnection `
                                            -SqlQuery $strSqlQuery

    If ($arrReturnValueDatabaseQuery[0])
    {
      If ($arrReturnValueDatabaseQuery[1].Rows.Count -gt 0) {$blnOpenInvoiceLines = $true}
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "INFO" -LogMessage "Number of open invoice lines [$($arrReturnValueDatabaseQuery[1].Rows.Count)]"}
    }
    Else
    {
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "ERROR" -LogMessage "NO invoice lines found"}
    }

    If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Get open invoice lines]"}
    #endregion Get open invoice lines

    If ($blnOpenInvoiceLines)
    {
      #region Generate invoice file
      If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Generate invoice file]"}

      [array]$arrInvoiceLine = @()
      [boolean]$blnMailInvoice = $true

      If ($arrReturnValueDatabaseQuery[1].Rows.MailFactuur -contains 1)
      {
        #region Create MS Exchange service object
        If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Create MS Exchange service object]"}
        [boolean]$blnContinue = $false
        If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Create MS Exchange service object"}
        [array]$arrNewExchangeServiceObjectReturnValue = New-ExchangeServiceObject -ExchangeUsername $strExchangeUsername `
                                                                                   -ExchangePassword $(New-DecryptedString -InputString $strExchangeUserEncryptedPassword -Passphrase $strExchangeUsername) `
                                                                                   -Log @($true, $objLogValue)
        If ($arrNewExchangeServiceObjectReturnValue[0])
        {
          [object]$objExchangeService = $arrNewExchangeServiceObjectReturnValue[1]
          $blnContinue = $true
        }
        If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Create MS Exchange service object]"}
        #endregion Create MS Exchange service object
      }
      Else
      {
        [boolean]$blnContinue = $true
      }

      If ($blnContinue)
      {
        For ($intIndexCounter = 0; $intIndexCounter -lt $arrReturnValueDatabaseQuery[1].Rows.Count; $intIndexCounter++)
        {
          $arrInvoiceLine += $arrReturnValueDatabaseQuery[1].Rows[$intIndexCounter]
          $blnMailInvoice = $blnMailInvoice -and $arrReturnValueDatabaseQuery[1].Rows[$intIndexCounter].MailFactuur
          If ($($arrReturnValueDatabaseQuery[1].Rows[$intIndexCounter].Bedrijfsnaam) -ne $($arrReturnValueDatabaseQuery[1].Rows[$intIndexCounter + 1].Bedrijfsnaam))
          {
            [array]$arrReturnValueInvoiceFile = New-InvoiceFile -InvoiceNumber $strInvoiceNumber -InvoiceLine $arrInvoiceLine

            If ($arrReturnValueInvoiceFile[0] -and $blnMailInvoice)
            {
              Send-InvoiceMail -ExchangeServiceObject $objExchangeService `
                               -InvoiceNumber $strInvoiceNumber `
                               -InvoiceInformation $arrInvoiceLine `
                               -InvoiceFullFilePath $arrReturnValueInvoiceFile[1]
            }
            $strInvoiceNumber = "$($strInvoiceNumber.Split('-')[0])-$("{0:D3}" -f ([int]($strInvoiceNumber.Split('-')[1]) + 1))"
            $arrInvoiceLine = @()
            $blnMailInvoice = $true
          }
        }

        If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Generate invoice file]"}
        #endregion Generate invoice file
      }
    }
  }

  #region Disconnect from database server
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "Start region:`t`t[Disconnect from database server]"}
  $objDatabaseConnection.Close()
  $blnDatabaseConnection = $false
  If ($blnLog) {Write-Log -LogValue $objLogValue -LogMessageLevel "DEBUG" -LogMessage "End region:`t`t[Disconnect from database server]"}
  #endregion Disconnect from database server
}

If ($blnLog) {Stop-Log -LogValue $objLogValue}
