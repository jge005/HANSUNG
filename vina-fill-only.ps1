param(
  [string]$PayloadPath = (Join-Path $PSScriptRoot 'vina-payload.sample.json'),
  [string]$OutputPath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Copy-CellText($targetWs, [string]$targetAddress, $sourceWs, [string]$sourceAddress) {
  $targetWs.Range($targetAddress).Value2 = [string]$sourceWs.Range($sourceAddress).Text
}

function Get-CellText($ws, [string]$address) {
  return ([string]$ws.Range($address).Text).Trim()
}

function Set-CellText($ws, [string]$address, [string]$value) {
  try {
    $ws.Range($address).Value2 = $value
  } catch {
    throw ("Set-CellText failed at {0}: {1}" -f $address, $_.Exception.Message)
  }
}

function Set-CellNumber($ws, [string]$address, $value) {
  if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) {
    $ws.Range($address).ClearContents() | Out-Null
    return
  }
  try {
    $ws.Range($address).Value2 = [double]$value
  } catch {
    throw ("Set-CellNumber failed at {0}: {1}" -f $address, $_.Exception.Message)
  }
}

function Set-MergedCellTextByPosition($ws, [int]$row, [int]$col, [string]$value) {
  try {
    $cell = $ws.Cells.Item($row, $col)
    $mergeArea = $cell.MergeArea
    if ($mergeArea -and $mergeArea.Cells.Count -ge 1) {
      $mergeArea.Cells.Item(1, 1).Value2 = $value
    } else {
      $cell.Value2 = $value
    }
  } catch {
    throw ("Set-MergedCellText failed at r{0} c{1}: {2}" -f $row, $col, $_.Exception.Message)
  }
}

function Parse-WeightNumber($text) {
  $raw = [string]$text
  if ([string]::IsNullOrWhiteSpace($raw)) { return 0.0 }
  $clean = $raw.ToUpper().Replace('KG','').Replace(',','').Trim()
  $value = 0.0
  if ([double]::TryParse($clean, [ref]$value)) { return $value }
  return 0.0
}

function Format-WeightText($value) {
  if ($value -le 0) { return '' }
  return ('{0:0.##}kg' -f [double]$value)
}

function Normalize-Boxes($payloadBoxes, $payloadItems) {
  $items = @($payloadItems)
  $boxes = @($payloadBoxes)
  if ($boxes.Count -gt 0) { return $boxes }
  return @(
    [pscustomobject]@{
      boxNo = 1
      items = @(
        $items | ForEach-Object {
          [pscustomobject]@{
            code = [string]$_.code
            qty = $_.qty
          }
        }
      )
    }
  )
}

function Get-KoreanString([int[]]$codes) {
  return -join ($codes | ForEach-Object { [char]$_ })
}

function Release-ComObjectSafe($obj) {
  if ($null -ne $obj) {
    try { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj) } catch {}
  }
}

$payload = Get-Content -LiteralPath $PayloadPath -Raw | ConvertFrom-Json
$shipDate = [datetime]$payload.shipDate
$transport = if ([string]::IsNullOrWhiteSpace([string]$payload.transport)) { 'AIR' } elseif ([string]$payload.transport -match 'VSL') { 'VSL' } else { 'AIR' }
$boxes = Normalize-Boxes $payload.boxes $payload.items

$desktop = [Environment]::GetFolderPath('Desktop')
$sourcePath = Join-Path $desktop ((Get-KoreanString @(0xCD9C,0xD558,0xC815,0xBCF4)) + '.xlsx')
$templatePath = Get-ChildItem -LiteralPath $desktop -Filter '*.xlsx' | Where-Object {
  if ($transport -eq 'VSL') { $_.Name -like '*Invoice+Packing*VSL*VM(DEV).xlsx' }
  else { $_.Name -like '*Invoice+Packing*AIR*VM(DEV).xlsx' }
} | Select-Object -First 1 -ExpandProperty FullName

if (-not $templatePath) { throw 'Template workbook not found.' }

$token = $shipDate.ToString('yyMMdd')
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
  $OutputPath = Join-Path $PSScriptRoot ("[HANSUNG] {0} VINA ({1}) filled.xlsx" -f $token, $transport)
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$sourceWb = $null
$targetWb = $null

try {
  $sourceWb = $excel.Workbooks.Open($sourcePath)
  $shipWs = $sourceWb.Worksheets.Item(4)

  $shipWs.Range('B3:E610').ClearContents() | Out-Null

  $writeRow = 3
  foreach ($item in $payload.items) {
    Set-CellText $shipWs ("B{0}" -f $writeRow) ([string]$item.code)
    Set-CellText $shipWs ("C{0}" -f $writeRow) ([string]$item.name)
    Set-CellNumber $shipWs ("D{0}" -f $writeRow) $item.qty
    if ($item.PSObject.Properties.Name -contains 'price') {
      Set-CellNumber $shipWs ("E{0}" -f $writeRow) $item.price
    }
    $writeRow++
  }

  if ($transport -eq 'AIR') {
    $sourceWb.Worksheets.Item(6).Range('C4').Value2 = $shipDate
  } else {
    $sourceWb.Worksheets.Item(8).Range('C4').Value2 = $shipDate
    if ($payload.PSObject.Properties.Name -contains 'remarkCode' -and -not [string]::IsNullOrWhiteSpace([string]$payload.remarkCode)) {
      Set-CellText $sourceWb.Worksheets.Item(8) 'C14' ([string]$payload.remarkCode)
    }
  }

  $sourceWb.RefreshAll()
  $excel.CalculateFullRebuild()

  $sourceInvoiceWs = if ($transport -eq 'AIR') { $sourceWb.Worksheets.Item(6) } else { $sourceWb.Worksheets.Item(8) }
  $sourcePackingWs = if ($transport -eq 'AIR') { $sourceWb.Worksheets.Item(7) } else { $sourceWb.Worksheets.Item(9) }

  $targetWb = $excel.Workbooks.Open($templatePath)

  $targetInvoiceWs = $targetWb.Worksheets.Item(1)
  $targetPackingWs = $targetWb.Worksheets.Item(2)
  $targetDetailWs = $targetWb.Worksheets.Item(3)
  $targetDetailPackingWs = $targetWb.Worksheets.Item(4)
  $targetCntrWs = $targetWb.Worksheets.Item(5)
  $targetShipmarkWs = $targetWb.Worksheets.Item(6)

  $targetInvoiceWs.Range('A25:G120').ClearContents() | Out-Null
  Copy-CellText $targetInvoiceWs 'A4' $sourceInvoiceWs 'A4'
  Copy-CellText $targetInvoiceWs 'C4' $sourceInvoiceWs 'C4'
  Copy-CellText $targetInvoiceWs 'A10' $sourceInvoiceWs 'A8'
  Copy-CellText $targetInvoiceWs 'C10' $sourceInvoiceWs 'C8'
  Copy-CellText $targetInvoiceWs 'A16' $sourceInvoiceWs 'A12'
  Copy-CellText $targetInvoiceWs 'C13' $sourceInvoiceWs 'C10'
  Copy-CellText $targetInvoiceWs 'B22' $sourceInvoiceWs 'B16'
  Copy-CellText $targetInvoiceWs 'C18' $sourceInvoiceWs 'C14'

  $row = 18
  $targetRow = 25
  while ($targetRow -le 120) {
    $code = Get-CellText $sourceInvoiceWs ("B{0}" -f $row)
    if ([string]::IsNullOrWhiteSpace($code)) { break }
    Set-CellText $targetInvoiceWs ("A{0}" -f $targetRow) (Get-CellText $sourceInvoiceWs ("A{0}" -f $row))
    Set-CellText $targetInvoiceWs ("B{0}" -f $targetRow) $code
    Set-CellText $targetInvoiceWs ("C{0}" -f $targetRow) (Get-CellText $sourceInvoiceWs ("C{0}" -f $row))
    Set-CellText $targetInvoiceWs ("D{0}" -f $targetRow) 'ea'
    Set-CellText $targetInvoiceWs ("E{0}" -f $targetRow) (Get-CellText $sourceInvoiceWs ("D{0}" -f $row))
    Set-CellText $targetInvoiceWs ("F{0}" -f $targetRow) (Get-CellText $sourceInvoiceWs ("E{0}" -f $row))
    Set-CellText $targetInvoiceWs ("G{0}" -f $targetRow) (Get-CellText $sourceInvoiceWs ("F{0}" -f $row))
    $row++
    $targetRow++
  }

  $itemMeta = @{}
  foreach ($item in @($payload.items)) {
    $code = [string]$item.code
    if ([string]::IsNullOrWhiteSpace($code)) { continue }
    $qtyValue = if ($null -ne $item.qty -and [string]$item.qty -ne '') { [double]$item.qty } else { 0.0 }
    $priceValue = if ($null -ne $item.price -and [string]$item.price -ne '') { [double]$item.price } else { 0.0 }
    $itemMeta[$code] = [pscustomobject]@{
      code = $code
      specification = [string]$item.name
      qty = $qtyValue
      price = $priceValue
      amount = ($qtyValue * $priceValue)
      net = 0.0
      gross = 0.0
    }
  }

  $row = if ($transport -eq 'AIR') { 22 } else { 20 }
  while ($true) {
    $code = Get-CellText $sourcePackingWs ("B{0}" -f $row)
    if ([string]::IsNullOrWhiteSpace($code)) { break }
    $existing = if ($itemMeta.ContainsKey($code)) { $itemMeta[$code] } else { [pscustomobject]@{ price=0.0; amount=0.0; specification=$code } }
    $itemMeta[$code] = [pscustomobject]@{
      code = $code
      specification = [string]$existing.specification
      qty = [double](([string](Get-CellText $sourcePackingWs ("C{0}" -f $row))).Replace(',',''))
      price = [double]$existing.price
      amount = [double]$existing.amount
      net = Parse-WeightNumber (Get-CellText $sourcePackingWs ("D{0}" -f $row))
      gross = Parse-WeightNumber (Get-CellText $sourcePackingWs ("E{0}" -f $row))
    }
    $row++
  }

  $row = 18
  while ($true) {
    $code = Get-CellText $sourceInvoiceWs ("B{0}" -f $row)
    if ([string]::IsNullOrWhiteSpace($code)) { break }
    $existing = if ($itemMeta.ContainsKey($code)) { $itemMeta[$code] } else { [pscustomobject]@{ code=$code; qty=0.0; net=0.0; gross=0.0 } }
    $amountText = (Get-CellText $sourceInvoiceWs ("E{0}" -f $row)).Replace(',','')
    $priceText = (Get-CellText $sourceInvoiceWs ("D{0}" -f $row)).Replace(',','')
    $qtyText = (Get-CellText $sourceInvoiceWs ("C{0}" -f $row)).Replace(',','')
    $itemMeta[$code] = [pscustomobject]@{
      code = $code
      specification = if ($payload.items | Where-Object { [string]$_.code -eq $code } | Select-Object -First 1) {
        [string](($payload.items | Where-Object { [string]$_.code -eq $code } | Select-Object -First 1).name)
      } else { $existing.specification }
      qty = if ($qtyText) { [double]$qtyText } else { [double]$existing.qty }
      price = if ($priceText) { [double]$priceText } else { 0.0 }
      amount = if ($amountText) { [double]$amountText } else { 0.0 }
      net = [double]$existing.net
      gross = [double]$existing.gross
    }
    $row++
  }

  $targetPackingWs.Range('A23:F120').ClearContents() | Out-Null
  $consigneeAddress = Get-CellText $sourcePackingWs 'A10'
  if ([string]::IsNullOrWhiteSpace($consigneeAddress)) { $consigneeAddress = Get-CellText $sourcePackingWs 'A9' }
  Copy-CellText $targetPackingWs 'A4' $sourcePackingWs 'A4'
  Copy-CellText $targetPackingWs 'D4' $sourcePackingWs 'D4'
  Set-CellText $targetPackingWs 'A10' $consigneeAddress
  Copy-CellText $targetPackingWs 'D10' $sourcePackingWs 'D10'
  Copy-CellText $targetPackingWs 'A17' $sourcePackingWs 'A15'
  Copy-CellText $targetPackingWs 'B17' $sourcePackingWs 'B16'
  Copy-CellText $targetPackingWs 'B20' $sourcePackingWs 'B19'
  if ([string]::IsNullOrWhiteSpace((Get-CellText $targetPackingWs 'B20'))) { Copy-CellText $targetPackingWs 'B20' $sourcePackingWs 'B18' }
  Copy-CellText $targetPackingWs 'D20' $sourcePackingWs 'D19'
  if ([string]::IsNullOrWhiteSpace((Get-CellText $targetPackingWs 'D20'))) { Copy-CellText $targetPackingWs 'D20' $sourcePackingWs 'D18' }

  $row = if ($transport -eq 'VSL') { 20 } else { 22 }
  $targetRow = 23
  while ($targetRow -le 120) {
    $code = Get-CellText $sourcePackingWs ("B{0}" -f $row)
    if ([string]::IsNullOrWhiteSpace($code)) { break }
    Set-CellText $targetPackingWs ("A{0}" -f $targetRow) (Get-CellText $sourcePackingWs ("A{0}" -f $row))
    Set-CellText $targetPackingWs ("B{0}" -f $targetRow) $code
    Set-CellText $targetPackingWs ("C{0}" -f $targetRow) (Get-CellText $sourcePackingWs ("C{0}" -f $row))
    Set-CellText $targetPackingWs ("D{0}" -f $targetRow) (Get-CellText $sourcePackingWs ("D{0}" -f $row))
    Set-CellText $targetPackingWs ("E{0}" -f $targetRow) (Get-CellText $sourcePackingWs ("E{0}" -f $row))
    Set-CellText $targetPackingWs ("F{0}" -f $targetRow) (Get-CellText $sourcePackingWs ("F{0}" -f $row))
    $row++
    $targetRow++
  }

  if ($transport -eq 'AIR') {
    $targetDetailWs.Range('B3:M36').ClearContents() | Out-Null
    $detailStartRow = 3
    $detailQtyCol = 'F'; $detailPriceCol = 'H'; $detailAmountCol = 'I'; $detailNetCol = 'J'; $detailGrossCol = 'K'; $detailOriginCol1 = 'L'; $detailOriginCol2 = 'M'; $detailTotalRow = 36
  } else {
    $targetDetailWs.Range('B3:K37').ClearContents() | Out-Null
    $detailStartRow = 3
    $detailQtyCol = 'F'; $detailPriceCol = 'G'; $detailAmountCol = 'H'; $detailNetCol = 'I'; $detailGrossCol = 'J'; $detailOriginCol1 = 'K'; $detailOriginCol2 = $null; $detailTotalRow = 37
  }

  $detailRow = $detailStartRow
  $sumQty = 0.0
  $sumAmount = 0.0
  $sumNet = 0.0
  $sumGross = 0.0
  foreach ($item in @($payload.items)) {
    $code = [string]$item.code
    if ([string]::IsNullOrWhiteSpace($code)) { continue }
    $meta = if ($itemMeta.ContainsKey($code)) { $itemMeta[$code] } else { [pscustomobject]@{ qty=0.0; price=0.0; amount=0.0; net=0.0; gross=0.0; specification=[string]$item.name } }
    Set-CellText $targetDetailWs ("B{0}" -f $detailRow) $code
    Set-CellText $targetDetailWs ("C{0}" -f $detailRow) 'DISPLAY FILM'
    Set-CellText $targetDetailWs ("D{0}" -f $detailRow) ([string]$meta.specification)
    Set-CellText $targetDetailWs ("E{0}" -f $detailRow) 'EA'
    Set-CellNumber $targetDetailWs ("{0}{1}" -f $detailQtyCol, $detailRow) $meta.qty
    Set-CellNumber $targetDetailWs ("{0}{1}" -f $detailPriceCol, $detailRow) $meta.price
    Set-CellNumber $targetDetailWs ("{0}{1}" -f $detailAmountCol, $detailRow) $meta.amount
    Set-CellText $targetDetailWs ("{0}{1}" -f $detailNetCol, $detailRow) (Format-WeightText $meta.net)
    Set-CellText $targetDetailWs ("{0}{1}" -f $detailGrossCol, $detailRow) (Format-WeightText $meta.gross)
    Set-CellText $targetDetailWs ("{0}{1}" -f $detailOriginCol1, $detailRow) 'Korea'
    if ($detailOriginCol2) { Set-CellText $targetDetailWs ("{0}{1}" -f $detailOriginCol2, $detailRow) 'Korea' }
    $sumQty += [double]$meta.qty
    $sumAmount += [double]$meta.amount
    $sumNet += [double]$meta.net
    $sumGross += [double]$meta.gross
    $detailRow++
  }
  Set-CellNumber $targetDetailWs ("{0}{1}" -f $detailQtyCol, $detailTotalRow) $sumQty
  Set-CellNumber $targetDetailWs ("{0}{1}" -f $detailAmountCol, $detailTotalRow) $sumAmount
  Set-CellNumber $targetDetailWs ("{0}{1}" -f $detailNetCol, $detailTotalRow) $sumNet
  Set-CellNumber $targetDetailWs ("{0}{1}" -f $detailGrossCol, $detailTotalRow) $sumGross

  $outerL = [string]$targetDetailPackingWs.Range('D4').Text
  $outerW = [string]$targetDetailPackingWs.Range('F4').Text
  $outerH = [string]$targetDetailPackingWs.Range('H4').Text
  $boxQtyText = [string]$targetDetailPackingWs.Range('I4').Text
  $cbmText = [string]$targetDetailPackingWs.Range('J4').Text
  $packingType = [string]$targetDetailPackingWs.Range('M4').Text
  if ([string]::IsNullOrWhiteSpace($outerL)) { $outerL = '45' }
  if ([string]::IsNullOrWhiteSpace($outerW)) { $outerW = '31' }
  if ([string]::IsNullOrWhiteSpace($outerH)) { $outerH = '30' }
  if ([string]::IsNullOrWhiteSpace($boxQtyText)) { $boxQtyText = '1' }
  if ([string]::IsNullOrWhiteSpace($cbmText)) { $cbmText = '0.042' }
  if ([string]::IsNullOrWhiteSpace($packingType)) { $packingType = 'box' }
  $targetDetailPackingWs.Range('A4:N40').ClearContents() | Out-Null

  $packingRow = 4
  $totalBoxNet = 0.0
  $totalBoxGross = 0.0
  foreach ($box in @($boxes)) {
    $boxNo = if ($box.boxNo) { [int]$box.boxNo } else { ($packingRow - 3) }
    $boxQty = 0.0
    $boxNet = 0.0
    $boxGross = 0.0
    foreach ($boxItem in @($box.items)) {
      $code = [string]$boxItem.code
      if (-not $itemMeta.ContainsKey($code)) { continue }
      $meta = $itemMeta[$code]
      $qtyValue = [double]$boxItem.qty
      $boxQty += $qtyValue
      if ([double]$meta.qty -gt 0) {
        $ratio = $qtyValue / [double]$meta.qty
        $boxNet += ([double]$meta.net * $ratio)
        $boxGross += ([double]$meta.gross * $ratio)
      }
    }
    Set-CellText $targetDetailPackingWs ("A{0}" -f $packingRow) ([string]$boxNo)
    Set-CellText $targetDetailPackingWs ("B{0}" -f $packingRow) 'DISPLAY FILM'
    Set-CellText $targetDetailPackingWs ("C{0}" -f $packingRow) ('{0:N0} EA' -f $boxQty)
    Set-CellText $targetDetailPackingWs ("D{0}" -f $packingRow) $outerL
    Set-CellText $targetDetailPackingWs ("E{0}" -f $packingRow) '*'
    Set-CellText $targetDetailPackingWs ("F{0}" -f $packingRow) $outerW
    Set-CellText $targetDetailPackingWs ("G{0}" -f $packingRow) '*'
    Set-CellText $targetDetailPackingWs ("H{0}" -f $packingRow) $outerH
    Set-CellText $targetDetailPackingWs ("I{0}" -f $packingRow) $boxQtyText
    Set-CellText $targetDetailPackingWs ("J{0}" -f $packingRow) $cbmText
    Set-CellText $targetDetailPackingWs ("K{0}" -f $packingRow) (('{0:0.##}' -f $boxNet) + ' kg')
    Set-CellText $targetDetailPackingWs ("L{0}" -f $packingRow) (('{0:0.##}' -f $boxGross) + ' kg')
    Set-CellText $targetDetailPackingWs ("M{0}" -f $packingRow) $packingType
    $totalBoxNet += $boxNet
    $totalBoxGross += $boxGross
    $packingRow++
  }
  Set-CellText $targetDetailPackingWs ("A{0}" -f $packingRow) 'TOTAL'
  Set-CellText $targetDetailPackingWs ("I{0}" -f $packingRow) ('{0} Box' -f @($boxes).Count)
  Set-CellText $targetDetailPackingWs ("J{0}" -f $packingRow) ('{0:0.###}' -f ([double]$cbmText * @($boxes).Count))
  Set-CellText $targetDetailPackingWs ("K{0}" -f $packingRow) (('{0:0.##}' -f $totalBoxNet) + ' kg')
  Set-CellText $targetDetailPackingWs ("L{0}" -f $packingRow) (('{0:0.##}' -f $totalBoxGross) + ' kg')

  $targetCntrWs.Range('A5:AP84').ClearContents() | Out-Null
  Set-CellText $targetCntrWs 'A3' 'CT NO : '
  Set-CellText $targetCntrWs 'S3' 'SEAL NO :'
  Set-MergedCellTextByPosition $targetCntrWs 5 1 ("DISPLAY FILM`n" + @($boxes).Count)
  $cntrRow = 5
  foreach ($item in @($payload.items)) {
    $code = [string]$item.code
    if ([string]::IsNullOrWhiteSpace($code)) { continue }
    $meta = if ($itemMeta.ContainsKey($code)) { $itemMeta[$code] } else { [pscustomobject]@{ qty=$item.qty; specification=[string]$item.name } }
    Set-CellText $targetCntrWs ("AK{0}" -f $cntrRow) '1'
    Set-CellText $targetCntrWs ("AL{0}" -f $cntrRow) $code
    Set-CellText $targetCntrWs ("AM{0}" -f $cntrRow) 'DISPLAY FILM'
    Set-CellText $targetCntrWs ("AN{0}" -f $cntrRow) ([string]$meta.specification)
    Set-CellText $targetCntrWs ("AO{0}" -f $cntrRow) 'EA'
    Set-CellNumber $targetCntrWs ("AP{0}" -f $cntrRow) $meta.qty
    $cntrRow++
  }

  $startCols = @(4,17,30,43,56,69)
  foreach ($col in $startCols) {
    Set-MergedCellTextByPosition $targetShipmarkWs 14 $col ''
    Set-MergedCellTextByPosition $targetShipmarkWs 16 $col ''
    Set-MergedCellTextByPosition $targetShipmarkWs 18 $col ''
    Set-MergedCellTextByPosition $targetShipmarkWs 20 $col ''
  }
  $boxIndex = 0
  foreach ($box in $boxes) {
    if ($boxIndex -ge $startCols.Count) { break }
    $codes = @($box.items | ForEach-Object { [string]$_.code } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    $commodity = if ($codes.Count -eq 0) { 'COMMODITY :' } elseif ($codes.Count -eq 1) { 'COMMODITY : ' + $codes[0] } else { 'COMMODITY : ' + $codes[0] + ', ' + $codes[1] + ' ETC.' }
    $col = $startCols[$boxIndex]
    Set-MergedCellTextByPosition $targetShipmarkWs 14 $col 'HO CHI MINH, VIETNAM'
    Set-MergedCellTextByPosition $targetShipmarkWs 16 $col $commodity
    Set-MergedCellTextByPosition $targetShipmarkWs 18 $col ('C/NO(CASE NUMBER) : CT NO. ' + [int]$box.boxNo)
    Set-MergedCellTextByPosition $targetShipmarkWs 20 $col 'MADE IN KOREA'
    $boxIndex++
  }

  if (Test-Path -LiteralPath $OutputPath) {
    Remove-Item -LiteralPath $OutputPath -Force
  }
  $targetWb.SaveCopyAs($OutputPath)

  Write-Output ("FILLED`t" + $OutputPath)
}
finally {
  if ($targetWb) { try { $targetWb.Close($false) } catch {} }
  if ($sourceWb) { try { $sourceWb.Close($false) } catch {} }
  if ($excel) { try { $excel.Quit() | Out-Null } catch {} }
  Release-ComObjectSafe $targetWb
  Release-ComObjectSafe $sourceWb
  Release-ComObjectSafe $excel
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
