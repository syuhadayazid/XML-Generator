$inputPath = "C:\Users\syuhada.yazid\OneDrive - WiseTech Global\Desktop\sample files\xmlpath.txt"
$outputPath = "C:\Users\syuhada.yazid\OneDrive - WiseTech Global\Desktop\XML Generator\sample.xml"
$ediOutputPath = [System.IO.Path]::ChangeExtension($outputPath, ".edi")
$sefSchemaPath = "C:\Users\syuhada.yazid\OneDrive - WiseTech Global\Desktop\sample files\X12\X12-861-4010.sef"
$NoMappingToken = '__NO_MAPPING__'
$xlsxPathColumn = "Element Xpath or Segment, Loop, Element Identifier"
$xlsxValueColumn = "Value"
$xlsxWorksheetName = $null
$outputDir = Split-Path -Parent $outputPath

if (-not (Test-Path $inputPath)) {
    Write-Error "Input file not found: $inputPath"
    exit 1
}

if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$lines = @()

function Normalize-PathForLookup {
    param([string]$line)

    $clean = ([string]$line).Trim()
    if ($clean -and -not $clean.StartsWith('/')) {
        $clean = '/' + $clean
    }
    while ($clean -match '(^|/)px:px:') {
        $clean = $clean -replace '(^|/)px:px:', '$1px:'
    }
    $clean = $clean -replace 'px:/Shipment', 'px:Shipment'
    $clean = $clean -replace 'px:/', 'px:'
    $clean = [regex]::Replace($clean, '(?<=[A-Za-z0-9_\]])(px:)(?=[A-Za-z_])', '/$1')
    $clean = [regex]::Replace($clean, '(\[[^\]]+\])(?=[A-Za-z_])', '$1/')
    return $clean
}

function Resolve-EdiMappedValue {
    param([string]$rawValue)

    $text = [string]$rawValue
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $text = $text -replace "`r", ""
    $text = $text -split "`n" | Select-Object -First 1
    $text = $text.Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    if ($text -match '(?i)hardcode\s*=?\s*["'']([^"'']+)["'']') {
        return $Matches[1]
    }

    if ($text -match '(?i)hardcode\s+to\s*["'']([^"'']+)["'']') {
        return $Matches[1].Trim()
    }

    if ($text -match '(?i)hardcode\s+to\s+([A-Za-z0-9_\-]+)') {
        return $Matches[1].Trim()
    }

    if ($text -match '(?i)hardcode\s*\?([^\?]+)\?') {
        return $Matches[1].Trim()
    }

    if ($text -match '(?i)^default\s+is\s*["'']([^"'']+)["'']') {
        return $Matches[1].Trim()
    }

    if ($text -match '(?i)^default\s+is\s+([A-Za-z0-9_\-]+)') {
        return $Matches[1].Trim()
    }

    if ($text -match '(?i)\bshould\s+be\s+([A-Za-z0-9_\-]+)\b') {
        return $Matches[1].Trim()
    }

    if ($text -match '\b(\d{4})[-/](\d{2})[-/](\d{2})\b') {
        return "$($Matches[1])$($Matches[2])$($Matches[3])"
    }

    if ($text -match '\b(\d{8})\b') {
        return $Matches[1]
    }

    if ($text -match '\b(\d{2})[-/](\d{2})[-/](\d{4})\b') {
        return "$($Matches[3])$($Matches[1])$($Matches[2])"
    }

    if ($text -match '^([A-Za-z0-9_\-]+)\s*=\s*.+$') {
        return $Matches[1]
    }

    if ($text -match '^([A-Za-z0-9_\-]+)\s*\([^\)]*\)$') {
        return $Matches[1]
    }

    if ($text -match '(?i)^ten\s+empty\s+spaces$') {
        return '          '
    }

    if ($text -match '(?i)^current\s+date') {
        return (Get-Date -Format 'yyyyMMdd')
    }

    if ($text -match '(?i)^current\s+time') {
        return (Get-Date -Format 'HHmm')
    }

    if ($text -match '(?i)sequential\s+number') {
        return '0001'
    }

    if ($text -match '^(<[^>]+>|\[[^\]]+\])$') {
        return 'SAMPLE_VALUE'
    }

    if ($text -match '(?i)^no\s+mapping$') {
        return $NoMappingToken
    }

    # If the cell looks like a guidance sentence, prefer fallback sample defaults.
    if ($text -match '(?i)(format|hardcode|current\s+date|current\s+time|expressed\s+as|for\s+example)') {
        return $null
    }

    return $text
}

function Convert-EdiToPathLines {
    param([string]$ediText)

    $result = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($ediText)) {
        return @($result)
    }

    $segmentParts = [regex]::Split($ediText, '~|\r?\n')
    foreach ($rawSegment in $segmentParts) {
        $segment = [string]$rawSegment
        if ([string]::IsNullOrWhiteSpace($segment)) {
            continue
        }

        $segment = $segment.Trim()
        if (-not $segment) {
            continue
        }

        $elements = $segment -split '\*'
        if (-not $elements -or $elements.Count -eq 0) {
            continue
        }

        $segmentId = ([string]$elements[0]).Trim()
        $segmentId = ($segmentId -replace '[^A-Za-z0-9_:\-]', '')
        if ([string]::IsNullOrWhiteSpace($segmentId)) {
            continue
        }

        for ($i = 1; $i -lt $elements.Count; $i++) {
            $value = ([string]$elements[$i]).Trim()
            if ([string]::IsNullOrWhiteSpace($value)) {
                continue
            }

            $elementName = "{0}{1:D2}" -f $segmentId, $i
            $result.Add("/X12/$segmentId/$elementName")
        }
    }

    return @($result)
}

function Test-LooksLikeEdiContent {
    param([string[]]$lines)

    if (-not $lines -or $lines.Count -eq 0) {
        return $false
    }

    $segments = New-Object System.Collections.Generic.List[string]
    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        $parts = [regex]::Split([string]$line, '~')
        foreach ($part in $parts) {
            $trimmed = ([string]$part).Trim()
            if (-not [string]::IsNullOrWhiteSpace($trimmed)) {
                $segments.Add($trimmed)
            }
        }
    }

    if ($segments.Count -lt 2) {
        return $false
    }

    $ediLike = 0
    foreach ($segment in $segments) {
        if ($segment -match '^[A-Za-z0-9]{2,3}\*') {
            $ediLike++
        }
    }

    return (($ediLike / [double]$segments.Count) -ge 0.6)
}

function Get-InputLines {
    param(
        [string]$sourcePath,
        [string]$pathColumn,
        [string]$valueColumn,
        [string]$worksheetName
    )

    $pathValues = @{}

    $ext = [System.IO.Path]::GetExtension($sourcePath).ToLowerInvariant()

    if ($ext -eq ".txt") {
        $text = Get-Content -Raw $sourcePath
        $rawLines = @($text -split '\r?\n')
        if (Test-LooksLikeEdiContent -lines $rawLines) {
            $converted = Convert-EdiToPathLines -ediText $text
            if ($converted -and $converted.Count -gt 0) {
                Write-Warning "Detected EDI content in TXT input. Converting segments to XML-like paths."
                return @{ Lines = @($converted); PathValues = $pathValues }
            }
        }

        return @{ Lines = @($rawLines); PathValues = $pathValues }
    }

    if ($ext -eq ".edi" -or $ext -eq ".x12") {
        $ediText = Get-Content -Raw $sourcePath
        $converted = Convert-EdiToPathLines -ediText $ediText
        if (-not $converted -or $converted.Count -eq 0) {
            Write-Error "No readable EDI segments found in '$sourcePath'."
            exit 1
        }

        return @{ Lines = @($converted); PathValues = $pathValues }
    }

    if ($ext -eq ".xlsx") {
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            Write-Error "XLSX input requires ImportExcel module. Install once with: Install-Module ImportExcel -Scope CurrentUser"
            exit 1
        }

        if (-not (Get-Module -Name ImportExcel)) {
            Import-Module ImportExcel -ErrorAction Stop
        }

        function Import-XlsxRows {
            param(
                [string]$xlsxPath,
                [string]$sheetName
            )

            try {
                if ([string]::IsNullOrWhiteSpace($sheetName)) {
                    return @{ Rows = @(Import-Excel -Path $xlsxPath); UsedNoHeader = $false }
                }

                return @{ Rows = @(Import-Excel -Path $xlsxPath -WorksheetName $sheetName); UsedNoHeader = $false }
            } catch {
                if ($_.Exception.Message -notmatch "No column headers found on top row") {
                    throw
                }

                if ([string]::IsNullOrWhiteSpace($sheetName)) {
                    return @{ Rows = @(Import-Excel -Path $xlsxPath -NoHeader); UsedNoHeader = $true }
                }

                return @{ Rows = @(Import-Excel -Path $xlsxPath -WorksheetName $sheetName -NoHeader); UsedNoHeader = $true }
            }
        }

        function Find-EmbeddedHeaderValues {
            param(
                [string]$xlsxPath,
                [string]$sheetName,
                [scriptblock]$normalizeFn,
                [string]$normalizedHeader
            )

            $searchPath = $xlsxPath
            $tempSearchPath = $null

            try {
                if ([string]::IsNullOrWhiteSpace($sheetName)) {
                    $null = Get-ExcelSheetInfo -Path $searchPath
                }
            } catch {
                $tempSearchPath = Join-Path ([System.IO.Path]::GetTempPath()) ("xmlgen-header-" + [System.Guid]::NewGuid().ToString() + ".xlsx")
                Copy-Item -LiteralPath $xlsxPath -Destination $tempSearchPath -Force -ErrorAction Stop
                $searchPath = $tempSearchPath
            }

            $sheetNames = @()
            if ([string]::IsNullOrWhiteSpace($sheetName)) {
                $sheetNames = @((Get-ExcelSheetInfo -Path $searchPath).Name)
            } else {
                $sheetNames = @($sheetName)
            }

            try {
                $bestMatch = $null
                foreach ($candidateSheet in $sheetNames) {
                    try {
                        $rawRows = @(Import-Excel -Path $searchPath -WorksheetName $candidateSheet -NoHeader)
                    } catch {
                        continue
                    }

                    if (-not $rawRows -or $rawRows.Count -eq 0) {
                        continue
                    }

                    $props = @($rawRows[0].PSObject.Properties.Name)
                    $matches = New-Object System.Collections.Generic.List[object]
                    for ($rowIndex = 0; $rowIndex -lt $rawRows.Count; $rowIndex++) {
                        foreach ($prop in $props) {
                            $cellValue = [string]$rawRows[$rowIndex].$prop
                            if ((& $normalizeFn $cellValue) -eq $normalizedHeader) {
                                $matches.Add([pscustomobject]@{ RowIndex = $rowIndex; Column = $prop })
                            }
                        }
                    }

                    foreach ($match in $matches) {
                        $values = New-Object System.Collections.Generic.List[string]
                        $pathLikeCount = 0
                        for ($dataIndex = ($match.RowIndex + 1); $dataIndex -lt $rawRows.Count; $dataIndex++) {
                            $value = [string]$rawRows[$dataIndex].$($match.Column)
                            if ([string]::IsNullOrWhiteSpace($value)) {
                                continue
                            }

                            $trimmed = $value.Trim()
                            $values.Add($trimmed)
                            if ($trimmed -match '^\s*/?[A-Za-z_][\w:\-]*(/|\[)') {
                                $pathLikeCount++
                            }
                        }

                        if ($values.Count -eq 0) {
                            continue
                        }

                        $candidate = [pscustomobject]@{
                            Sheet = $candidateSheet
                            RowNumber = $match.RowIndex + 1
                            Column = $match.Column
                            Values = @($values)
                            PathLikeCount = $pathLikeCount
                            Score = ($pathLikeCount * 2) + $values.Count
                        }

                        if (-not $bestMatch -or $candidate.Score -gt $bestMatch.Score) {
                            $bestMatch = $candidate
                        }
                    }
                }

                return $bestMatch
            } finally {
                if ($tempSearchPath -and (Test-Path -LiteralPath $tempSearchPath)) {
                    Remove-Item -LiteralPath $tempSearchPath -Force -ErrorAction SilentlyContinue
                }
            }
        }

        $rows = $null
        $directReadError = $null
        $usedNoHeader = $false
        try {
            $importResult = Import-XlsxRows -xlsxPath $sourcePath -sheetName $worksheetName
            $rows = @($importResult.Rows)
            $usedNoHeader = [bool]$importResult.UsedNoHeader
        } catch {
            $directReadError = $_.Exception.Message
        }

        if (-not $rows -or $rows.Count -eq 0) {
            $tempXlsxPath = Join-Path ([System.IO.Path]::GetTempPath()) ("xmlgen-" + [System.Guid]::NewGuid().ToString() + ".xlsx")
            $tempReadError = $null

            try {
                Copy-Item -LiteralPath $sourcePath -Destination $tempXlsxPath -Force -ErrorAction Stop
                $importResult = Import-XlsxRows -xlsxPath $tempXlsxPath -sheetName $worksheetName
                $rows = @($importResult.Rows)
                $usedNoHeader = [bool]$importResult.UsedNoHeader
            } catch {
                $tempReadError = $_.Exception.Message
            } finally {
                if (Test-Path -LiteralPath $tempXlsxPath) {
                    Remove-Item -LiteralPath $tempXlsxPath -Force -ErrorAction SilentlyContinue
                }
            }

            if (-not $rows -or $rows.Count -eq 0) {
                $detail = "direct read failed"
                if (-not [string]::IsNullOrWhiteSpace($directReadError)) {
                    $detail += ": $directReadError"
                }
                if (-not [string]::IsNullOrWhiteSpace($tempReadError)) {
                    $detail += " | temp-copy read failed: $tempReadError"
                }

                Write-Error "Unable to read XLSX '$sourcePath'. Close the file in Excel or wait for OneDrive sync, then retry. Details: $detail"
                exit 1
            }
        }

        if (-not $rows -or $rows.Count -eq 0) {
            Write-Error "No rows found in XLSX '$sourcePath'."
            exit 1
        }

        $firstRowProps = @($rows[0].PSObject.Properties.Name)
        $normalize = {
            param([string]$text)
            return (($text -replace '\s+', '').Trim()).ToLowerInvariant()
        }

        $targetNormalized = & $normalize $pathColumn
        $targetValueNormalized = & $normalize $valueColumn
        $resolvedColumn = $null
        $resolvedValueColumn = $null
        foreach ($prop in $firstRowProps) {
            if ((& $normalize $prop) -eq $targetNormalized) {
                $resolvedColumn = $prop
            }

            if ((& $normalize $prop) -eq $targetValueNormalized) {
                $resolvedValueColumn = $prop
            }
        }

        if (-not $resolvedColumn) {
            $embeddedHeaderMatch = Find-EmbeddedHeaderValues -xlsxPath $sourcePath -sheetName $worksheetName -normalizeFn $normalize -normalizedHeader $targetNormalized
            if ($embeddedHeaderMatch) {
                Write-Warning "Detected embedded header '$pathColumn' on worksheet '$($embeddedHeaderMatch.Sheet)' row $($embeddedHeaderMatch.RowNumber), column '$($embeddedHeaderMatch.Column)'."
                $embeddedLines = @($embeddedHeaderMatch.Values | Where-Object { $_ -match '^\s*/?[A-Za-z_][\w:\-]*(/|\[)' })
                return @{ Lines = $embeddedLines; PathValues = $pathValues }
            }
        }

        if (-not $resolvedColumn -and $usedNoHeader) {
            $bestColumn = $null
            $bestScore = -1

            foreach ($prop in $firstRowProps) {
                $nonEmptyCount = 0
                $pathLikeCount = 0
                foreach ($row in $rows) {
                    $value = [string]$row.$prop
                    if ([string]::IsNullOrWhiteSpace($value)) {
                        continue
                    }

                    $nonEmptyCount++
                    if ($value -match "^\s*/?[A-Za-z_][\w:\-]*(/|\[)") {
                        $pathLikeCount++
                    }
                }

                $score = ($pathLikeCount * 2) + $nonEmptyCount
                if ($score -gt $bestScore) {
                    $bestScore = $score
                    $bestColumn = $prop
                }
            }

            if ($bestColumn) {
                $resolvedColumn = $bestColumn
                Write-Warning "No header row detected in XLSX. Using detected path column '$resolvedColumn'."
            }
        }

        if (-not $resolvedColumn) {
            $available = ($firstRowProps -join ', ')
            Write-Error "Column '$pathColumn' not found in XLSX input. Available columns: $available"
            exit 1
        }

        if (-not $resolvedValueColumn) {
            $candidateNames = @('value', 'mapped value', 'output value', 'target value', 'populate value', 'edi value')
            foreach ($prop in $firstRowProps) {
                $normalizedProp = & $normalize $prop
                if ($candidateNames -contains $normalizedProp) {
                    $resolvedValueColumn = $prop
                    break
                }
            }
        }

        if (-not $resolvedValueColumn) {
            Write-Warning "Value column '$valueColumn' not found. EDI output will use sample placeholder values for missing mappings."
        }

        $result = New-Object System.Collections.Generic.List[string]
        foreach ($row in $rows) {
            $pathValue = [string]$row.$resolvedColumn
            if ([string]::IsNullOrWhiteSpace($pathValue)) {
                continue
            }

            $trimmedPath = $pathValue.Trim()
            if ($trimmedPath -notmatch '^\s*/?[A-Za-z_][\w:\-]*(/|\[)') {
                continue
            }

            $result.Add($trimmedPath)

            if ($resolvedValueColumn) {
                $mappedValue = Resolve-EdiMappedValue -rawValue ([string]$row.$resolvedValueColumn)
                if (-not [string]::IsNullOrWhiteSpace($mappedValue)) {
                    $lookupKey = Normalize-PathForLookup -line $trimmedPath
                    if (-not $pathValues.ContainsKey($lookupKey)) {
                        $pathValues[$lookupKey] = New-Object System.Collections.Queue
                    }

                    if ($pathValues[$lookupKey] -is [System.Collections.Queue]) {
                        $pathValues[$lookupKey].Enqueue(([string]$mappedValue).Trim())
                    }
                }
            }
        }

        if (Test-LooksLikeEdiContent -lines @($result)) {
            $converted = Convert-EdiToPathLines -ediText (($result -join [Environment]::NewLine))
            if ($converted -and $converted.Count -gt 0) {
                Write-Warning "Detected EDI content in XLSX column '$resolvedColumn'. Converting segments to XML-like paths."
                return @{ Lines = @($converted); PathValues = $pathValues }
            }
        }

        return @{ Lines = @($result); PathValues = $pathValues }
    }

    Write-Error "Unsupported input file type '$ext'. Supported: .txt, .xlsx, .edi, .x12"
    exit 1
}

$inputResult = Get-InputLines -sourcePath $inputPath -pathColumn $xlsxPathColumn -valueColumn $xlsxValueColumn -worksheetName $xlsxWorksheetName
$lines = @($inputResult.Lines)
$pathValueMap = @{}
if ($inputResult.PathValues) {
    $pathValueMap = $inputResult.PathValues
}
$used = 0
$skipped = 0
$skippedDetails = New-Object System.Collections.Generic.List[string]
$malformed = New-Object System.Collections.Generic.List[string]

$xmlDoc = New-Object System.Xml.XmlDocument
$ns = "urn:px"
$root = $null

function Split-XmlPathSegments {
    param([string]$path)

    $segments = New-Object System.Collections.Generic.List[string]
    $sb = New-Object System.Text.StringBuilder
    $depth = 0

    for ($i = 0; $i -lt $path.Length; $i++) {
        $ch = $path[$i]

        if ($ch -eq '[') {
            $depth++
            $null = $sb.Append($ch)
            continue
        }

        if ($ch -eq ']') {
            $depth--
            $null = $sb.Append($ch)
            continue
        }

        if ($ch -eq '/' -and $depth -eq 0) {
            $part = $sb.ToString().Trim()
            if (-not [string]::IsNullOrWhiteSpace($part)) {
                $segments.Add($part)
            }
            $null = $sb.Clear()
            continue
        }

        $null = $sb.Append($ch)
    }

    $tail = $sb.ToString().Trim()
    if (-not [string]::IsNullOrWhiteSpace($tail)) {
        $segments.Add($tail)
    }

    if ($depth -ne 0) {
        return @{ Valid = $false; Segments = @(); Error = "unbalanced brackets" }
    }

    return @{ Valid = $true; Segments = $segments; Error = "" }
}

function Normalize-PathLine {
    param([string]$line)

    $clean = $line.Trim()
    while ($clean -match '(^|/)px:px:') {
        $clean = $clean -replace '(^|/)px:px:', '$1px:'
    }
    $clean = $clean -replace 'px:/Shipment', 'px:Shipment'
    $clean = $clean -replace 'px:/', 'px:'
    $clean = [regex]::Replace($clean, '(?<=[A-Za-z0-9_\]])(px:)(?=[A-Za-z_])', '/$1')
    $clean = [regex]::Replace($clean, '(\[[^\]]+\])(?=[A-Za-z_])', '$1/')
    return $clean
}

function Parse-Segment {
    param(
        [string]$segment,
        [int]$lineNo
    )

    $result = @{
        Valid = $false
        Name = ""
        Attributes = @{}
        IsAttributeOnly = $false
        Error = ""
    }

    if (-not ($segment -match '^([^\[]+)(\[(.*)\])?$')) {
        $result.Error = "Line ${lineNo}: malformed segment '$segment'"
        return $result
    }

    $name = $Matches[1].Trim()
    while ($name -match '(^|:)px:px:') {
        $name = $name -replace '(^|:)px:px:', '$1px:'
    }
    $predicate = $Matches[3]

        if ($name -match '^@([\w:.-]+)$') {
        $result.Valid = $true
        $result.Name = $name
        $result.IsAttributeOnly = $true
        return $result
    }

    if ([string]::IsNullOrWhiteSpace($name) -or ($name -match '[\[\]@=\s]')) {
        $result.Error = "Line ${lineNo}: invalid element name '$name'"
        return $result
    }

    $attrs = @{}
    if ($predicate) {
        $quoteCount = ([regex]::Matches($predicate, '"')).Count
        if (($quoteCount % 2) -ne 0) {
            $result.Error = "Line ${lineNo}: malformed predicate quotes in '$segment'"
            return $result
        }

        # Ignore path predicates that are not element attributes, e.g. [/px:@typeCode=...]
        if ($predicate -notmatch '/\s*px:') {
            $attrMatches = [regex]::Matches($predicate, '@([\w:]+)\s*=\s*(["''])(.*?)\2')
            foreach ($m in $attrMatches) {
                $attrName = $m.Groups[1].Value
                $attrs[$attrName] = $m.Groups[3].Value
            }

            # Support attribute-existence predicates like [@unitCode] by assigning a sample value.
            $trimmedPredicate = $predicate.Trim()
            if ($trimmedPredicate -match '^@([\w:]+)$') {
                $attrName = $Matches[1]
                if (-not $attrs.ContainsKey($attrName)) {
                    $attrs[$attrName] = 'SAMPLE_VALUE'
                }
            }

            if (($predicate -match '@') -and $attrs.Count -eq 0) {
                $result.Error = "Line ${lineNo}: unsupported/malformed predicate '$predicate'"
                return $result
            }
        }
    }

    $result.Valid = $true
    $result.Name = $name
    $result.Attributes = $attrs
    return $result
}

function Get-LocalName {
    param([string]$name)

    if ($name -match ':') {
        return $name.Split(':')[-1]
    }

    return $name
}

function Get-SampleEdiElementValue {
    param(
        [string]$segmentId,
        [int]$position,
        [string]$transactionSet
    )

    $segment = $segmentId.ToUpperInvariant()
    switch ($segment) {
        'ISA' {
            switch ($position) {
                1 { return '00' }
                2 { return '          ' }
                3 { return '00' }
                4 { return '          ' }
                5 { return 'ZZ' }
                6 { return 'SENDERID      ' }
                7 { return 'ZZ' }
                8 { return 'RECEIVERID    ' }
                9 { return (Get-Date -Format 'yyMMdd') }
                10 { return (Get-Date -Format 'HHmm') }
                11 { return 'U' }
                12 { return '00401' }
                13 { return '000000001' }
                14 { return '0' }
                15 { return 'T' }
                16 { return ':' }
                default { return 'SAMPLE_VALUE' }
            }
        }
        'GS' {
            switch ($position) {
                1 {
                    switch ($transactionSet) {
                        '214' { return 'QM' }
                        default { return 'OW' }
                    }
                }
                2 { return 'SENDER' }
                3 { return 'RECEIVER' }
                4 { return (Get-Date -Format 'yyyyMMdd') }
                5 { return (Get-Date -Format 'HHmm') }
                6 { return '1' }
                7 { return 'X' }
                8 { return '004010' }
                default { return 'SAMPLE_VALUE' }
            }
        }
        'ST' {
            if ($position -eq 1) { return $transactionSet }
            if ($position -eq 2) { return '0001' }
            return 'SAMPLE_VALUE'
        }
        'SE' {
            if ($position -eq 1) { return '0000' }
            if ($position -eq 2) { return '0001' }
            return 'SAMPLE_VALUE'
        }
        'GE' {
            if ($position -eq 1) { return '1' }
            if ($position -eq 2) { return '1' }
            return 'SAMPLE_VALUE'
        }
        'IEA' {
            if ($position -eq 1) { return '1' }
            if ($position -eq 2) { return '000000001' }
            return 'SAMPLE_VALUE'
        }
        'DTM' {
            if ($position -eq 1) { return '' }
            if ($position -eq 2) { return (Get-Date -Format 'yyyyMMdd') }
            if ($position -eq 3) { return (Get-Date -Format 'HHmm') }
            return 'SAMPLE_VALUE'
        }
        default {
            if ($position -eq 1 -and $segment -match '^N[1-4]$') {
                return 'SAMPLE_VALUE'
            }
            return 'SAMPLE_VALUE'
        }
    }
}

function Format-EdiElementValue {
    param(
        [string]$segmentId,
        [int]$position,
        [string]$value
    )

    $text = [string]$value

    # X12 ISA sender/receiver IDs are fixed-width 15 characters.
    if ($segmentId -eq 'ISA' -and ($position -eq 6 -or $position -eq 8)) {
        if ($text.Length -lt 15) {
            return $text.PadRight(15)
        }

        return $text.Substring(0, [Math]::Min(15, $text.Length))
    }

    if ($segmentId -eq 'ISA' -and $position -eq 9) {
        if ($text -match '^\d{8}$') {
            return $text.Substring(2, 6)
        }

        if ($text -match '^\d{6}$') {
            return $text
        }

        return (Get-Date -Format 'yyMMdd')
    }

    if ($segmentId -eq 'ISA' -and $position -eq 10) {
        if ($text -match '^\d{4}$') {
            return $text
        }

        return (Get-Date -Format 'HHmm')
    }

    if ($segmentId -eq 'ISA' -and $position -eq 13) {
        $digits = ($text -replace '\D', '')
        if ([string]::IsNullOrWhiteSpace($digits)) {
            return '000000001'
        }

        if ($digits.Length -lt 9) {
            return $digits.PadLeft(9, '0')
        }

        if ($digits.Length -gt 9) {
            return $digits.Substring($digits.Length - 9)
        }

        return $digits
    }

    return $text
}

function Get-TransactionSetHintFromInputPath {
    param([string]$path)

    if ([string]::IsNullOrWhiteSpace($path)) {
        return $null
    }

    $name = [System.IO.Path]::GetFileNameWithoutExtension($path)
    if ([string]::IsNullOrWhiteSpace($name)) {
        return $null
    }

    $m = [regex]::Match($name, '(?i)x12[_\-]?(\d{3})')
    if ($m.Success) {
        return $m.Groups[1].Value
    }

    return $null
}

function Resolve-TransactionSetId {
    param([string]$transactionSetHint)

    $candidate = ([string]$transactionSetHint).Trim()
    if ($candidate -match '^\d{3}$') {
        return $candidate
    }

    return '861'
}

function Resolve-SefSchemaPath {
    param(
        [string]$transactionSetId,
        [string]$preferredPath
    )

    if (-not [string]::IsNullOrWhiteSpace($preferredPath) -and (Test-Path -LiteralPath $preferredPath)) {
        $preferredName = [System.IO.Path]::GetFileName($preferredPath)
        if ([string]::IsNullOrWhiteSpace($transactionSetId) -or $preferredName -match "(?i)$transactionSetId") {
            return $preferredPath
        }
    }

    $preferredDir = if ([string]::IsNullOrWhiteSpace($preferredPath)) { $null } else { [System.IO.Path]::GetDirectoryName($preferredPath) }
    $candidatePaths = @(
        (if ($preferredDir) { Join-Path $preferredDir ("X12-$transactionSetId-4010.sef") } else { $null }),
        (Join-Path 'C:\Users\syuhada.yazid\OneDrive - WiseTech Global\Desktop\sample files\X12' ("X12-$transactionSetId-4010.sef")),
        (Join-Path 'C:\Users\syuhada.yazid\OneDrive - WiseTech Global\Desktop\sample files' ("X12-$transactionSetId-4010.sef"))
    )

    foreach ($candidatePath in $candidatePaths) {
        if (-not [string]::IsNullOrWhiteSpace($candidatePath) -and (Test-Path -LiteralPath $candidatePath)) {
            return $candidatePath
        }
    }

    return $preferredPath
}

function Parse-SefSetTokens {
    param([string]$setExpression)

    $tokens = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($setExpression)) {
        return [object[]]$tokens.ToArray()
    }

    $depth = 0
    for ($index = 0; $index -lt $setExpression.Length; $index++) {
        $char = $setExpression[$index]
        if ($char -eq '{') {
            $depth++
            continue
        }

        if ($char -eq '}') {
            if ($depth -gt 0) { $depth-- }
            continue
        }

        if ($char -ne '[' -or $depth -gt 0) {
            continue
        }

        $endIndex = $setExpression.IndexOf(']', $index)
        if ($endIndex -lt 0) {
            break
        }

        $content = $setExpression.Substring($index + 1, $endIndex - $index - 1)
        $parts = @($content.Split(','))
        if ($parts.Count -eq 0) {
            $index = $endIndex
            continue
        }

        $segmentField = ([string]$parts[0]).Trim()
        if ([string]::IsNullOrWhiteSpace($segmentField)) {
            continue
        }

        $segmentId = ($segmentField.Split('@')[0]).Trim().ToUpperInvariant()
        if ($segmentId -notmatch '^[A-Z0-9]{2,3}$') {
            continue
        }

        $usage = if ($parts.Count -ge 2) { ([string]$parts[1]).Trim().ToUpperInvariant() } else { '' }
        $maxUseText = if ($parts.Count -ge 3) { ([string]$parts[2]).Trim() } else { '' }
        $maxUse = 1
        if ($maxUseText -eq '>1') {
            $maxUse = 999999
        } elseif ($maxUseText -match '^\d+$') {
            $maxUse = [int]$maxUseText
        }

        $tokens.Add([pscustomobject]@{
            SegmentId = $segmentId
            IsRequired = ($usage -eq 'M')
            MaxUse = $maxUse
        })

        $index = $endIndex
    }

    return [object[]]$tokens.ToArray()
}

function Parse-SefSegmentDefinitions {
    param([string[]]$lines)

    $definitions = @{}
    foreach ($line in $lines) {
        $text = ([string]$line).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) {
            continue
        }

        $equalsIndex = $text.IndexOf('=')
        if ($equalsIndex -lt 1) {
            continue
        }

        $segmentId = $text.Substring(0, $equalsIndex).Trim().ToUpperInvariant()
        if ($segmentId -notmatch '^[A-Z0-9]{2,3}$') {
            continue
        }

        $body = $text.Substring($equalsIndex + 1)
        $matches = [regex]::Matches($body, '\[([^\]]+)\]')
        $requiredPositions = New-Object System.Collections.Generic.List[int]
        $elementIds = New-Object System.Collections.Generic.List[string]
        $position = 0

        foreach ($match in $matches) {
            $position++
            $fields = @($match.Groups[1].Value.Split(','))
            $elementId = if ($fields.Count -ge 1) { ([string]$fields[0]).Trim() } else { '' }
            $elementIds.Add($elementId)
            $usage = if ($fields.Count -ge 2) { ([string]$fields[1]).Trim().ToUpperInvariant() } else { '' }
            if ($usage -eq 'M') {
                $requiredPositions.Add($position)
            }
        }

        $definitions[$segmentId] = [pscustomobject]@{
            SegmentId = $segmentId
            ElementCount = $position
            RequiredPositions = @($requiredPositions)
            ElementIds = @($elementIds)
        }
    }

    return $definitions
}

function Parse-SefElementDefinitions {
    param([string[]]$lines)

    $definitions = @{}
    foreach ($line in $lines) {
        $text = ([string]$line).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) {
            continue
        }

        if ($text -match '^(\d+)=([A-Z]{1,3}),(\d+),(\d+)\s*$') {
            $elementId = [string]$Matches[1]
            if (-not $definitions.ContainsKey($elementId)) {
                $definitions[$elementId] = [pscustomobject]@{
                    ElementId = $elementId
                    DataType = [string]$Matches[2]
                    MinLength = [int]$Matches[3]
                    MaxLength = [int]$Matches[4]
                }
            }
        }
    }

    return $definitions
}

function Get-SefSchemaModel {
    param(
        [string]$schemaPath,
        [string]$transactionSetId
    )

    if ([string]::IsNullOrWhiteSpace($schemaPath) -or -not (Test-Path -LiteralPath $schemaPath)) {
        return $null
    }

    $lines = Get-Content -LiteralPath $schemaPath
    $setExpression = $null
    $segmentLines = New-Object System.Collections.Generic.List[string]
    $elementLines = New-Object System.Collections.Generic.List[string]
    $inSets = $false
    $inSegs = $false
    $inElms = $false

    foreach ($rawLine in $lines) {
        $line = [string]$rawLine
        if ($line -match '^\.SETS') {
            $inSets = $true
            $inSegs = $false
            $inElms = $false
            continue
        }

        if ($line -match '^\.SEGS') {
            $inSets = $false
            $inSegs = $true
            $inElms = $false
            continue
        }

        if ($line -match '^\.ELMS') {
            $inSets = $false
            $inSegs = $false
            $inElms = $true
            continue
        }

        if ($line -match '^\.[A-Z]') {
            $inSets = $false
            $inSegs = $false
            $inElms = $false
            continue
        }

        if ($inSets -and $line -match '^(\d{3})=') {
            if ($Matches[1] -eq $transactionSetId) {
                $setExpression = $line.Substring(4)
            }
            continue
        }

        if ($inSegs) {
            $segmentLines.Add($line)
            continue
        }

        if ($inElms) {
            $elementLines.Add($line)
        }
    }

    if ([string]::IsNullOrWhiteSpace($setExpression)) {
        return $null
    }

    return [pscustomobject]@{
        SchemaPath = $schemaPath
        SetTokens = @(Parse-SefSetTokens -setExpression $setExpression)
        SegmentDefinitions = Parse-SefSegmentDefinitions -lines @($segmentLines)
        ElementDefinitions = Parse-SefElementDefinitions -lines @($elementLines)
    }
}

function Parse-EdiSegments {
    param([string]$ediText)

    $segments = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($ediText)) {
        return [object[]]$segments.ToArray()
    }

    $rawSegments = [regex]::Split($ediText, '~\s*')
    foreach ($rawSegment in $rawSegments) {
        $segmentText = ([string]$rawSegment).Trim()
        if ([string]::IsNullOrWhiteSpace($segmentText)) {
            continue
        }

        $parts = @([regex]::Split($segmentText, '\*'))
        $segmentId = ([string]$parts[0]).Trim().ToUpperInvariant()
        if ($segmentId -notmatch '^[A-Z0-9]{2,3}$') {
            continue
        }

        $elements = @()
        if ($parts.Count -gt 1) {
            $elements = @($parts[1..($parts.Count - 1)])
        }

        $segments.Add([pscustomobject]@{
            SegmentId = $segmentId
            Elements = $elements
            RawText = $segmentText
        })
    }

    return [object[]]$segments.ToArray()
}

function Validate-EdiAgainstSef {
    param(
        [string]$ediText,
        [object]$schemaModel,
        [string]$transactionSetId
    )

    $errors = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]

    if (-not $schemaModel) {
        $warnings.Add('SEF schema not available. Skipping EDI validation.')
        return [pscustomobject]@{ Errors = @($errors); Warnings = @($warnings) }
    }

    $segments = @(Parse-EdiSegments -ediText $ediText)
    if ($segments.Count -eq 0) {
        $errors.Add('Generated EDI contains no readable segments.')
        return [pscustomobject]@{ Errors = @($errors); Warnings = @($warnings) }
    }

    $setTokens = @($schemaModel.SetTokens)
    $segmentDefinitions = $schemaModel.SegmentDefinitions
    $elementDefinitions = if ($schemaModel.PSObject.Properties.Name -contains 'ElementDefinitions') { $schemaModel.ElementDefinitions } else { @{} }
    $tokenCounts = @{}
    $patternIndex = 0
    $topLevelSegmentIds = @{}
    foreach ($token in $setTokens) {
        $topLevelSegmentIds[$token.SegmentId] = $true
    }

    for ($segmentIndex = 0; $segmentIndex -lt $segments.Count; $segmentIndex++) {
        $segment = $segments[$segmentIndex]
        $lineNo = $segmentIndex + 1

        if ($segmentDefinitions.ContainsKey($segment.SegmentId)) {
            $definition = $segmentDefinitions[$segment.SegmentId]
            if ($segment.Elements.Count -gt $definition.ElementCount) {
                $errors.Add("Segment $($segment.SegmentId) at position $lineNo has $($segment.Elements.Count) elements, exceeding SEF definition of $($definition.ElementCount).")
            }

            foreach ($requiredPosition in $definition.RequiredPositions) {
                $value = if ($segment.Elements.Count -ge $requiredPosition) { [string]$segment.Elements[$requiredPosition - 1] } else { $null }
                if ($null -eq $value -or $value.Length -eq 0) {
                    $errors.Add("Segment $($segment.SegmentId) at position $lineNo is missing required element $($segment.SegmentId)$('{0:D2}' -f $requiredPosition).")
                }
            }

            for ($elementPos = 1; $elementPos -le $segment.Elements.Count; $elementPos++) {
                if ($definition.ElementIds.Count -lt $elementPos) {
                    continue
                }

                $elementId = [string]$definition.ElementIds[$elementPos - 1]
                if ([string]::IsNullOrWhiteSpace($elementId) -or -not $elementDefinitions.ContainsKey($elementId)) {
                    continue
                }

                $elementRule = $elementDefinitions[$elementId]
                $value = [string]$segment.Elements[$elementPos - 1]
                if ([string]::IsNullOrEmpty($value)) {
                    continue
                }

                $actualLength = $value.Length
                $minLength = [int]$elementRule.MinLength
                $maxLength = [int]$elementRule.MaxLength
                $elementName = "$($segment.SegmentId)$('{0:D2}' -f $elementPos)"

                if ($maxLength -gt 0 -and $actualLength -gt $maxLength) {
                    $errors.Add("${elementName}: element too long, actual: $actualLength, standard: $maxLength")
                    continue
                }

                if ($minLength -gt 0 -and $actualLength -lt $minLength) {
                    $errors.Add("${elementName}: element too short, actual: $actualLength, standard: $minLength")
                }
            }
        } else {
            $warnings.Add("Segment $($segment.SegmentId) at position $lineNo is not defined in the loaded SEF.")
        }

        if ($topLevelSegmentIds.ContainsKey($segment.SegmentId)) {
            $matched = $false
            for ($candidateIndex = $patternIndex; $candidateIndex -lt $setTokens.Count; $candidateIndex++) {
                $candidate = $setTokens[$candidateIndex]
                if ($candidate.SegmentId -ne $segment.SegmentId) {
                    continue
                }

                for ($skippedIndex = $patternIndex; $skippedIndex -lt $candidateIndex; $skippedIndex++) {
                    $skipped = $setTokens[$skippedIndex]
                    $skippedCount = if ($tokenCounts.ContainsKey($skippedIndex)) { [int]$tokenCounts[$skippedIndex] } else { 0 }
                    if ($skipped.IsRequired -and $skippedCount -eq 0) {
                        $warnings.Add("Segment $($segment.SegmentId) at position $lineNo appears before required segment $($skipped.SegmentId) in the SEF sequence.")
                    }
                }

                $currentCount = if ($tokenCounts.ContainsKey($candidateIndex)) { [int]$tokenCounts[$candidateIndex] } else { 0 }
                if ($currentCount -ge [int]$candidate.MaxUse) {
                    continue
                }

                $tokenCounts[$candidateIndex] = $currentCount + 1
                $patternIndex = $candidateIndex
                $matched = $true
                break
            }

        }
    }

    for ($index = 0; $index -lt $setTokens.Count; $index++) {
        $token = $setTokens[$index]
        $count = if ($tokenCounts.ContainsKey($index)) { [int]$tokenCounts[$index] } else { 0 }
        if ($token.IsRequired -and $count -eq 0) {
            $errors.Add("Required segment $($token.SegmentId) is missing according to the SEF $transactionSetId definition.")
        }
    }

    return [pscustomobject]@{
        Errors = @($errors)
        Warnings = @($warnings)
    }
}

function Build-SampleEdiFromPathLines {
    param(
        [string[]]$pathLines,
        [string]$transactionSetHint,
        [hashtable]$pathValueMap
    )

    $controlSegments = @('ISA', 'GS', 'ST', 'SE', 'GE', 'IEA')
    $controlElementMaps = @{}
    $segmentOccurrences = @{}
    $occurrenceOrder = New-Object System.Collections.Generic.List[string]
    $occurrenceCounters = @{}
    $latestOccurrenceByBase = @{}
    $effectiveTransactionSetHint = $transactionSetHint

    foreach ($line in $pathLines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        $clean = Normalize-PathLine -line $line
        $lineValue = $null
        $lookupKey = Normalize-PathForLookup -line $clean
        if ($pathValueMap -and $pathValueMap.ContainsKey($lookupKey)) {
            $mappedEntry = $pathValueMap[$lookupKey]
            if ($mappedEntry -is [System.Collections.Queue]) {
                if ($mappedEntry.Count -gt 0) {
                    $lineValue = [string]$mappedEntry.Dequeue()
                }
            } elseif ($mappedEntry -is [System.Collections.IList]) {
                if ($mappedEntry.Count -gt 0) {
                    $lineValue = [string]$mappedEntry[0]
                }
            } else {
                $lineValue = [string]$mappedEntry
            }
        }
        $split = Split-XmlPathSegments -path $clean
        if (-not $split.Valid -or $split.Segments.Count -lt 3) {
            continue
        }

        $parsed = New-Object System.Collections.Generic.List[hashtable]
        $lineMalformed = $false
        foreach ($seg in $split.Segments) {
            $parsedSeg = Parse-Segment -segment $seg -lineNo 0
            if (-not $parsedSeg.Valid -or $parsedSeg.IsAttributeOnly) {
                $lineMalformed = $true
                break
            }
            $parsed.Add($parsedSeg)
        }

        if ($lineMalformed -or $parsed.Count -lt 3) {
            continue
        }

        $rootName = Get-LocalName -name $parsed[0].Name
        if ($rootName -ne 'X12') {
            continue
        }

        $segmentIndex = 1
        $tsNodeName = Get-LocalName -name $parsed[1].Name
        if ($parsed.Count -ge 4 -and $tsNodeName -match '^TS_(\d{3})$') {
            $segmentIndex = 2
            if ([string]::IsNullOrWhiteSpace($effectiveTransactionSetHint)) {
                $effectiveTransactionSetHint = $Matches[1]
            }
        }

        # Skip loop container nodes such as GROUP_1, GROUP_2, etc.
        while ($segmentIndex -lt ($parsed.Count - 1)) {
            $candidateSegmentName = Get-LocalName -name $parsed[$segmentIndex].Name
            if ($candidateSegmentName -match '^GROUP_\d+$') {
                $segmentIndex++
                continue
            }

            break
        }

        $elementIndex = $segmentIndex + 1
        if ($parsed.Count -le $elementIndex) {
            continue
        }

        $segmentId = (Get-LocalName -name $parsed[$segmentIndex].Name).ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($segmentId) -or $segmentId -notmatch '^[A-Z0-9]{2,3}$') {
            continue
        }

        $elementName = Get-LocalName -name $parsed[$elementIndex].Name
        $position = $null
        if ($elementName -match ([regex]::Escape($segmentId) + '(\d{2})$')) {
            $position = [int]$Matches[1]
        } elseif ($elementName -match '(\d{2})$') {
            $position = [int]$Matches[1]
        }

        if (-not $position -or $position -lt 1) {
            continue
        }

        $positionKey = [string]$position
        if ($controlSegments -contains $segmentId) {
            if (-not $controlElementMaps.ContainsKey($segmentId)) {
                $controlElementMaps[$segmentId] = [ordered]@{}
            }

            if (-not $controlElementMaps[$segmentId].Contains($positionKey)) {
                $controlElementMaps[$segmentId][$positionKey] = $lineValue
            } elseif ([string]::IsNullOrWhiteSpace([string]$controlElementMaps[$segmentId][$positionKey]) -and -not [string]::IsNullOrWhiteSpace($lineValue)) {
                $controlElementMaps[$segmentId][$positionKey] = $lineValue
            }
            continue
        }

        $occurrenceParts = New-Object System.Collections.Generic.List[string]
        for ($idx = 0; $idx -le $segmentIndex; $idx++) {
            $occurrenceParts.Add((Get-LocalName -name $parsed[$idx].Name))
        }
        $baseOccurrenceKey = ($occurrenceParts -join '/')
        $occurrenceKey = if ($latestOccurrenceByBase.ContainsKey($baseOccurrenceKey)) { $latestOccurrenceByBase[$baseOccurrenceKey] } else { $baseOccurrenceKey }

        if (($segmentId -eq 'REF' -or $segmentId -eq 'N1') -and $position -eq 1) {
            $existingOccurrence = if ($segmentOccurrences.ContainsKey($occurrenceKey)) { $segmentOccurrences[$occurrenceKey] } else { $null }
            if ($existingOccurrence -and $existingOccurrence.Elements.Count -gt 0) {
                $nextCount = if ($occurrenceCounters.ContainsKey($baseOccurrenceKey)) { [int]$occurrenceCounters[$baseOccurrenceKey] + 1 } else { 1 }
                $occurrenceCounters[$baseOccurrenceKey] = $nextCount
                $occurrenceKey = "$baseOccurrenceKey#$nextCount"
            }
        }

        if (-not $segmentOccurrences.ContainsKey($occurrenceKey)) {
            $segmentOccurrences[$occurrenceKey] = [ordered]@{
                SegmentId = $segmentId
                Elements = [ordered]@{}
            }
            $occurrenceOrder.Add($occurrenceKey)
        }

        if ($segmentId -eq 'REF' -or $segmentId -eq 'N1') {
            $latestOccurrenceByBase[$baseOccurrenceKey] = $occurrenceKey
        }

        if (-not $segmentOccurrences[$occurrenceKey].Elements.Contains($positionKey)) {
            $segmentOccurrences[$occurrenceKey].Elements[$positionKey] = $lineValue
        } elseif ([string]::IsNullOrWhiteSpace([string]$segmentOccurrences[$occurrenceKey].Elements[$positionKey]) -and -not [string]::IsNullOrWhiteSpace($lineValue)) {
            $segmentOccurrences[$occurrenceKey].Elements[$positionKey] = $lineValue
        }
    }

    if ($segmentOccurrences.Count -eq 0 -and -not $controlElementMaps.ContainsKey('ST')) {
        return $null
    }

    $transactionSet = Resolve-TransactionSetId -transactionSetHint $effectiveTransactionSetHint

    if ($occurrenceOrder.Count -eq 0) {
        $fallbackSegmentId = if ($transactionSet -eq '214') { 'B10' } else { 'BRA' }
        $fallbackKey = "X12/$fallbackSegmentId"
        $segmentOccurrences[$fallbackKey] = [ordered]@{
            SegmentId = $fallbackSegmentId
            Elements = [ordered]@{ '1' = $null; '2' = $null; '3' = $null }
        }
        $occurrenceOrder.Add($fallbackKey)
    }

    $segmentLines = New-Object System.Collections.Generic.List[string]
    $isaControlNumber = $null
    $stControlNumber = $null
    $generationWarnings = New-Object System.Collections.Generic.List[string]

    $addGenerationWarning = {
        param([string]$message)
        if ([string]::IsNullOrWhiteSpace($message)) {
            return
        }

        if (-not $generationWarnings.Contains($message)) {
            $generationWarnings.Add($message)
        }
    }

    $emitSegment = {
        param([string]$segmentId, [hashtable]$sourceMap, [int]$forceElementCount)

        $elementMap = [ordered]@{}
        if ($sourceMap) {
            foreach ($k in $sourceMap.Keys) {
                $elementMap[[string]$k] = $sourceMap[$k]
            }
        }

        if ($forceElementCount -gt 0) {
            for ($i = 1; $i -le $forceElementCount; $i++) {
                $key = [string]$i
                if (-not $elementMap.Contains($key)) {
                    $elementMap[$key] = $null
                }
            }
        }

        if ($elementMap.Count -eq 0) {
            $elementMap['1'] = $null
        }

        $presentPositions = New-Object System.Collections.Generic.List[int]
        foreach ($k in $elementMap.Keys) {
            $presentPositions.Add([int]$k)
        }

        $maxPos = (($elementMap.Keys | ForEach-Object { [int]$_ }) | Measure-Object -Maximum).Maximum
        if (-not $maxPos) { $maxPos = 1 }

        $values = New-Object System.Collections.Generic.List[string]
        for ($pos = 1; $pos -le [int]$maxPos; $pos++) {
            if ($segmentId -eq 'ST' -and $pos -eq 1) {
                $values.Add($transactionSet)
                continue
            }

            if ($segmentId -eq 'IEA' -and $pos -eq 2 -and -not [string]::IsNullOrWhiteSpace($isaControlNumber)) {
                $values.Add($isaControlNumber)
                continue
            }

            if ($segmentId -eq 'GE' -and $pos -eq 2 -and -not [string]::IsNullOrWhiteSpace($stControlNumber)) {
                $values.Add($stControlNumber)
                continue
            }

            $explicitValue = $null
            $posKey = [string]$pos
            $hasExplicitPath = $elementMap.Contains($posKey)
            if ($hasExplicitPath) {
                $explicitValue = [string]$elementMap[$posKey]
            }

            if (-not [string]::IsNullOrWhiteSpace($explicitValue)) {
                if ($segmentId -eq 'GS' -and $pos -eq 2 -and $explicitValue -eq 'SAMPLE_VALUE' -and -not [string]::IsNullOrWhiteSpace($gsSenderFallback)) {
                    $values.Add($gsSenderFallback)
                    & $addGenerationWarning "GS02 used ISA06 fallback because input mapped to placeholder text."
                    continue
                }

                if ($segmentId -eq 'GS' -and $pos -eq 3 -and $explicitValue -eq 'SAMPLE_VALUE' -and -not [string]::IsNullOrWhiteSpace($gsReceiverFallback)) {
                    $values.Add($gsReceiverFallback)
                    & $addGenerationWarning "GS03 used ISA08 fallback because input mapped to placeholder text."
                    continue
                }

                if ($explicitValue -eq $NoMappingToken) {
                    $values.Add('')
                    & $addGenerationWarning "$segmentId$('{0:D2}' -f $pos) was explicitly marked as no mapping and left blank."
                    continue
                }

                $formattedValue = Format-EdiElementValue -segmentId $segmentId -position $pos -value $explicitValue
                if ($segmentId -eq 'ISA' -and ($pos -eq 6 -or $pos -eq 8 -or $pos -eq 9 -or $pos -eq 10 -or $pos -eq 13) -and $formattedValue -ne $explicitValue) {
                    & $addGenerationWarning "$segmentId$('{0:D2}' -f $pos) value was auto-adjusted to required X12 length/format."
                }
                if ($segmentId -eq 'ISA' -and $pos -eq 13) {
                    $isaControlNumber = $formattedValue
                }
                if ($segmentId -eq 'ST' -and $pos -eq 2) {
                    $stControlNumber = $formattedValue
                }
                $values.Add($formattedValue)
            } else {
                if ($hasExplicitPath) {
                    if ($segmentId -eq 'PID') {
                        # For explicit PID no-mapping rows, keep element blank instead of SAMPLE_VALUE.
                        $values.Add('')
                        continue
                    }

                    $sampleValue = Get-SampleEdiElementValue -segmentId $segmentId -position $pos -transactionSet $transactionSet
                    $formattedValue = Format-EdiElementValue -segmentId $segmentId -position $pos -value $sampleValue
                    if ($segmentId -eq 'ISA' -and $pos -eq 13) {
                        $isaControlNumber = $formattedValue
                    }
                    if ($segmentId -eq 'ST' -and $pos -eq 2) {
                        $stControlNumber = $formattedValue
                    }
                    $values.Add($formattedValue)
                    continue
                }

                $hasLowerPresent = $false
                $hasHigherPresent = $false
                foreach ($presentPos in $presentPositions) {
                    if ($presentPos -lt $pos) { $hasLowerPresent = $true }
                    if ($presentPos -gt $pos) { $hasHigherPresent = $true }
                }

                if (($controlSegments -notcontains $segmentId) -and $hasLowerPresent -and $hasHigherPresent) {
                    # Keep sparse interior element gaps blank when this position is not mapped.
                    $values.Add('')
                    continue
                }

                if ($segmentId -eq 'RCD' -and $pos -ge 2) {
                    # Keep optional trailing RCD elements blank when no mapped value is provided.
                    $values.Add('')
                    continue
                }

                if ($segmentId -eq 'PRF' -and ($pos -ge 2 -and $pos -le 4)) {
                    # Keep PRF02-PRF04 intentionally blank when only PRF01/PRF05 are mapped.
                    $values.Add('')
                    continue
                }

                if ($segmentId -eq 'PID' -and ($pos -ge 2 -and $pos -le 4)) {
                    # Keep PID02-PID04 intentionally blank when no mapped value is provided.
                    $values.Add('')
                    continue
                }

                $sampleValue = Get-SampleEdiElementValue -segmentId $segmentId -position $pos -transactionSet $transactionSet
                $formattedValue = Format-EdiElementValue -segmentId $segmentId -position $pos -value $sampleValue
                if ($segmentId -eq 'ISA' -and $pos -eq 13) {
                    $isaControlNumber = $formattedValue
                }
                if ($segmentId -eq 'ST' -and $pos -eq 2) {
                    $stControlNumber = $formattedValue
                }
                $values.Add($formattedValue)
            }
        }

        $segmentText = $segmentId
        if ($segmentId -eq 'RCD') {
            while ($values.Count -gt 0 -and [string]::IsNullOrEmpty($values[$values.Count - 1])) {
                $values.RemoveAt($values.Count - 1)
            }
        }

        if ($values.Count -gt 0) {
            $segmentText += '*' + ($values -join '*')
        }
        $segmentText += '~'
        $segmentLines.Add($segmentText)
    }

    $isaMap = if ($controlElementMaps.ContainsKey('ISA')) { $controlElementMaps['ISA'] } else { $null }
    $gsMap = if ($controlElementMaps.ContainsKey('GS')) { $controlElementMaps['GS'] } else { $null }
    $stMap = if ($controlElementMaps.ContainsKey('ST')) { $controlElementMaps['ST'] } else { $null }

    $gsSenderSource = if ($isaMap -and $isaMap.Contains('6')) { [string]$isaMap['6'] } else { [string](Get-SampleEdiElementValue -segmentId 'ISA' -position 6 -transactionSet $transactionSet) }
    $gsSenderFallback = [string](Format-EdiElementValue -segmentId 'ISA' -position 6 -value $gsSenderSource)
    $gsSenderFallback = $gsSenderFallback.Trim()
    $gsReceiverSource = if ($isaMap -and $isaMap.Contains('8')) { [string]$isaMap['8'] } else { [string](Get-SampleEdiElementValue -segmentId 'ISA' -position 8 -transactionSet $transactionSet) }
    $gsReceiverFallback = [string](Format-EdiElementValue -segmentId 'ISA' -position 8 -value $gsReceiverSource)
    $gsReceiverFallback = $gsReceiverFallback.Trim()

    & $emitSegment 'ISA' $isaMap 16
    & $emitSegment 'GS' $gsMap 8
    & $emitSegment 'ST' $stMap 2

    $txnDataSegmentCount = 0
    foreach ($occurrenceKey in $occurrenceOrder) {
        $occurrence = $segmentOccurrences[$occurrenceKey]
        if (-not $occurrence) { continue }

        $occurrenceSegmentId = [string]$occurrence.SegmentId
        if ([string]::IsNullOrWhiteSpace($occurrenceSegmentId)) { continue }
        if ($controlSegments -contains $occurrenceSegmentId) { continue }

        & $emitSegment $occurrenceSegmentId $occurrence.Elements 0
        $txnDataSegmentCount++
    }

    $txnSegmentCount = 2 + $txnDataSegmentCount
    $interchangeControlNumber = if ([string]::IsNullOrWhiteSpace($isaControlNumber)) { '000000001' } else { $isaControlNumber }
    $functionalGroupControlNumber = if ([string]::IsNullOrWhiteSpace($stControlNumber)) { '0001' } else { $stControlNumber }
    $seLine = "SE*$txnSegmentCount*0001~"
    $segmentLines.Add($seLine)
    $segmentLines.Add("GE*1*$functionalGroupControlNumber~")
    $segmentLines.Add("IEA*1*$interchangeControlNumber~")

    return [pscustomobject]@{
        EdiText = ($segmentLines -join [Environment]::NewLine)
        Warnings = @($generationWarnings)
    }
}

function New-ElementFromSpec {
    param(
        [System.Xml.XmlDocument]$doc,
        [string]$fullName,
        [hashtable]$attributes,
        [string]$nsValue
    )

    $parts = $fullName.Split(':')
    if ($parts.Count -eq 2) {
        if ([string]::IsNullOrWhiteSpace($parts[0]) -or [string]::IsNullOrWhiteSpace($parts[1]) -or $parts[1] -match '[@\s]') {
            throw "Invalid element name '$fullName'"
        }
    } elseif ($fullName -match ':' -or [string]::IsNullOrWhiteSpace($fullName) -or $fullName -match '[@\s]') {
        throw "Invalid element name '$fullName'"
    }

    try {
        if ($parts.Count -eq 2) {
            $node = $doc.CreateElement($parts[0], $parts[1], $nsValue)
        } else {
            if ([string]::IsNullOrWhiteSpace($nsValue)) {
                $node = $doc.CreateElement($fullName)
            } else {
                $node = $doc.CreateElement($fullName, $nsValue)
            }
        }
    } catch {
        throw "Invalid element name '$fullName'"
    }

    foreach ($k in $attributes.Keys) {
        $null = $node.SetAttribute($k, $attributes[$k])
    }

    return $node
}

function Is-SameSpec {
    param(
        [string]$nameA,
        [hashtable]$attrsA,
        [string]$nameB,
        [hashtable]$attrsB
    )

    if ($nameA -ne $nameB) { return $false }
    if ($attrsA.Count -ne $attrsB.Count) { return $false }

    foreach ($k in $attrsA.Keys) {
        if (-not $attrsB.ContainsKey($k)) { return $false }
        if ($attrsA[$k] -ne $attrsB[$k]) { return $false }
    }

    return $true
}

function Get-OrCreate-Child {
    param(
        [System.Xml.XmlDocument]$doc,
        [System.Xml.XmlElement]$parent,
        [string]$fullName,
        [hashtable]$attributes,
        [string]$nsValue
    )

    $effectiveAttributes = @{}
    foreach ($k in $attributes.Keys) {
        if ($k -eq 'xmlns' -or $k.StartsWith('xmlns:')) { continue }
        $effectiveAttributes[$k] = $attributes[$k]
    }

    # Domain-specific merge: Party entries with same typeCode under the same parent
    # should be represented by one node that aggregates children.
    if (($fullName -eq 'px:Party' -or $fullName -eq 'Party') -and $effectiveAttributes.ContainsKey('typeCode')) {
        foreach ($child in $parent.ChildNodes) {
            if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
            if ($child.LocalName -ne 'Party') { continue }
            if ($child.GetAttribute('typeCode') -eq $effectiveAttributes['typeCode']) {
                return $child
            }
        }
    }

    foreach ($child in $parent.ChildNodes) {
        if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
        if ($child.Name -ne $fullName) { continue }

        $same = $true
        $nonNsAttrCount = 0
        foreach ($attr in $child.Attributes) {
            if ($attr.Prefix -eq 'xmlns' -or $attr.Name -eq 'xmlns') { continue }
            $nonNsAttrCount++
        }

        if ($nonNsAttrCount -ne $effectiveAttributes.Count) {
            $same = $false
        }

        if ($same) {
            foreach ($k in $effectiveAttributes.Keys) {
                if ($child.GetAttribute($k) -ne $effectiveAttributes[$k]) {
                    $same = $false
                    break
                }
            }
        }

        if ($same) { return $child }
    }

    $node = New-ElementFromSpec -doc $doc -fullName $fullName -attributes $effectiveAttributes -nsValue $nsValue

    $null = $parent.AppendChild($node)
    return $node
}

function Get-ElementKey {
    param([System.Xml.XmlElement]$element)

    $pairs = New-Object System.Collections.Generic.List[string]
    foreach ($attr in $element.Attributes) {
        if ($attr.Prefix -eq 'xmlns' -or $attr.Name -eq 'xmlns') { continue }
        $pairs.Add("$($attr.Name)=$($attr.Value)")
    }

    $sorted = $pairs | Sort-Object
    return "$($element.Name)|$($sorted -join ';')"
}

function Merge-DuplicateSiblings {
    param([System.Xml.XmlElement]$parent)

    # First recurse into descendants.
    $childElements = @($parent.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element })
    foreach ($child in $childElements) {
        Merge-DuplicateSiblings -parent $child
    }

    $seen = @{}
    $currentChildren = @($parent.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element })
    foreach ($child in $currentChildren) {
        $key = Get-ElementKey -element $child
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $child
            continue
        }

        $target = [System.Xml.XmlElement]$seen[$key]

        if ([string]::IsNullOrWhiteSpace($target.InnerText) -and -not [string]::IsNullOrWhiteSpace($child.InnerText) -and $child.ChildNodes.Count -eq 1) {
            $target.InnerText = $child.InnerText
        }

        $toMove = @($child.ChildNodes)
        foreach ($n in $toMove) {
            $imported = $parent.OwnerDocument.ImportNode($n, $true)
            $null = $target.AppendChild($imported)
        }

        $null = $parent.RemoveChild($child)
    }
}

function Merge-PartySiblingsByTypeCode {
    param([System.Xml.XmlElement]$parent)

    for ($idx = 0; $idx -lt $parent.ChildNodes.Count; $idx++) {
        $child = $parent.ChildNodes.Item($idx)
        if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            Merge-PartySiblingsByTypeCode -parent $child
        }
    }

    $i = 0
    while ($i -lt $parent.ChildNodes.Count) {
        $baseNode = $parent.ChildNodes.Item($i)
        if ($baseNode.NodeType -ne [System.Xml.XmlNodeType]::Element -or $baseNode.LocalName -ne 'Party') {
            $i++
            continue
        }

        $baseTypeCode = $baseNode.GetAttribute('typeCode').Trim()
        if ([string]::IsNullOrWhiteSpace($baseTypeCode)) {
            $i++
            continue
        }

        $j = $i + 1
        while ($j -lt $parent.ChildNodes.Count) {
            $candidate = $parent.ChildNodes.Item($j)
            if ($candidate.NodeType -eq [System.Xml.XmlNodeType]::Element -and $candidate.LocalName -eq 'Party') {
                $candidateTypeCode = $candidate.GetAttribute('typeCode').Trim()
                if ($candidateTypeCode -eq $baseTypeCode) {
                    $toMove = @($candidate.ChildNodes)
                    foreach ($n in $toMove) {
                        $imported = $parent.OwnerDocument.ImportNode($n, $true)
                        $null = $baseNode.AppendChild($imported)
                    }
                    $null = $parent.RemoveChild($candidate)
                    continue
                }
            }
            $j++
        }

        $i++
    }
}

function Wrap-ShipmentUnitDocumentReferences {
    param([System.Xml.XmlElement]$parent)

    for ($idx = 0; $idx -lt $parent.ChildNodes.Count; $idx++) {
        $child = $parent.ChildNodes.Item($idx)
        if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            Wrap-ShipmentUnitDocumentReferences -parent $child
        }
    }

    if ($parent.LocalName -ne 'ShipmentUnit') {
        return
    }

    $docRefs = @($parent.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element -and $_.LocalName -eq 'DocumentReference' })
    if ($docRefs.Count -eq 0) {
        return
    }

    $firstDocRef = [System.Xml.XmlElement]$docRefs[0]
    $wrapper = New-ElementFromSpec -doc $parent.OwnerDocument -fullName 'px:DocumentReferences' -attributes @{} -nsValue $parent.NamespaceURI

    $null = $parent.InsertBefore($wrapper, $firstDocRef)

    foreach ($sourceNode in $docRefs) {
        $imported = $parent.OwnerDocument.ImportNode($sourceNode, $true)
        $null = $wrapper.AppendChild($imported)
    }

    foreach ($sourceNode in $docRefs) {
        $null = $parent.RemoveChild($sourceNode)
    }
}

for ($lineIndex = 0; $lineIndex -lt $lines.Count; $lineIndex++) {
    $lineNo = $lineIndex + 1
    $line = $lines[$lineIndex]

    if ([string]::IsNullOrWhiteSpace($line)) {
        $skipped++
        $skippedDetails.Add("Line ${lineNo}: blank or whitespace-only input")
        continue
    }

    $clean = $line.Trim()
    $clean = Normalize-PathLine -line $line

    $split = Split-XmlPathSegments -path $clean
    if (-not $split.Valid) {
        $skipped++
        $skippedDetails.Add("Line ${lineNo}: $($split.Error) in '$clean'")
        $malformed.Add("Line ${lineNo}: $($split.Error) in '$clean'")
        continue
    }

    $segments = $split.Segments
    if ($segments.Count -eq 0) {
        $skipped++
        $skippedDetails.Add("Line ${lineNo}: path produced no readable segments")
        continue
    }

    $parsed = New-Object System.Collections.Generic.List[hashtable]
    $lineMalformed = $false

    foreach ($segment in $segments) {
        $p = Parse-Segment -segment $segment -lineNo $lineNo
        if (-not $p.Valid) {
            $lineMalformed = $true
            $malformed.Add($p.Error)
            break
        }
        $parsed.Add($p)
    }

    if ($lineMalformed -or $parsed.Count -eq 0) {
        $skipped++
        if ($parsed.Count -eq 0) {
            $skippedDetails.Add("Line ${lineNo}: no valid parsed segments")
        } else {
            $skippedDetails.Add("Line ${lineNo}: parsing failed for one or more segments")
        }
        continue
    }

    if (-not $root) {
        $rootSpec = $parsed[0]
        $rootAttrs = @{}
        $rootXmlnsAttrs = @{}
        foreach ($k in $rootSpec.Attributes.Keys) {
            if ($k -eq 'xmlns' -or $k.StartsWith('xmlns:')) {
                $rootXmlnsAttrs[$k] = $rootSpec.Attributes[$k]
            } else {
                $rootAttrs[$k] = $rootSpec.Attributes[$k]
            }
        }

        $defaultRootNs = $null
        if ($rootXmlnsAttrs.ContainsKey('xmlns')) {
            $defaultRootNs = $rootXmlnsAttrs['xmlns']
        }

        $rootParts = $rootSpec.Name.Split(':')
        if ($rootParts.Count -eq 2) {
            $rootNsToUse = if ($defaultRootNs) { $defaultRootNs } else { $ns }
            $root = $xmlDoc.CreateElement($rootParts[0], $rootParts[1], $rootNsToUse)
        } else {
            if ($defaultRootNs) {
                $root = $xmlDoc.CreateElement($rootSpec.Name, $defaultRootNs)
            } else {
                $root = $xmlDoc.CreateElement($rootSpec.Name)
            }
        }

        foreach ($k in $rootAttrs.Keys) {
            $null = $root.SetAttribute($k, $rootAttrs[$k])
        }

        foreach ($k in $rootXmlnsAttrs.Keys) {
            if ($k.StartsWith('xmlns:')) {
                $nsPrefix = $k.Substring(6)
                $null = $root.SetAttribute('xmlns', $nsPrefix, 'http://www.w3.org/2000/xmlns/', $rootXmlnsAttrs[$k])
            }
        }

        $null = $xmlDoc.AppendChild($root)
    }

    $childNs = $root.NamespaceURI

    $start = 0
    if ($parsed[0].Name -eq $root.Name) {
        $start = 1
    } else {
        $skipped++
        $skippedDetails.Add("Line ${lineNo}: root '$($parsed[0].Name)' does not match detected root '$($root.Name)'")
        $malformed.Add("Line ${lineNo}: root '$($parsed[0].Name)' does not match detected root '$($root.Name)'")
        continue
    }

    $cursor = $root
    for ($i = $start; $i -lt $parsed.Count; $i++) {
        if ($parsed[$i].IsAttributeOnly) {
            if ($i -ne ($parsed.Count - 1)) {
                $skipped++
                $skippedDetails.Add("Line ${lineNo}: attribute segment '$($parsed[$i].Name)' must be the last segment in the path")
                $malformed.Add("Line ${lineNo}: attribute segment '$($parsed[$i].Name)' must be the last segment in the path")
                continue
            }

            $attrName = $parsed[$i].Name.Substring(1)
            $attrValue = 'SAMPLE_VALUE'
            if ($attrName -eq 'xmlns' -or $attrName.StartsWith('xmlns:')) {
                $attrValue = if ($childNs) { $childNs } else { $ns }
            }

            $null = $cursor.SetAttribute($attrName, $attrValue)
            continue
        }

        $cursor = Get-OrCreate-Child -doc $xmlDoc -parent $cursor -fullName $parsed[$i].Name -attributes $parsed[$i].Attributes -nsValue $childNs
        if ($i -eq ($parsed.Count - 1) -and [string]::IsNullOrWhiteSpace($cursor.InnerText)) {
            if ($parsed[$i].Name -match 'Date|Time') {
                $cursor.InnerText = '2026-05-22T10:30:00Z'
            } else {
                $cursor.InnerText = 'SAMPLE_VALUE'
            }
        }
    }

    $used++
}

if (-not $root) {
    Write-Error "No valid paths found. XML was not generated."
    exit 1
}

Merge-PartySiblingsByTypeCode -parent $root
Wrap-ShipmentUnitDocumentReferences -parent $root
Merge-DuplicateSiblings -parent $root

$settings = New-Object System.Xml.XmlWriterSettings
$settings.Indent = $true
$settings.IndentChars = "  "
$settings.Encoding = [System.Text.Encoding]::UTF8
$settings.OmitXmlDeclaration = $false

$ms = New-Object System.IO.MemoryStream
$writer = [System.Xml.XmlWriter]::Create($ms, $settings)
$xmlDoc.Save($writer)
$writer.Close()

$xmlString = [System.Text.Encoding]::UTF8.GetString($ms.ToArray())
$xmlString = $xmlString -replace ' xmlns:px="urn:px"', ''

$outputFormat = 'XML'
$finalOutputPath = $outputPath
$isX12Root = ($root.Name -eq 'X12' -or $root.LocalName -eq 'X12')
$ediValidation = $null
$ediGenerationWarnings = @()

if ($isX12Root) {
    $transactionSetHint = Get-TransactionSetHintFromInputPath -path $inputPath
    $transactionSetId = Resolve-TransactionSetId -transactionSetHint $transactionSetHint
    $ediBuildResult = Build-SampleEdiFromPathLines -pathLines $lines -transactionSetHint $transactionSetHint -pathValueMap $pathValueMap
    $ediText = if ($ediBuildResult -and $ediBuildResult.PSObject.Properties.Name -contains 'EdiText') { [string]$ediBuildResult.EdiText } else { [string]$ediBuildResult }
    if ($ediBuildResult -and $ediBuildResult.PSObject.Properties.Name -contains 'Warnings') {
        $ediGenerationWarnings = @($ediBuildResult.Warnings)
    }
    if (-not [string]::IsNullOrWhiteSpace($ediText)) {
        $ediText | Out-File -FilePath $ediOutputPath -Encoding ascii
        $resolvedSefSchemaPath = Resolve-SefSchemaPath -transactionSetId $transactionSetId -preferredPath $sefSchemaPath
        $sefSchema = Get-SefSchemaModel -schemaPath $resolvedSefSchemaPath -transactionSetId $transactionSetId
        $ediValidation = Validate-EdiAgainstSef -ediText $ediText -schemaModel $sefSchema -transactionSetId $transactionSetId
        $outputFormat = 'EDI'
        $finalOutputPath = $ediOutputPath
    } else {
        Write-Warning "X12 root detected but EDI output could not be constructed. Falling back to XML output."
        $xmlString | Out-File -FilePath $outputPath -Encoding UTF8
    }
} else {
    $xmlString | Out-File -FilePath $outputPath -Encoding UTF8
}

Write-Output "Summary:"
Write-Output "Output Format: $outputFormat"
Write-Output "Output Path: $finalOutputPath"
Write-Output "Detected Root: $($root.Name)"
Write-Output "Lines Read: $($lines.Count)"
Write-Output "Lines Used: $used"
Write-Output "Lines Skipped: $skipped"
Write-Output "Malformed Lines: $($malformed.Count)"
if ($outputFormat -eq 'EDI') {
    Write-Output "Generation Warnings: $($ediGenerationWarnings.Count)"
    foreach ($message in $ediGenerationWarnings) {
        Write-Output "GEN-WARNING: $message"
    }
}
if ($ediValidation) {
    Write-Output "Validation Errors: $($ediValidation.Errors.Count)"
    Write-Output "Validation Warnings: $($ediValidation.Warnings.Count)"
    foreach ($message in $ediValidation.Errors) {
        Write-Output "ERROR: $message"
    }
    foreach ($message in $ediValidation.Warnings) {
        Write-Output "WARNING: $message"
    }
}
if ($malformed.Count -gt 0) {
    Write-Output "Malformed Details:"
    foreach ($entry in $malformed) {
        Write-Output "- $entry"
    }
}
if ($skippedDetails.Count -gt 0) {
    Write-Output "Skipped Details:"
    foreach ($entry in $skippedDetails) {
        Write-Output "- $entry"
    }
}
