$inputPath = "C:\Users\syuhada.yazid\OneDrive - WiseTech Global\Desktop\sample files\xmlpath.txt"
$outputPath = "C:\Users\syuhada.yazid\OneDrive - WiseTech Global\Desktop\XML Generator\sample.xml"
$xlsxPathColumn = "Element Xpath or Segment, Loop, Element Identifier"
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

function Get-InputLines {
    param(
        [string]$sourcePath,
        [string]$pathColumn,
        [string]$worksheetName
    )

    $ext = [System.IO.Path]::GetExtension($sourcePath).ToLowerInvariant()

    if ($ext -eq ".txt") {
        return @(Get-Content $sourcePath)
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
        $resolvedColumn = $null
        foreach ($prop in $firstRowProps) {
            if ((& $normalize $prop) -eq $targetNormalized) {
                $resolvedColumn = $prop
                break
            }
        }

        if (-not $resolvedColumn) {
            $embeddedHeaderMatch = Find-EmbeddedHeaderValues -xlsxPath $sourcePath -sheetName $worksheetName -normalizeFn $normalize -normalizedHeader $targetNormalized
            if ($embeddedHeaderMatch) {
                Write-Warning "Detected embedded header '$pathColumn' on worksheet '$($embeddedHeaderMatch.Sheet)' row $($embeddedHeaderMatch.RowNumber), column '$($embeddedHeaderMatch.Column)'."
                return @($embeddedHeaderMatch.Values | Where-Object { $_ -match '^\s*/?[A-Za-z_][\w:\-]*(/|\[)' })
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

        $result = New-Object System.Collections.Generic.List[string]
        foreach ($row in $rows) {
            $value = [string]$row.$resolvedColumn
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                $result.Add($value)
            }
        }
        return @($result)
    }

    Write-Error "Unsupported input file type '$ext'. Supported: .txt, .xlsx"
    exit 1
}

$lines = Get-InputLines -sourcePath $inputPath -pathColumn $xlsxPathColumn -worksheetName $xlsxWorksheetName
$used = 0
$skipped = 0
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

function Parse-Segment {
    param(
        [string]$segment,
        [int]$lineNo
    )

    $result = @{
        Valid = $false
        Name = ""
        Attributes = @{}
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
            if (($predicate -match '@') -and $attrMatches.Count -eq 0) {
                $result.Error = "Line ${lineNo}: unsupported/malformed predicate '$predicate'"
                return $result
            }
            foreach ($m in $attrMatches) {
                $attrName = $m.Groups[1].Value
                $attrs[$attrName] = $m.Groups[3].Value
            }
        }
    }

    $result.Valid = $true
    $result.Name = $name
    $result.Attributes = $attrs
    return $result
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
        continue
    }

    $clean = $line.Trim()
    while ($clean -match '(^|/)px:px:') {
        $clean = $clean -replace '(^|/)px:px:', '$1px:'
    }
    $clean = $clean -replace 'px:/Shipment', 'px:Shipment'
    $clean = $clean -replace 'px:/', 'px:'
    $clean = [regex]::Replace($clean, '(?<=[A-Za-z0-9_\]])(px:)(?=[A-Za-z_])', '/$1')
    # Fix common malformed pattern: "]TagName" should be "]/TagName".
    $clean = [regex]::Replace($clean, '(\[[^\]]+\])(?=[A-Za-z_])', '$1/')

    $split = Split-XmlPathSegments -path $clean
    if (-not $split.Valid) {
        $skipped++
        $malformed.Add("Line ${lineNo}: $($split.Error) in '$clean'")
        continue
    }

    $segments = $split.Segments
    if ($segments.Count -eq 0) {
        $skipped++
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
        $malformed.Add("Line ${lineNo}: root '$($parsed[0].Name)' does not match detected root '$($root.Name)'")
        continue
    }

    $cursor = $root
    for ($i = $start; $i -lt $parsed.Count; $i++) {
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
$xmlString | Out-File -FilePath $outputPath -Encoding UTF8

Write-Output "Summary:"
Write-Output "Output Path: $outputPath"
Write-Output "Detected Root: $($root.Name)"
Write-Output "Lines Read: $($lines.Count)"
Write-Output "Lines Used: $used"
Write-Output "Lines Skipped: $skipped"
Write-Output "Malformed Lines: $($malformed.Count)"
if ($malformed.Count -gt 0) {
    Write-Output "Malformed Details:"
    foreach ($entry in $malformed) {
        Write-Output "- $entry"
    }
}
