$ErrorActionPreference = "Stop"

function Get-XmlAttribute {
    param(
        [System.Xml.XmlNode]$Node,
        [string[]]$Names
    )

    foreach ($name in $Names) {
        $attribute = $Node.Attributes[$name]
        if ($null -ne $attribute) {
            return $attribute.Value
        }
    }

    return $null
}

function ConvertTo-SafeFileName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        $Name = "Ohne Titel"
    }

    $safeName = ($Name -replace '[\\/:*?"<>|]', '_').Trim().TrimEnd('.', ' ')

    if ($safeName -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$') {
        $safeName = "_" + $safeName
    }

    if ($safeName.Length -gt 100) {
        $safeName = $safeName.Substring(0, 100).TrimEnd('.', ' ')
    }

    if ([string]::IsNullOrWhiteSpace($safeName)) {
        return "Ohne Titel"
    }

    return $safeName
}

function ConvertTo-MarkdownText {
    param([string]$Text)

    if ($null -eq $Text) {
        return ""
    }

    return ($Text -replace '\|', '\|' -replace "`r?`n", ' ' -replace '\s{2,}', ' ').Trim()
}

function Get-MarkdownPreview {
    param(
        [string]$Path,
        [int]$MaximumLength = 280
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return ""
    }

    $text = Get-Content -LiteralPath $Path -Raw -ErrorAction SilentlyContinue
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ""
    }

    # Metadatenkopf der Quelldatei aus dem Vorschautext entfernen
    $text = $text -replace '(?s)\A---\s*.*?---\s*', ''

    # Markdown-Formatierung für den knappen Indexauszug vereinfachen
    $text = $text -replace '!\[[^\]]*\]\([^)]+\)', ''
    $text = $text -replace '\[([^\]]+)\]\([^)]+\)', '$1'
    $text = $text -replace '^\s{0,3}#{1,6}\s*', ''
    $text = $text -replace '^\s*[-*+]\s+', ''
    $text = $text -replace '^\s*\|.*\|\s*$', ' '
    $text = $text -replace '[`*_>#]', ''
    $text = $text -replace '\s+', ' '
    $text = $text.Trim()

    if ($text.Length -gt $MaximumLength) {
        $text = $text.Substring(0, $MaximumLength).TrimEnd() + "…"
    }

    return $text
}

function Get-OneNotePageLink {
    param(
        $OneNoteApplication,
        [string]$PageId
    )

    try {
        $hyperlink = ""
        $OneNoteApplication.GetHyperlinkToObject($PageId, "", [ref]$hyperlink)
        return $hyperlink
    }
    catch {
        return ""
    }
}

function Write-MarkdownMetadata {
    param(
        [string]$MarkdownPath,
        [string]$PageTitle,
        [string]$NotebookName,
        [string]$SectionName,
        [string]$OneNotePageId,
        [string]$OneNoteLink,
        [string]$ExportedAt
    )

    if (-not (Test-Path -LiteralPath $MarkdownPath)) {
        return
    }

    $body = Get-Content -LiteralPath $MarkdownPath -Raw -ErrorAction Stop
    $body = $body -replace '(?s)\A---\s*.*?---\s*', ''
    $body = $body.TrimStart()

    $safeTitle = $PageTitle -replace "'", "''"
    $safeNotebook = $NotebookName -replace "'", "''"
    $safeSection = $SectionName -replace "'", "''"
    $safeLink = $OneNoteLink -replace "'", "''"

    $frontMatter = @"
---
title: '$safeTitle'
notebook: '$safeNotebook'
section: '$safeSection'
onenote_page_id: '$OneNotePageId'
onenote_link: '$safeLink'
exported_at: '$ExportedAt'
---

"@

    Set-Content -LiteralPath $MarkdownPath `
        -Value ($frontMatter + $body + "`r`n") `
        -Encoding UTF8
}

# Pandoc prüfen
$pandocCommand = Get-Command "pandoc.exe" -ErrorAction SilentlyContinue
if ($null -eq $pandocCommand) {
    $pandocCommand = Get-Command "pandoc" -ErrorAction SilentlyContinue
}

$pandocAvailable = ($null -ne $pandocCommand)
$pandocPath = $null

if ($pandocAvailable) {
    $pandocPath = $pandocCommand.Source
    Write-Host "Pandoc gefunden: $pandocPath"
}
else {
    Write-Host "`nPandoc wurde nicht gefunden."
    Write-Host "Windows-Installationshinweise: https://pandoc.org/installing.html"
    Write-Host "Laden Sie dort den Windows-Installer herunter."
    Write-Host "Nach der Installation PowerShell neu starten und das Skript erneut ausführen.`n"

    $answer = Read-Host "Trotzdem nur DOCX-Dateien exportieren? [J]a/[N]ein"
    if ($answer -notmatch '^(j|ja|y|yes)$') {
        Write-Host "Abgebrochen. Es wurden keine Dateien exportiert."
        return
    }

    Write-Host "Es werden nur DOCX-Dateien exportiert; Markdown-Index wird nicht erstellt."
}

# OneNote verbinden
try {
    $oneNote = New-Object -ComObject OneNote.Application
}
catch {
    throw "OneNote konnte nicht angesprochen werden. Öffnen Sie OneNote und versuchen Sie es erneut. Details: $($_.Exception.Message)"
}

# Notizbücher auflisten (2 = Notizbücher)
$notebookXml = ""
$oneNote.GetHierarchy("", 2, [ref]$notebookXml)

[xml]$notebookDocument = $notebookXml
$notebooks = @($notebookDocument.SelectNodes("//*[local-name()='Notebook']"))

if ($notebooks.Count -eq 0) {
    throw "OneNote hat keine verfügbaren Notizbücher zurückgegeben."
}

Write-Host "`nVerfügbare Notizbücher:`n"
for ($i = 0; $i -lt $notebooks.Count; $i++) {
    $name = Get-XmlAttribute $notebooks[$i] @("name", "Name")
    Write-Host ("[{0}] {1}" -f ($i + 1), $name)
}

$selection = Read-Host "`nNummer des zu exportierenden Notizbuchs"
$selectedIndex = 0

if (-not [int]::TryParse($selection, [ref]$selectedIndex) -or
    $selectedIndex -lt 1 -or
    $selectedIndex -gt $notebooks.Count) {
    throw "Ungültige Auswahl. Es wurde nichts exportiert."
}

$selectedNotebook = $notebooks[$selectedIndex - 1]
$notebookName = Get-XmlAttribute $selectedNotebook @("name", "Name")
$notebookId = Get-XmlAttribute $selectedNotebook @("ID", "id")

if ([string]::IsNullOrWhiteSpace($notebookId)) {
    throw "Die ID des ausgewählten Notizbuchs konnte nicht gelesen werden."
}

# Ausgabe-Basisordner
#$defaultOutputRoot = "C:\tmp"
$defaultOutputRoot ="C:\tmp\OneNoteExport4Cosma"
$enteredOutputRoot = Read-Host "Basis-Ausgabeordner eingeben oder Enter für [$defaultOutputRoot]"

if ([string]::IsNullOrWhiteSpace($enteredOutputRoot)) {
    $outputRoot = $defaultOutputRoot
}
else {
    $outputRoot = $enteredOutputRoot.Trim().Trim('"')
}

$outputRoot = [System.IO.Path]::GetFullPath($outputRoot)
New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null

# Gemeinsamer Hauptordner mit DOCX- und Markdown-Struktur
#$notebookFolder = Join-Path $outputRoot (ConvertTo-SafeFileName $notebookName)
#$docxRoot = Join-Path $notebookFolder "docx"
#$markdownRoot = Join-Path $notebookFolder "md"
#$sourceRoot = Join-Path $markdownRoot "quellen"
$notebookFolder = Join-Path $outputRoot (ConvertTo-SafeFileName $notebookName)
$docxRoot = Join-Path $notebookFolder "docx"
$markdownFolderName = "md_{0}" -f (ConvertTo-SafeFileName $notebookName)
$markdownRoot = Join-Path $notebookFolder $markdownFolderName
$sourceRoot = Join-Path $markdownRoot "quellen"

$exportMode = "NewOnly"

if (Test-Path -LiteralPath $notebookFolder) {
    Write-Host "`nDer Notizbuchordner existiert bereits:"
    Write-Host $notebookFolder
    Write-Host "[Ü] Vorhandene DOCX-Dateien überschreiben"
    Write-Host "[N] Nur neue DOCX-Dateien anlegen; vorhandene überspringen"
    Write-Host "[A] Abbrechen"

    $choice = (Read-Host "Auswahl [Ü/N/A]").Trim().ToUpperInvariant()

    if ($choice -in @("Ü", "U")) {
        $exportMode = "Overwrite"
    }
    elseif ($choice -eq "N") {
        $exportMode = "NewOnly"
    }
    else {
        Write-Host "Abgebrochen. Es wurden keine Dateien exportiert."
        return
    }
}
else {
    New-Item -ItemType Directory -Path $notebookFolder -Force | Out-Null
}

New-Item -ItemType Directory -Path $docxRoot -Force | Out-Null
if ($pandocAvailable) {
    New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null
}

Write-Host "`nNotizbuch: $notebookName"
Write-Host "Hauptordner: $notebookFolder"
Write-Host "Exportmodus: $exportMode`n"

# Abschnitte abrufen (3 = Abschnitte)
$sectionXml = ""
$oneNote.GetHierarchy($notebookId, 3, [ref]$sectionXml)

[xml]$sectionDocument = $sectionXml
$sections = @($sectionDocument.SelectNodes("//*[local-name()='Section']"))

if ($sections.Count -eq 0) {
    throw "Im ausgewählten Notizbuch wurden keine Abschnitte gefunden."
}

# Diagnose: Abschnittsname, OneNote-ID und vollständiger XML-Elternpfad protokollieren.
$hierarchyReport = foreach ($section in $sections) {
    $parts = [System.Collections.Generic.List[string]]::new()
    $node = $section

    while ($null -ne $node -and $node.NodeType -ne [System.Xml.XmlNodeType]::Document) {
        if ($node.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            $nodeName = $node.LocalName
            $nameAttribute = $node.Attributes["name"]

            if ($null -ne $nameAttribute -and
                $nodeName -in @("Notebook", "SectionGroup", "Section")) {
                $parts.Insert(0, "$nodeName='$($nameAttribute.Value)'")
            }
        }
        $node = $node.ParentNode
    }

    [pscustomobject]@{
        SectionName = Get-XmlAttribute $section @("name", "Name")
        SectionId   = Get-XmlAttribute $section @("ID", "id")
        XmlPath     = ($parts -join " > ")
    }
}

$hierarchyReportPath = Join-Path $notebookFolder "onenote-abschnittshierarchie.txt"
$hierarchyReport |
    Format-Table -AutoSize |
    Out-String -Width 300 |
    Set-Content -LiteralPath $hierarchyReportPath -Encoding UTF8

Write-Host "Abschnitts-Hierarchie gespeichert: $hierarchyReportPath"

$docxCount = 0
$markdownCount = 0
$skippedCount = 0
$failedCount = 0
$errors = @()
$indexEntries = @()

# get duration
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($section in $sections) {
    $sectionName = Get-XmlAttribute $section @("name", "Name")
    $sectionId = Get-XmlAttribute $section @("ID", "id")

    # Den gesamten OneNote-Papierkorb überspringen, einschließlich aller Unterabschnitte.
    $insideRecycleBin = $false
    $ancestor = $section.ParentNode

    while ($null -ne $ancestor -and $ancestor.NodeType -ne [System.Xml.XmlNodeType]::Document) {
        if ($ancestor.NodeType -eq [System.Xml.XmlNodeType]::Element -and
            $ancestor.LocalName -eq "SectionGroup") {
            $groupNameAttribute = $ancestor.Attributes["name"]
            if ($null -ne $groupNameAttribute -and
                $groupNameAttribute.Value -eq "OneNote_RecycleBin") {
                $insideRecycleBin = $true
                break
            }
        }
        $ancestor = $ancestor.ParentNode
    }

    if ($insideRecycleBin) {
        Write-Host "OneNote-Papierkorb (inkl. Unterabschnitten) übersprungen: $sectionName"
        continue
    }

    if ([string]::IsNullOrWhiteSpace($sectionId)) {
        $failedCount++
        $errors += "Abschnitt '$sectionName': Abschnitts-ID fehlt."
        continue
    }

    # Seiten abrufen (4 = Seiten)
    $pageXml = ""
    try {
        $oneNote.GetHierarchy($sectionId, 4, [ref]$pageXml)
    }
    catch {
        $failedCount++
        $errors += "Abschnitt '$sectionName': Seiten konnten nicht gelesen werden: $($_.Exception.Message)"
        continue
    }

    [xml]$pageDocument = $pageXml
    $pages = @($pageDocument.SelectNodes("//*[local-name()='Page']"))

    if ($pages.Count -eq 0) {
        Write-Host "Abschnitt ohne Seiten übersprungen: $sectionName"
        continue
    }

    # Ausgabeordner erst anlegen, nachdem OneNote mindestens eine Seite geliefert hat.
    $safeSectionName = ConvertTo-SafeFileName $sectionName
    $docxSectionFolder = Join-Path $docxRoot $safeSectionName
    $mdSectionFolder = Join-Path $sourceRoot $safeSectionName

    New-Item -ItemType Directory -Path $docxSectionFolder -Force | Out-Null
    Write-Host "New docx SectionFolder: $docxSectionFolder"

    if ($pandocAvailable) {
        New-Item -ItemType Directory -Path $mdSectionFolder -Force | Out-Null
    }

    $usedPageNames = @{}

    foreach ($page in $pages) {
        $pageName = Get-XmlAttribute $page @("name", "Name")
        $pageId = Get-XmlAttribute $page @("ID", "id")

        $deletedValue = Get-XmlAttribute $page @(
            "isDeleted",
            "IsDeleted",
            "isInRecycleBin",
            "IsInRecycleBin"
        )

        if ($deletedValue -match '^(true|1)$') {
            $skippedCount++
            continue
        }

        if ([string]::IsNullOrWhiteSpace($pageId)) {
            $failedCount++
            $errors += "Seite '$pageName' im Abschnitt '$sectionName': Seiten-ID fehlt."
            continue
        }

        $safePageName = ConvertTo-SafeFileName $pageName
        $pageFileBase = $safePageName
        $duplicateNumber = 2

        while ($usedPageNames.ContainsKey($pageFileBase)) {
            $pageFileBase = "{0} ({1})" -f $safePageName, $duplicateNumber
            $duplicateNumber++
        }
        $usedPageNames[$pageFileBase] = $true

        $docxPath = Join-Path $docxSectionFolder "$pageFileBase.docx"
        $mdPath = Join-Path $mdSectionFolder "$pageFileBase.md"

        $docxExists = Test-Path -LiteralPath $docxPath
        $shouldPublish = $true

        if ($docxExists -and $exportMode -eq "NewOnly") {
            $shouldPublish = $false
            $skippedCount++
        }

        try {
            if ($shouldPublish) {
                # OneNote PublishFormat 5 = Word/DOCX.
                # Publish benötigt einen noch nicht vorhandenen Zieldateinamen.
                if (Test-Path -LiteralPath $docxPath) {
                    Remove-Item -LiteralPath $docxPath -Force
                }

                if ([string]::IsNullOrWhiteSpace($pageId)) {
                    throw "OneNote-Seiten-ID fehlt; Export wird übersprungen."
                }

                if (Test-Path -LiteralPath $docxPath) {
                    Remove-Item -LiteralPath $docxPath -Force -ErrorAction Stop
                }

                if (Test-Path -LiteralPath $docxPath) {
                    throw "Die vorhandene Zieldatei konnte nicht entfernt werden: $docxPath"
                }
                #Write-Host "Publish mit: '$pageId', '$docxPath' "

                $oneNote.Publish($pageId, $docxPath, 5, "")

                #if (-not (Test-Path -LiteralPath $docxPath)) {
                #    throw "OneNote hat keine DOCX-Datei erzeugt."
                #}

                if (-not (Test-Path -LiteralPath $docxPath)) {
                    throw "Publish wurde ausgeführt, aber keine DOCX-Datei gefunden."
                }

                $fileInfo = Get-Item -LiteralPath $docxPath
                if ($fileInfo.Length -eq 0) {
                    throw "Publish hat eine leere DOCX-Datei erzeugt."
                }

                # Erst nach erfolgreichem Export und den Prüfungen zählen
                $docxCount++
            }

            if ($pandocAvailable) {
                # Im Nur-neue-Dateien-Modus vorhandene Markdown-Datei behalten.
                $shouldConvert = $shouldPublish -or
                    (-not (Test-Path -LiteralPath $mdPath))

                if ($shouldConvert) {
                    if (-not (Test-Path -LiteralPath $docxPath)) {
                        throw "Keine DOCX-Quelldatei für die Markdown-Konvertierung gefunden."
                    }

                    $mediaFolder = Join-Path $mdSectionFolder "${pageFileBase}_media"

                    if (Test-Path -LiteralPath $mediaFolder) {
                        Remove-Item -LiteralPath $mediaFolder -Recurse -Force
                    }
                    New-Item -ItemType Directory -Path $mediaFolder -Force | Out-Null

                    $pandocArguments = @(
                        "--from=docx",
                        "--to=gfm",
                        "--extract-media=$mediaFolder",
                        "--output=$mdPath",
                        $docxPath
                    )

                    & $pandocPath @pandocArguments

                    if ($LASTEXITCODE -ne 0) {
                        throw "Pandoc meldete den Exitcode $LASTEXITCODE."
                    }

                    if (-not (Test-Path -LiteralPath $mdPath)) {
                        throw "Pandoc hat keine Markdown-Datei erzeugt."
                    }

                    $oneNoteLink = Get-OneNotePageLink -OneNoteApplication $oneNote -PageId $pageId
                    $exportedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

                    Write-MarkdownMetadata `
                        -MarkdownPath $mdPath `
                        -PageTitle $pageName `
                        -NotebookName $notebookName `
                        -SectionName $sectionName `
                        -OneNotePageId $pageId `
                        -OneNoteLink $oneNoteLink `
                        -ExportedAt $exportedAt

                    $markdownCount++
                }

                if (Test-Path -LiteralPath $mdPath) {
                    $relativePath = "quellen/$safeSectionName/$pageFileBase.md"
                    $preview = Get-MarkdownPreview -Path $mdPath
                    $oneNoteLink = Get-OneNotePageLink -OneNoteApplication $oneNote -PageId $pageId

                    $indexEntries += [pscustomobject]@{
                        Section = $sectionName
                        Title = $pageName
                        RelativePath = $relativePath
                        Preview = $preview
                        OneNoteLink = $oneNoteLink
                    }
                }
            }

            if ($shouldPublish) {
                Write-Host "Exportiert: $sectionName / $pageName"
            }
            elseif ($pandocAvailable -and (Test-Path -LiteralPath $mdPath)) {
                Write-Host "Vorhanden, übersprungen: $sectionName / $pageName"
            }
        }
        catch {
            $failedCount++
            $errors += "Seite '$pageName' im Abschnitt '$sectionName' Section '$sectionName' OneNoteLink '$oneNoteLink'  RelativePath '$RelativePath' PageID '$pageId': $($_.Exception.Message)"
        }
    }
}

# Inhaltsverzeichnis erstellen
if ($pandocAvailable) {
    $indexPath = Join-Path $markdownRoot "inhaltsverzeichnis.md"
    $indexLines = @(
        "# Wissensindex: $notebookName",
        "",
        "Erstellt: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "",
        "Dieser Index verweist auf die einzelnen Markdown-Quellen unter `quellen/`.",
        "Zuerst anhand von Abschnitt, Seitentitel und Vorschau passende Einträge auswählen,",
        "danach die verlinkten Quelldateien lesen.",
        ""
    )

    $groupedEntries = $indexEntries | Sort-Object Section, Title | Group-Object Section

    foreach ($group in $groupedEntries) {
        $indexLines += "## $($group.Name)"
        $indexLines += ""

        foreach ($entry in $group.Group) {
            $title = ConvertTo-MarkdownText $entry.Title
            $preview = ConvertTo-MarkdownText $entry.Preview
            $link = $entry.RelativePath -replace ' ', '%20'

            $indexLines += "- [$title]($link)"
            if (-not [string]::IsNullOrWhiteSpace($preview)) {
                $indexLines += "  - Vorschau: $preview"
            }
            if (-not [string]::IsNullOrWhiteSpace($entry.OneNoteLink)) {
                $oneNoteUrl = $entry.OneNoteLink -replace ' ', '%20'
                $indexLines += "  - [OneNote-Originalseite]($oneNoteUrl)"
            }
        }

        $indexLines += ""
    }

    Set-Content -LiteralPath $indexPath -Value $indexLines -Encoding UTF8
}

# Fehlerprotokoll
if ($errors.Count -gt 0) {
    $logPath = Join-Path $notebookFolder "export-fehler.txt"
    $errors | Set-Content -LiteralPath $logPath -Encoding UTF8
}

# duration
$stopwatch.Stop()
$elapsed = $stopwatch.Elapsed
$elapsedText = "{0:00}:{1:00}:{2:00}" -f `
    [int]$elapsed.TotalHours, $elapsed.Minutes, $elapsed.Seconds
Write-Host "Dauer: $elapsedText (hh:mm:ss)"

Write-Host "`nExport abgeschlossen."
Write-Host "DOCX neu erstellt/überschrieben: $docxCount"

if ($pandocAvailable) {
    Write-Host "Markdown neu erstellt/aktualisiert: $markdownCount"
    Write-Host "Index: $(Join-Path $markdownRoot 'inhaltsverzeichnis.md')"
}
else {
    Write-Host "Markdown und Inhaltsverzeichnis wurden nicht erstellt (Pandoc nicht installiert)."
}

Write-Host "Übersprungene Seiten: $skippedCount"
Write-Host "Fehler: $failedCount"
Write-Host "Hauptordner: $notebookFolder"

if ($errors.Count -gt 0) {
    Write-Host "Fehlerprotokoll: $logPath"
}
