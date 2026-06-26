param(
  [Parameter(Mandatory=$true)][string]$InputPath,
  [string]$FileName = "source"
)

$ErrorActionPreference = "Stop"
$OutputEncoding = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

$MaxChars = 120000

function Write-Json($value) {
  [Console]::Out.WriteLine(($value | ConvertTo-Json -Compress -Depth 8))
}

function Read-ZipEntryText($entry) {
  $stream = $entry.Open()
  try {
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
    try {
      return $reader.ReadToEnd()
    } finally {
      $reader.Dispose()
    }
  } finally {
    $stream.Dispose()
  }
}

function Get-XmlTextNodes([string]$xmlText) {
  if ([string]::IsNullOrWhiteSpace($xmlText)) { return @() }
  [xml]$xml = $xmlText
  $nodes = $xml.SelectNodes("//*[local-name()='t']")
  $out = New-Object System.Collections.Generic.List[string]
  foreach ($node in $nodes) {
    $value = [string]$node.InnerText
    if ($value.Trim().Length -gt 0) {
      [void]$out.Add($value.Trim())
    }
  }
  return $out
}

function Get-WordText($zip) {
  $entry = $zip.GetEntry("word/document.xml")
  if ($null -eq $entry) { return "" }
  [xml]$xml = Read-ZipEntryText $entry
  $parts = New-Object System.Collections.Generic.List[string]
  $paragraphs = $xml.SelectNodes("//*[local-name()='p']")
  foreach ($paragraph in $paragraphs) {
    $texts = New-Object System.Collections.Generic.List[string]
    foreach ($node in $paragraph.SelectNodes(".//*[local-name()='t']")) {
      [void]$texts.Add([string]$node.InnerText)
    }
    $line = (($texts -join "")).Trim()
    if ($line.Length -gt 0) {
      [void]$parts.Add($line)
    }
  }
  return ($parts -join "`r`n")
}

function Get-PowerPointText($zip) {
  $slides = $zip.Entries |
    Where-Object { $_.FullName -match '^ppt/slides/slide\d+\.xml$' } |
    Sort-Object {
      $m = [regex]::Match($_.FullName, 'slide(\d+)\.xml$')
      if ($m.Success) { [int]$m.Groups[1].Value } else { 0 }
    }
  $parts = New-Object System.Collections.Generic.List[string]
  foreach ($entry in $slides) {
    $m = [regex]::Match($entry.FullName, 'slide(\d+)\.xml$')
    $slideNo = if ($m.Success) { [int]$m.Groups[1].Value } else { $parts.Count + 1 }
    $texts = Get-XmlTextNodes (Read-ZipEntryText $entry)
    if ($texts.Count -gt 0) {
      [void]$parts.Add(("Slide {0}: {1}" -f $slideNo, (($texts -join " ").Trim())))
    }
  }
  return ($parts -join "`r`n`r`n")
}

function Get-SharedStrings($zip) {
  $entry = $zip.GetEntry("xl/sharedStrings.xml")
  if ($null -eq $entry) { return @() }
  [xml]$xml = Read-ZipEntryText $entry
  $shared = New-Object System.Collections.Generic.List[string]
  foreach ($si in $xml.SelectNodes("//*[local-name()='si']")) {
    $texts = New-Object System.Collections.Generic.List[string]
    foreach ($node in $si.SelectNodes(".//*[local-name()='t']")) {
      [void]$texts.Add([string]$node.InnerText)
    }
    [void]$shared.Add(($texts -join ""))
  }
  return $shared
}

function Get-ExcelText($zip) {
  $shared = Get-SharedStrings $zip
  $sheets = $zip.Entries |
    Where-Object { $_.FullName -match '^xl/worksheets/sheet\d+\.xml$' } |
    Sort-Object {
      $m = [regex]::Match($_.FullName, 'sheet(\d+)\.xml$')
      if ($m.Success) { [int]$m.Groups[1].Value } else { 0 }
    }
  $parts = New-Object System.Collections.Generic.List[string]
  foreach ($entry in $sheets) {
    $m = [regex]::Match($entry.FullName, 'sheet(\d+)\.xml$')
    $sheetNo = if ($m.Success) { [int]$m.Groups[1].Value } else { $parts.Count + 1 }
    [xml]$xml = Read-ZipEntryText $entry
    [void]$parts.Add(("Sheet {0}" -f $sheetNo))
    foreach ($row in $xml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")) {
      $values = New-Object System.Collections.Generic.List[string]
      foreach ($cell in $row.ChildNodes) {
        if ($cell.LocalName -ne "c") { continue }
        $cellType = $cell.GetAttribute("t")
        $value = ""
        if ($cellType -eq "inlineStr") {
          $inline = $cell.SelectNodes(".//*[local-name()='t']")
          $inlineTexts = New-Object System.Collections.Generic.List[string]
          foreach ($node in $inline) { [void]$inlineTexts.Add([string]$node.InnerText) }
          $value = ($inlineTexts -join "")
        } else {
          $vNode = $cell.SelectSingleNode("*[local-name()='v']")
          if ($null -ne $vNode) { $value = [string]$vNode.InnerText }
          if ($cellType -eq "s" -and $value -match '^\d+$') {
            $idx = [int]$value
            if ($idx -ge 0 -and $idx -lt $shared.Count) {
              $value = [string]$shared[$idx]
            }
          }
        }
        if ($value.Trim().Length -gt 0) {
          [void]$values.Add($value.Trim())
        }
      }
      if ($values.Count -gt 0) {
        [void]$parts.Add(($values -join "`t"))
      }
    }
    [void]$parts.Add("")
  }
  return ($parts -join "`r`n")
}

function Read-TextFile([string]$path, [string]$extension) {
  $text = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
  if ($extension -eq ".html" -or $extension -eq ".htm") {
    $text = [regex]::Replace($text, '(?is)<(script|style).*?</\1>', ' ')
    $text = [regex]::Replace($text, '(?s)<[^>]+>', ' ')
    $text = [System.Net.WebUtility]::HtmlDecode($text)
  }
  return $text
}

function Read-PdfFile([string]$path) {
  $candidates = @(
    "C:\Users\saman\.cache\codex-runtimes\codex-primary-runtime\dependencies\bin\pdftotext.exe",
    "pdftotext.exe"
  )
  $tool = $null
  foreach ($candidate in $candidates) {
    if ($candidate -eq "pdftotext.exe") {
      try {
        $cmd = Get-Command $candidate -ErrorAction Stop
        $tool = $cmd.Source
        break
      } catch {}
    } elseif (Test-Path -LiteralPath $candidate) {
      $tool = $candidate
      break
    }
  }
  if ($null -eq $tool) {
    throw "PDF extraction needs pdftotext, but it was not found."
  }
  $text = & $tool -layout -enc UTF-8 $path -
  return ($text -join "`r`n")
}

try {
  if (!(Test-Path -LiteralPath $InputPath)) {
    throw "Source file was not found."
  }

  $extension = [System.IO.Path]::GetExtension($FileName).ToLowerInvariant()
  $text = ""
  if ($extension -in @(".txt", ".md", ".csv", ".tsv", ".json", ".xml", ".html", ".htm", ".log")) {
    $text = Read-TextFile $InputPath $extension
  } elseif ($extension -eq ".pdf") {
    $text = Read-PdfFile $InputPath
  } elseif ($extension -in @(".docx", ".pptx", ".xlsx")) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($InputPath)
    try {
      if ($extension -eq ".docx") {
        $text = Get-WordText $zip
      } elseif ($extension -eq ".pptx") {
        $text = Get-PowerPointText $zip
      } elseif ($extension -eq ".xlsx") {
        $text = Get-ExcelText $zip
      }
    } finally {
      $zip.Dispose()
    }
  } else {
    throw "Unsupported source file type: $extension"
  }

  $text = [regex]::Replace(([string]$text), "[`t ]{2,}", " ").Trim()
  $truncated = $false
  if ($text.Length -gt $MaxChars) {
    $text = $text.Substring(0, $MaxChars)
    $truncated = $true
  }
  Write-Json @{ ok = $true; name = $FileName; kind = $extension; text = $text; chars = $text.Length; truncated = $truncated }
  exit 0
} catch {
  Write-Json @{ ok = $false; error = [string]$_.Exception.Message }
  exit 1
}
