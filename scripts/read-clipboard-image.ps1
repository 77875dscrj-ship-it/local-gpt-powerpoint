param(
  [Parameter(Mandatory = $true)]
  [string]$OutputPath
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

function Write-ResultJson {
  param([Parameter(Mandatory = $true)]$Value)
  [Console]::Out.WriteLine(($Value | ConvertTo-Json -Compress -Depth 8))
}

function Convert-DibToBmpBytes {
  param([Parameter(Mandatory = $true)][byte[]]$DibBytes)

  if ($DibBytes.Length -lt 40) { return $null }
  $headerSize = [System.BitConverter]::ToUInt32($DibBytes, 0)
  $bitCount = [System.BitConverter]::ToUInt16($DibBytes, 14)
  $compression = [System.BitConverter]::ToUInt32($DibBytes, 16)
  $colorsUsed = 0
  if ($DibBytes.Length -ge 40) {
    $colorsUsed = [System.BitConverter]::ToUInt32($DibBytes, 32)
  }

  $paletteBytes = 0
  if ($colorsUsed -gt 0) {
    $paletteBytes = [int]$colorsUsed * 4
  } elseif ($bitCount -le 8) {
    $paletteBytes = [int]([Math]::Pow(2, $bitCount)) * 4
  }

  $maskBytes = 0
  if ($compression -eq 3 -and $headerSize -eq 40) {
    $maskBytes = 12
  }

  $pixelOffset = 14 + [int]$headerSize + $paletteBytes + $maskBytes
  $fileSize = 14 + $DibBytes.Length
  $bmpBytes = New-Object byte[] $fileSize
  $bmpBytes[0] = 0x42
  $bmpBytes[1] = 0x4D
  [System.BitConverter]::GetBytes([uint32]$fileSize).CopyTo($bmpBytes, 2)
  [System.BitConverter]::GetBytes([uint32]0).CopyTo($bmpBytes, 6)
  [System.BitConverter]::GetBytes([uint32]$pixelOffset).CopyTo($bmpBytes, 10)
  [System.Array]::Copy($DibBytes, 0, $bmpBytes, 14, $DibBytes.Length)
  return $bmpBytes
}

function Save-ImageObject {
  param(
    [Parameter(Mandatory = $true)]$Image,
    [Parameter(Mandatory = $true)][string]$Path
  )

  $Image.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
  return $true
}

function Try-SaveStreamImage {
  param(
    [Parameter(Mandatory = $true)][System.IO.Stream]$Stream,
    [Parameter(Mandatory = $true)][string]$Path
  )

  $position = $null
  if ($Stream.CanSeek) {
    $position = $Stream.Position
    $Stream.Position = 0
  }
  try {
    $image = [System.Drawing.Image]::FromStream($Stream)
    try {
      Save-ImageObject -Image $image -Path $Path | Out-Null
      return $true
    } finally {
      $image.Dispose()
    }
  } catch {
    return $false
  } finally {
    if ($null -ne $position -and $Stream.CanSeek) {
      $Stream.Position = $position
    }
  }
}

function Try-SaveBytesImage {
  param(
    [Parameter(Mandatory = $true)][byte[]]$Bytes,
    [Parameter(Mandatory = $true)][string]$Path,
    [bool]$IsDib = $false
  )

  $candidate = $Bytes
  if ($IsDib) {
    $candidate = Convert-DibToBmpBytes -DibBytes $Bytes
    if ($null -eq $candidate) { return $false }
  }

  $stream = New-Object System.IO.MemoryStream(,$candidate)
  try {
    return (Try-SaveStreamImage -Stream $stream -Path $Path)
  } finally {
    $stream.Dispose()
  }
}

function Try-SaveClipboardData {
  param(
    [Parameter(Mandatory = $true)]$Data,
    [Parameter(Mandatory = $true)][string]$Path,
    [bool]$IsDib = $false
  )

  if ($null -eq $Data) { return $false }
  if ($Data -is [System.Drawing.Image]) {
    Save-ImageObject -Image $Data -Path $Path | Out-Null
    return $true
  }
  if ($Data -is [System.IO.Stream]) {
    if ($IsDib) {
      $memory = New-Object System.IO.MemoryStream
      $position = $null
      if ($Data.CanSeek) {
        $position = $Data.Position
        $Data.Position = 0
      }
      try {
        $Data.CopyTo($memory)
        return (Try-SaveBytesImage -Bytes $memory.ToArray() -Path $Path -IsDib $true)
      } finally {
        if ($null -ne $position -and $Data.CanSeek) { $Data.Position = $position }
        $memory.Dispose()
      }
    }
    return (Try-SaveStreamImage -Stream $Data -Path $Path)
  }
  if ($Data -is [byte[]]) {
    return (Try-SaveBytesImage -Bytes $Data -Path $Path -IsDib $IsDib)
  }
  return $false
}

function Try-SaveClipboardFormat {
  param(
    [Parameter(Mandatory = $true)]$DataObject,
    [Parameter(Mandatory = $true)][string]$Format,
    [Parameter(Mandatory = $true)][string]$Path
  )

  try {
    if (-not $DataObject.GetDataPresent($Format)) { return $false }
    $data = $DataObject.GetData($Format)
    $isDib = $Format -eq [System.Windows.Forms.DataFormats]::Dib -or $Format -eq "DIB" -or $Format -eq "DeviceIndependentBitmap"
    return (Try-SaveClipboardData -Data $data -Path $Path -IsDib $isDib)
  } catch {
    return $false
  }
}

try {
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing

  $directory = [System.IO.Path]::GetDirectoryName($OutputPath)
  if (-not [System.IO.Directory]::Exists($directory)) {
    [System.IO.Directory]::CreateDirectory($directory) | Out-Null
  }

  $sourceFormat = $null

  if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
    $image = [System.Windows.Forms.Clipboard]::GetImage()
    if ($null -ne $image) {
      Save-ImageObject -Image $image -Path $OutputPath | Out-Null
      $sourceFormat = "ClipboardImage"
    }
  }

  if ($null -eq $sourceFormat -and [System.Windows.Forms.Clipboard]::ContainsFileDropList()) {
    $files = [System.Windows.Forms.Clipboard]::GetFileDropList()
    foreach ($file in $files) {
      $extension = [System.IO.Path]::GetExtension([string]$file).ToLowerInvariant()
      if ((".png", ".jpg", ".jpeg", ".bmp", ".gif") -contains $extension -and [System.IO.File]::Exists([string]$file)) {
        $imageFromFile = [System.Drawing.Image]::FromFile([string]$file)
        try {
          Save-ImageObject -Image $imageFromFile -Path $OutputPath | Out-Null
          $sourceFormat = "FileDrop"
          break
        } finally {
          $imageFromFile.Dispose()
        }
      }
    }
  }

  $formats = @()
  $dataObject = [System.Windows.Forms.Clipboard]::GetDataObject()
  if ($null -ne $dataObject) {
    $formats = @($dataObject.GetFormats())
  }

  if ($null -eq $sourceFormat -and $null -ne $dataObject) {
    $preferredFormats = @(
      "PNG",
      "image/png",
      "JFIF",
      "JPEG",
      "image/jpeg",
      [System.Windows.Forms.DataFormats]::Bitmap,
      [System.Windows.Forms.DataFormats]::Dib,
      "DIB",
      "DeviceIndependentBitmap"
    )
    foreach ($format in $preferredFormats) {
      if (Try-SaveClipboardFormat -DataObject $dataObject -Format $format -Path $OutputPath) {
        $sourceFormat = $format
        break
      }
    }
  }

  if ($null -eq $sourceFormat -and $null -ne $dataObject) {
    foreach ($format in $formats) {
      if (Try-SaveClipboardFormat -DataObject $dataObject -Format $format -Path $OutputPath) {
        $sourceFormat = $format
        break
      }
    }
  }

  if ($null -eq $sourceFormat) {
    Write-ResultJson @{ ok = $true; hasImage = $false; formats = $formats }
    exit 0
  }

  $item = Get-Item -LiteralPath $OutputPath
  Write-ResultJson @{
    ok = $true
    hasImage = $true
    path = $OutputPath
    mimeType = "image/png"
    sizeBytes = $item.Length
    sourceFormat = $sourceFormat
    formats = $formats
  }
  exit 0
} catch {
  Write-ResultJson @{ ok = $false; error = [string]$_.Exception.Message }
  exit 1
}
