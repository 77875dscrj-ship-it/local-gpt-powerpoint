param(
  [Parameter(Mandatory=$true)][string]$Action,
  [string]$InputPath = ""
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

function Write-Json($value) {
  [Console]::Out.WriteLine(($value | ConvertTo-Json -Compress -Depth 12))
}

function Rgb([int]$r, [int]$g, [int]$b) {
  return ($r + ($g * 256) + ($b * 65536))
}

function New-LocalGptId([string]$prefix) {
  return ($prefix + "-" + ([guid]::NewGuid().ToString("N").Substring(0, 12))).ToUpperInvariant()
}

function Set-ShapeTag($shape, [string]$name, [string]$value) {
  if ($null -eq $shape -or [string]::IsNullOrWhiteSpace($name)) { return }
  try { $shape.Tags.Delete($name) | Out-Null } catch {}
  try { $shape.Tags.Add($name, $value) | Out-Null } catch {}
}

function Set-LocalGptTags($shape, [string]$role) {
  if ($null -eq $shape) { return }
  $existingId = ""
  try { $existingId = [string]$shape.Tags.Item("LOCALGPT_ID") } catch {}
  if ([string]::IsNullOrWhiteSpace($existingId)) {
    Set-ShapeTag $shape "LOCALGPT_ID" (New-LocalGptId "SHAPE")
  }
  Set-ShapeTag $shape "LOCALGPT_OWNER" "LOCALGPT"
  if (![string]::IsNullOrWhiteSpace($role)) {
    Set-ShapeTag $shape "LOCALGPT_ROLE" $role
  }
}

function Get-ShapeTags($shape) {
  $tags = @{}
  try {
    for ($i = 1; $i -le $shape.Tags.Count; $i++) {
      $name = [string]$shape.Tags.Name($i)
      if (![string]::IsNullOrWhiteSpace($name)) {
        $tags[$name] = [string]$shape.Tags.Value($i)
      }
    }
  } catch {}
  return $tags
}

function Get-PowerPointApp {
  try {
    return [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
  } catch {
    throw "실행 중인 PowerPoint를 찾지 못했습니다."
  }
}

function Get-ShapeText($shape) {
  try {
    if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
      return [string]$shape.TextFrame.TextRange.Text
    }
  } catch {}
  return ""
}

function Get-FontSize($shape) {
  try {
    if ($shape.HasTextFrame -eq -1) {
      return [double]$shape.TextFrame.TextRange.Font.Size
    }
  } catch {}
  return 0
}

function Get-FontName($shape) {
  try {
    if ($shape.HasTextFrame -eq -1) {
      return [string]$shape.TextFrame.TextRange.Font.Name
    }
  } catch {}
  return ""
}

function Get-TextStyleInfo($shape) {
  $info = @{
    fontName = ""
    fontSize = 0
    bold = $null
    fontRgb = $null
    lineSpacing = $null
    spaceBefore = $null
    spaceAfter = $null
    lineRuleWithin = $null
    lineRuleBefore = $null
    lineRuleAfter = $null
  }
  try {
    if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
      $range = $shape.TextFrame.TextRange
      $info.fontName = [string]$range.Font.Name
      $info.fontSize = [math]::Round([double]$range.Font.Size, 2)
      $info.bold = ([int]$range.Font.Bold -eq -1)
      $info.fontRgb = Get-OptionalInt { $range.Font.Color.RGB }
      $para = $range.ParagraphFormat
      $info.lineSpacing = Get-OptionalDouble { $para.SpaceWithin }
      $info.spaceBefore = Get-OptionalDouble { $para.SpaceBefore }
      $info.spaceAfter = Get-OptionalDouble { $para.SpaceAfter }
      $info.lineRuleWithin = Get-OptionalInt { $para.LineRuleWithin }
      $info.lineRuleBefore = Get-OptionalInt { $para.LineRuleBefore }
      $info.lineRuleAfter = Get-OptionalInt { $para.LineRuleAfter }
    }
  } catch {}
  return $info
}

function Get-OptionalInt($scriptBlock) {
  try {
    $value = & $scriptBlock
    if ($null -ne $value) { return [int]$value }
  } catch {}
  return $null
}

function Get-OptionalDouble($scriptBlock) {
  try {
    $value = & $scriptBlock
    if ($null -ne $value) { return [math]::Round([double]$value, 2) }
  } catch {}
  return $null
}

function Get-OptionalString($scriptBlock) {
  try {
    $value = & $scriptBlock
    if ($null -ne $value) { return [string]$value }
  } catch {}
  return $null
}

function Limit-Text([string]$text, [int]$maxChars) {
  if ([string]::IsNullOrEmpty($text)) { return "" }
  if ($text.Length -le $maxChars) { return $text }
  return $text.Substring(0, $maxChars) + "..."
}

function Set-TextStyle($shape, [double]$fontSize, [bool]$bold) {
  try {
    $shape.TextFrame.TextRange.Font.Name = "맑은 고딕"
    if ($fontSize -gt 0) { $shape.TextFrame.TextRange.Font.Size = $fontSize }
    if ($bold) { $shape.TextFrame.TextRange.Font.Bold = -1 }
  } catch {}
}

function Add-TextBox($slide, [string]$text, [double]$left, [double]$top, [double]$width, [double]$height, [double]$fontSize, [bool]$bold, [int]$color) {
  $shape = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
  Set-LocalGptTags $shape "TEXT"
  $shape.TextFrame.WordWrap = -1
  $shape.TextFrame.TextRange.Text = $text
  Set-TextStyle $shape $fontSize $bold
  try { $shape.TextFrame.TextRange.Font.Color.RGB = $color } catch {}
  return $shape
}

function Add-Rect($slide, [double]$left, [double]$top, [double]$width, [double]$height, [int]$fill, [int]$line, [double]$transparency) {
  $shape = $slide.Shapes.AddShape(1, $left, $top, $width, $height)
  Set-LocalGptTags $shape "SHAPE"
  $shape.Fill.ForeColor.RGB = $fill
  $shape.Fill.Transparency = $transparency
  if ($line -lt 0) {
    $shape.Line.Visible = 0
  } else {
    $shape.Line.Visible = -1
    $shape.Line.ForeColor.RGB = $line
  }
  return $shape
}

function Clear-Slide($slide) {
  for ($i = $slide.Shapes.Count; $i -ge 1; $i--) {
    try { $slide.Shapes.Item($i).Delete() } catch {}
  }
}

function Get-ActiveSlide($ppt) {
  return $ppt.ActiveWindow.View.Slide
}

function Get-SlideById($pres, [int]$slideId) {
  for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $slide = $pres.Slides.Item($i)
    if ([int]$slide.SlideID -eq $slideId) { return $slide }
  }
  return $null
}

function Export-SlideImage($slide, [string]$path) {
  $dir = [System.IO.Path]::GetDirectoryName($path)
  if (![System.IO.Directory]::Exists($dir)) {
    [System.IO.Directory]::CreateDirectory($dir) | Out-Null
  }
  $firstError = ""
  try {
    $slide.Export($path, "PNG", 1280, 720) | Out-Null
  } catch {
    $firstError = [string]$_.Exception.Message
    try {
      $slide.Export($path, "PNG") | Out-Null
    } catch {
      throw ("Slide.Export failed path=" + $path + " slideIndex=" + [string]$slide.SlideIndex + " first=" + $firstError + " second=" + [string]$_.Exception.Message)
    }
  }
  return $path
}

function Get-SlideTitle($slide) {
  try {
    $title = Get-ShapeText $slide.Shapes.Title
    if ($title.Trim().Length -gt 0) { return $title.Trim() }
  } catch {}
  return ""
}

function Get-ActiveSlideText($slide) {
  $parts = New-Object System.Collections.Generic.List[string]
  for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
    $shape = $slide.Shapes.Item($i)
    $text = Get-ShapeText $shape
    if ($text.Trim().Length -gt 0) {
      [void]$parts.Add($text.Trim())
    }
  }
  return ($parts -join "`r`n`r`n")
}

function Get-DeckText($ppt) {
  $parts = New-Object System.Collections.Generic.List[string]
  $slides = $ppt.ActivePresentation.Slides
  for ($i = 1; $i -le $slides.Count; $i++) {
    $slide = $slides.Item($i)
    $title = Get-SlideTitle $slide
    $text = Get-ActiveSlideText $slide
    if ($text.Trim().Length -gt 0) {
      [void]$parts.Add(("Slide {0}: {1}`r`n{2}" -f $i, $title, $text.Trim()))
    } else {
      [void]$parts.Add(("Slide {0}: {1}" -f $i, $title))
    }
  }
  return ($parts -join "`r`n`r`n---`r`n`r`n")
}

function Get-ShapeInfo($shape) {
  $text = Get-ShapeText $shape
  $textStyle = Get-TextStyleInfo $shape
  $placeholderType = $null
  $isPlaceholder = $false
  try {
    $placeholderType = [int]$shape.PlaceholderFormat.Type
    $isPlaceholder = $true
  } catch {}
  return @{
    id = $shape.Id
    name = [string]$shape.Name
    type = Get-OptionalInt { $shape.Type }
    autoShapeType = Get-OptionalInt { $shape.AutoShapeType }
    placeholderType = $placeholderType
    isPlaceholder = $isPlaceholder
    zOrderPosition = Get-OptionalInt { $shape.ZOrderPosition }
    left = [math]::Round([double]$shape.Left, 1)
    top = [math]::Round([double]$shape.Top, 1)
    width = [math]::Round([double]$shape.Width, 1)
    height = [math]::Round([double]$shape.Height, 1)
    rotation = Get-OptionalDouble { $shape.Rotation }
    text = $text
    textLength = $text.Length
    textPreview = Limit-Text $text 360
    fontName = $textStyle.fontName
    fontSize = $textStyle.fontSize
    bold = $textStyle.bold
    fontRgb = $textStyle.fontRgb
    paragraph = @{
      lineSpacing = $textStyle.lineSpacing
      spaceBefore = $textStyle.spaceBefore
      spaceAfter = $textStyle.spaceAfter
      lineRuleWithin = $textStyle.lineRuleWithin
      lineRuleBefore = $textStyle.lineRuleBefore
      lineRuleAfter = $textStyle.lineRuleAfter
    }
    hasTextFrame = ($shape.HasTextFrame -eq -1)
    hasTable = (Get-OptionalInt { $shape.HasTable }) -eq -1
    hasChart = (Get-OptionalInt { $shape.HasChart }) -eq -1
    fillRgb = Get-OptionalInt { $shape.Fill.ForeColor.RGB }
    lineRgb = Get-OptionalInt { $shape.Line.ForeColor.RGB }
    altText = Get-OptionalString { $shape.AlternativeText }
    tags = Get-ShapeTags $shape
  }
}

function Get-LayoutInfo($layout) {
  if ($null -eq $layout) { return $null }
  $placeholderCount = 0
  $textPlaceholderCount = 0
  $shapeCount = 0
  try { $shapeCount = [int]$layout.Shapes.Count } catch {}
  $sampleShapes = New-Object System.Collections.Generic.List[object]
  $maxShapes = [Math]::Min($shapeCount, 16)
  for ($i = 1; $i -le $maxShapes; $i++) {
    try {
      $shape = $layout.Shapes.Item($i)
      $isPlaceholder = $false
      $placeholderType = $null
      try {
        $placeholderType = [int]$shape.PlaceholderFormat.Type
        $isPlaceholder = $true
        $placeholderCount++
      } catch {}
      if ($shape.HasTextFrame -eq -1) { $textPlaceholderCount++ }
      [void]$sampleShapes.Add(@{
        index = $i
        name = [string]$shape.Name
        type = Get-OptionalInt { $shape.Type }
        placeholderType = $placeholderType
        isPlaceholder = $isPlaceholder
        left = Get-OptionalDouble { $shape.Left }
        top = Get-OptionalDouble { $shape.Top }
        width = Get-OptionalDouble { $shape.Width }
        height = Get-OptionalDouble { $shape.Height }
        textPreview = Limit-Text (Get-ShapeText $shape) 120
      })
    } catch {}
  }
  return @{
    name = Get-OptionalString { $layout.Name }
    index = Get-OptionalInt { $layout.Index }
    matchingName = Get-OptionalString { $layout.MatchingName }
    shapeCount = $shapeCount
    placeholderCount = $placeholderCount
    textPlaceholderCount = $textPlaceholderCount
    sampleShapes = $sampleShapes
  }
}

function Get-DesignInfo($design) {
  if ($null -eq $design) { return $null }
  $layouts = New-Object System.Collections.Generic.List[object]
  $layoutCount = 0
  try { $layoutCount = [int]$design.SlideMaster.CustomLayouts.Count } catch {}
  $maxLayouts = [Math]::Min($layoutCount, 20)
  for ($i = 1; $i -le $maxLayouts; $i++) {
    try { [void]$layouts.Add((Get-LayoutInfo $design.SlideMaster.CustomLayouts.Item($i))) } catch {}
  }
  return @{
    name = Get-OptionalString { $design.Name }
    index = Get-OptionalInt { $design.Index }
    slideMasterName = Get-OptionalString { $design.SlideMaster.Name }
    masterShapeCount = Get-OptionalInt { $design.SlideMaster.Shapes.Count }
    customLayoutCount = $layoutCount
    layouts = $layouts
  }
}

function Get-ThemeInfo($pres) {
  $theme = @{
    name = ""
    bodyFont = ""
    headingFont = ""
    colors = @()
  }
  try { $theme.name = [string]$pres.SlideMaster.Theme.Name } catch {}
  try { $theme.bodyFont = [string]$pres.SlideMaster.Theme.ThemeFontScheme.MajorFont.Item(1).Name } catch {}
  try { $theme.headingFont = [string]$pres.SlideMaster.Theme.ThemeFontScheme.MinorFont.Item(1).Name } catch {}
  try {
    $colors = New-Object System.Collections.Generic.List[object]
    for ($i = 1; $i -le 12; $i++) {
      try {
        $color = $pres.SlideMaster.Theme.ThemeColorScheme.Colors($i)
        [void]$colors.Add(@{ index = $i; rgb = Get-OptionalInt { $color.RGB } })
      } catch {}
    }
    $theme.colors = $colors
  } catch {}
  return $theme
}

function Get-TemplateContext($ppt) {
  $pres = $ppt.ActivePresentation
  $designs = New-Object System.Collections.Generic.List[object]
  $designCount = 0
  try { $designCount = [int]$pres.Designs.Count } catch {}
  $maxDesigns = [Math]::Min($designCount, 8)
  for ($i = 1; $i -le $maxDesigns; $i++) {
    try { [void]$designs.Add((Get-DesignInfo $pres.Designs.Item($i))) } catch {}
  }

  $usedLayouts = @{}
  $usedDesigns = @{}
  for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    try {
      $slide = $pres.Slides.Item($i)
      $layoutName = Get-OptionalString { $slide.CustomLayout.Name }
      $designName = Get-OptionalString { $slide.Design.Name }
      if (![string]::IsNullOrWhiteSpace($layoutName)) {
        if (!$usedLayouts.ContainsKey($layoutName)) { $usedLayouts[$layoutName] = 0 }
        $usedLayouts[$layoutName] = [int]$usedLayouts[$layoutName] + 1
      }
      if (![string]::IsNullOrWhiteSpace($designName)) {
        if (!$usedDesigns.ContainsKey($designName)) { $usedDesigns[$designName] = 0 }
        $usedDesigns[$designName] = [int]$usedDesigns[$designName] + 1
      }
    } catch {}
  }

  return @{
    theme = Get-ThemeInfo $pres
    designCount = $designCount
    designs = $designs
    usedLayouts = $usedLayouts
    usedDesigns = $usedDesigns
  }
}

function Get-SlideShapeMap($slide) {
  $shapes = New-Object System.Collections.Generic.List[object]
  for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
    $shape = $slide.Shapes.Item($i)
    $info = Get-ShapeInfo $shape
    $info.shapeIndex = $i
    [void]$shapes.Add($info)
  }
  return @{
    slideIndex = $slide.SlideIndex
    slideId = $slide.SlideID
    slideName = [string]$slide.Name
    title = Get-SlideTitle $slide
    designName = Get-OptionalString { $slide.Design.Name }
    layoutName = Get-OptionalString { $slide.CustomLayout.Name }
    layoutIndex = Get-OptionalInt { $slide.CustomLayout.Index }
    shapeCount = $slide.Shapes.Count
    shapes = $shapes
  }
}

function Get-ShapeMap($ppt, [string]$jsonText) {
  $payload = @{}
  if (![string]::IsNullOrWhiteSpace($jsonText)) {
    $payload = $jsonText | ConvertFrom-Json
  }
  $pres = $ppt.ActivePresentation
  $scope = "active"
  try {
    if (![string]::IsNullOrWhiteSpace([string]$payload.scope)) { $scope = [string]$payload.scope }
  } catch {}
  $maxSlides = 12
  try {
    if ($null -ne $payload.maxSlides) { $maxSlides = [Math]::Max(1, [Math]::Min(200, [int]$payload.maxSlides)) }
  } catch {}
  $targetSlideIndex = 0
  try { $targetSlideIndex = [int]$payload.slideIndex } catch {}

  $slides = New-Object System.Collections.Generic.List[object]
  if ($scope -eq "deck") {
    $count = [Math]::Min($pres.Slides.Count, $maxSlides)
    for ($i = 1; $i -le $count; $i++) {
      [void]$slides.Add((Get-SlideShapeMap $pres.Slides.Item($i)))
    }
  } elseif ($targetSlideIndex -ge 1 -and $targetSlideIndex -le $pres.Slides.Count) {
    [void]$slides.Add((Get-SlideShapeMap $pres.Slides.Item($targetSlideIndex)))
  } else {
    [void]$slides.Add((Get-SlideShapeMap (Get-ActiveSlide $ppt)))
  }

  return @{
    ok = $true
    scope = $scope
    presentationName = [string]$pres.Name
    presentationFullName = [string]$pres.FullName
    slideCount = $pres.Slides.Count
    slideWidth = [math]::Round([double]$pres.PageSetup.SlideWidth, 1)
    slideHeight = [math]::Round([double]$pres.PageSetup.SlideHeight, 1)
    slides = $slides
  }
}

function Get-SelectionContext($ppt) {
  $sel = $ppt.ActiveWindow.Selection
  $slide = Get-ActiveSlide $ppt
  $info = @{
    type = [int]$sel.Type
    rawType = [int]$sel.Type
    slideIndex = $slide.SlideIndex
    slideId = $slide.SlideID
    text = ""
    textSelection = ""
    selectionMode = "none"
    shapes = @()
  }

  if ($sel.Type -eq 3) {
    $selectedText = ""
    try { $selectedText = [string]$sel.TextRange.Text } catch {}
    $info.text = $selectedText
    $info.textSelection = $selectedText
    $info.selectionMode = "text"

    # In PowerPoint, clicking inside a text box can report "text selection"
    # even though the visible block/shape is selected. Walk back to that parent shape.
    $parentShape = $null
    try { $parentShape = $sel.TextRange.Parent.Parent } catch {}
    if ($null -ne $parentShape) {
      $shapeInfo = Get-ShapeInfo $parentShape
      $info.shapes = @($shapeInfo)
      if ([string]::IsNullOrWhiteSpace($info.text)) {
        $info.type = 2
        $info.text = [string]$shapeInfo.text
        $info.selectionMode = "text_cursor_parent_shape"
      } else {
        $info.selectionMode = "text_range_parent_shape"
      }
    }
    return $info
  }

  if ($sel.Type -eq 2) {
    $shapes = New-Object System.Collections.Generic.List[object]
    for ($i = 1; $i -le $sel.ShapeRange.Count; $i++) {
      $shape = $sel.ShapeRange.Item($i)
      [void]$shapes.Add((Get-ShapeInfo $shape))
    }
    $info.shapes = $shapes
    $info.text = (($shapes | ForEach-Object { $_.text } | Where-Object { $_ }) -join "`r`n`r`n")
    $info.selectionMode = "shape_range"
    return $info
  }

  return $info
}

function Get-SlideSnapshot($slide) {
  $text = Get-ActiveSlideText $slide
  return @{
    slideIndex = $slide.SlideIndex
    slideId = $slide.SlideID
    name = [string]$slide.Name
    title = Get-SlideTitle $slide
    text = $text
    textLength = $text.Length
    shapeCount = $slide.Shapes.Count
  }
}

function Get-SlidesSnapshot($ppt) {
  $items = New-Object System.Collections.Generic.List[object]
  $slides = $ppt.ActivePresentation.Slides
  for ($i = 1; $i -le $slides.Count; $i++) {
    [void]$items.Add((Get-SlideSnapshot $slides.Item($i)))
  }
  return $items
}

function Get-Context($ppt) {
  $pres = $ppt.ActivePresentation
  $slide = Get-ActiveSlide $ppt
  return @{
    ok = $true
    officeVersion = [string]$ppt.Version
    officeName = [string]$ppt.Name
    presentationName = [string]$pres.Name
    presentationFullName = [string]$pres.FullName
    slideIndex = $slide.SlideIndex
    slideId = $slide.SlideID
    slideCount = $pres.Slides.Count
    slideWidth = [math]::Round([double]$pres.PageSetup.SlideWidth, 1)
    slideHeight = [math]::Round([double]$pres.PageSetup.SlideHeight, 1)
    activeSlideText = Get-ActiveSlideText $slide
    deckText = Get-DeckText $ppt
    slides = Get-SlidesSnapshot $ppt
    template = Get-TemplateContext $ppt
    activeSlideShapeMap = Get-SlideShapeMap $slide
    selection = Get-SelectionContext $ppt
  }
}

function Get-BodyShape($slide) {
  $best = $null
  $bestArea = 0
  for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
    $shape = $slide.Shapes.Item($i)
    try {
      if ($shape.HasTextFrame -eq -1) {
        $area = [double]$shape.Width * [double]$shape.Height
        if ($area -gt $bestArea -and [double]$shape.Top -gt 80) {
          $best = $shape
          $bestArea = $area
        }
      }
    } catch {}
  }
  return $best
}

function Get-TitleShape($slide) {
  try {
    $candidate = $slide.Shapes.Title
    if ($null -ne $candidate -and $candidate.HasTextFrame -eq -1) {
      return $candidate
    }
  } catch {}

  $best = $null
  $bestScore = -1
  for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
    $shape = $slide.Shapes.Item($i)
    try {
      if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
        $text = (Get-ShapeText $shape).Trim()
        $fontSize = Get-FontSize $shape
        $top = [double]$shape.Top
        if ($text.Length -gt 0 -and $top -lt 175 -and $fontSize -ge 18) {
          $score = (200 - $top) + ($fontSize * 3) + ([double]$shape.Width / 100)
          if ($score -gt $bestScore) {
            $best = $shape
            $bestScore = $score
          }
        }
      }
    } catch {}
  }
  return $best
}

function Set-SlideTitle($slide, [string]$text) {
  $candidate = Get-TitleShape $slide
  if ($null -eq $candidate) {
    $candidate = Add-TextBox $slide "" 54 34 610 52 28 $true (Rgb 36 41 54)
  }
  Set-LocalGptTags $candidate "TITLE"
  $candidate.TextFrame.TextRange.Text = $text
  Set-TextStyle $candidate 28 $true
  return $candidate.Name
}

function Set-SlideBody($slide, [string]$text) {
  $shape = Get-BodyShape $slide
  if ($null -eq $shape) {
    $shape = Add-TextBox $slide "" 76 130 560 280 20 $false (Rgb 42 46 56)
  }
  Set-LocalGptTags $shape "BODY"
  $shape.TextFrame.TextRange.Text = $text
  Set-TextStyle $shape 20 $false
  return $shape.Name
}

function Set-SpeakerNotes($slide, [string]$text) {
  try {
    $notesShape = $slide.NotesPage.Shapes.Placeholders.Item(2)
    $notesShape.TextFrame.TextRange.Text = $text
    return $notesShape.Name
  } catch {}
  for ($i = 1; $i -le $slide.NotesPage.Shapes.Count; $i++) {
    $shape = $slide.NotesPage.Shapes.Item($i)
    try {
      if ($shape.HasTextFrame -eq -1) {
        $shape.TextFrame.TextRange.Text = $text
        return $shape.Name
      }
    } catch {}
  }
  throw "발표자 노트 영역을 찾지 못했습니다."
}

function Get-SlideSpecArray($value) {
  if ($null -eq $value) { return @() }
  if ($value -is [System.Array]) { return @($value) }
  return @($value)
}

function Get-BulletText($spec) {
  $bullets = @()
  if ($null -ne $spec.bullets) {
    $bullets = @($spec.bullets) | ForEach-Object { [string]$_ } | Where-Object { $_.Trim().Length -gt 0 }
  } elseif ($null -ne $spec.body) {
    $bullets = @([string]$spec.body)
  }
  return ($bullets -join "`r`n")
}

function Add-StyledSlide($pres, [int]$index, $spec, [bool]$reuseExisting) {
  if ($reuseExisting) {
    $slide = $pres.Slides.Item($index)
    Clear-Slide $slide
  } else {
    $slide = $pres.Slides.Add($index, 12)
  }

  $w = [double]$pres.PageSetup.SlideWidth
  $h = [double]$pres.PageSetup.SlideHeight
  $accent = Rgb 86 69 160
  $ink = Rgb 33 37 48
  $muted = Rgb 91 96 112
  $soft = Rgb 245 243 252
  $line = Rgb 216 210 236
  $kind = ([string]$spec.kind).ToLowerInvariant()
  $title = [string]$spec.title
  $subtitle = [string]$spec.subtitle
  $body = Get-BulletText $spec

  Add-Rect $slide 0 0 $w $h (Rgb 255 255 255) -1 0 | Out-Null
  Add-Rect $slide 0 0 $w 16 $accent -1 0 | Out-Null

  if ($kind -eq "title") {
    Add-Rect $slide 44 74 ($w - 88) 260 $soft $line 0 | Out-Null
    Add-TextBox $slide $title 70 116 ($w - 140) 78 34 $true $ink | Out-Null
    if (![string]::IsNullOrWhiteSpace($subtitle)) {
      Add-TextBox $slide $subtitle 72 202 ($w - 144) 80 18 $false $muted | Out-Null
    }
    Add-Rect $slide 72 304 90 5 $accent -1 0 | Out-Null
  } elseif ($kind -eq "section") {
    Add-Rect $slide 0 0 ($w * 0.34) $h $accent -1 0 | Out-Null
    Add-TextBox $slide $title 46 88 ($w * 0.28) 230 32 $true (Rgb 255 255 255) | Out-Null
    if ($body) {
      Add-TextBox $slide $body ($w * 0.40) 114 ($w * 0.50) 250 22 $false $ink | Out-Null
    }
  } elseif ($kind -eq "comparison") {
    Add-TextBox $slide $title 54 34 ($w - 108) 50 28 $true $ink | Out-Null
    Add-Rect $slide 54 92 ($w - 108) 2 $accent -1 0 | Out-Null
    $leftBody = $body
    Add-Rect $slide 58 126 (($w - 148) / 2) 270 $soft $line 0 | Out-Null
    Add-Rect $slide (($w / 2) + 16) 126 (($w - 148) / 2) 270 (Rgb 250 251 253) $line 0 | Out-Null
    Add-TextBox $slide $leftBody 82 154 (($w - 196) / 2) 220 18 $false $ink | Out-Null
  } else {
    Add-TextBox $slide $title 54 32 ($w - 108) 54 28 $true $ink | Out-Null
    Add-Rect $slide 54 92 ($w - 108) 2 $accent -1 0 | Out-Null
    if (![string]::IsNullOrWhiteSpace($subtitle)) {
      Add-TextBox $slide $subtitle 58 106 ($w - 116) 38 15 $false $muted | Out-Null
    }
    Add-TextBox $slide $body 78 150 ($w - 156) 250 20 $false $ink | Out-Null
  }

  $footer = Add-TextBox $slide ("Local GPT  |  " + $slide.SlideIndex) ($w - 170) ($h - 28) 130 18 8 $false $muted
  try { $footer.TextFrame.TextRange.ParagraphFormat.Alignment = 3 } catch {}
  if (![string]::IsNullOrWhiteSpace([string]$spec.notes)) {
    Set-SpeakerNotes $slide ([string]$spec.notes) | Out-Null
  }
  return $slide
}

function Replace-Deck($ppt, $slides, $presOverride = $null) {
  $pres = $ppt.ActivePresentation
  if ($null -ne $presOverride) { $pres = $presOverride }
  if ($pres.Slides.Count -lt 1) {
    $pres.Slides.Add(1, 12) | Out-Null
  }
  while ($pres.Slides.Count -gt 1) {
    $pres.Slides.Item($pres.Slides.Count).Delete()
  }
  $list = Get-SlideSpecArray $slides
  if ($list.Count -eq 0) { throw "생성할 슬라이드가 없습니다." }
  $results = New-Object System.Collections.Generic.List[object]
  for ($i = 0; $i -lt $list.Count; $i++) {
    $slide = Add-StyledSlide $pres ($i + 1) $list[$i] ($i -eq 0)
    [void]$results.Add(@{ type = "slide"; slide = $slide.SlideIndex; target = "styled_slide" })
  }
  if ($null -eq $presOverride) { $pres.Slides.Item(1).Select() }
  return $results
}

function Add-Slides($ppt, [int]$afterIndex, $slides, $presOverride = $null) {
  $pres = $ppt.ActivePresentation
  if ($null -ne $presOverride) { $pres = $presOverride }
  if ($afterIndex -lt 0 -or $afterIndex -gt $pres.Slides.Count) {
    if ($null -eq $presOverride) {
      $afterIndex = (Get-ActiveSlide $ppt).SlideIndex
    } else {
      $afterIndex = $pres.Slides.Count
    }
  }
  $list = Get-SlideSpecArray $slides
  $results = New-Object System.Collections.Generic.List[object]
  $insertAt = $afterIndex + 1
  foreach ($spec in $list) {
    $slide = Add-StyledSlide $pres $insertAt $spec $false
    [void]$results.Add(@{ type = "slide"; slide = $slide.SlideIndex; target = "styled_slide" })
    $insertAt += 1
  }
  if ($results.Count -gt 0 -and $null -eq $presOverride) {
    $pres.Slides.Item($results[$results.Count - 1].slide).Select()
  }
  return $results
}

function Replace-SlideText($slide, [string]$find, [string]$replace) {
  $changed = 0
  if ([string]::IsNullOrEmpty($find)) { return $changed }
  for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
    $shape = $slide.Shapes.Item($i)
    try {
      if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
        $text = [string]$shape.TextFrame.TextRange.Text
        if ($text.Contains($find)) {
          $shape.TextFrame.TextRange.Text = $text.Replace($find, $replace)
          $changed += 1
        }
      }
    } catch {}
  }
  return $changed
}

function Format-Selection($ppt, $action) {
  $sel = $ppt.ActiveWindow.Selection
  $targets = New-Object System.Collections.Generic.List[object]
  if ($sel.Type -eq 3) {
    $changed = New-Object System.Collections.Generic.List[string]
    try {
      if ($null -ne $action.fontSize) { $sel.TextRange.Font.Size = [double]$action.fontSize; Add-ChangedProperty $changed "fontSize" }
      if ($null -ne $action.bold) { $sel.TextRange.Font.Bold = if ([bool]$action.bold) { -1 } else { 0 }; Add-ChangedProperty $changed "bold" }
      Apply-ParagraphFormat $sel.TextRange $action $changed
    } catch {}
    [void]$targets.Add(@{ target = "selected_text"; slide = (Get-ActiveSlide $ppt).SlideIndex; changed = $changed.Count; changedProperties = @($changed) })
    return $targets
  }
  if ($sel.Type -ne 2) {
    throw "선택된 텍스트 상자나 도형이 없습니다."
  }
  for ($i = 1; $i -le $sel.ShapeRange.Count; $i++) {
    $shape = $sel.ShapeRange.Item($i)
    $changedProperties = @(Apply-FormatToShape $shape $action)
    [void]$targets.Add(@{ target = $shape.Name; slide = (Get-ActiveSlide $ppt).SlideIndex; changed = $changedProperties.Count; changedProperties = $changedProperties })
  }
  return $targets
}

function Get-SlideByFrozenSelection($pres, $frozen) {
  $slide = $null
  try {
    if ($null -ne $frozen.slideId) {
      $slide = Get-SlideById $pres ([int]$frozen.slideId)
    }
  } catch {}
  if ($null -eq $slide) {
    $idx = 0
    try { $idx = [int]$frozen.slideIndex } catch {}
    if ($idx -ge 1 -and $idx -le $pres.Slides.Count) {
      $slide = $pres.Slides.Item($idx)
    }
  }
  if ($null -eq $slide) { throw "Frozen selection slide was not found." }
  return $slide
}

function Find-FrozenShape($slide, $target) {
  $shape = $null
  try {
    $targetId = [int]$target.id
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
      $candidate = $slide.Shapes.Item($i)
      if ([int]$candidate.Id -eq $targetId) { return $candidate }
    }
  } catch {}
  try {
    $targetName = [string]$target.name
    if (![string]::IsNullOrWhiteSpace($targetName)) {
      for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $candidate = $slide.Shapes.Item($i)
        if ([string]$candidate.Name -eq $targetName) { return $candidate }
      }
    }
  } catch {}
  try {
    $targetIndex = [int]$target.shapeIndex
    if ($targetIndex -ge 1 -and $targetIndex -le $slide.Shapes.Count) {
      return $slide.Shapes.Item($targetIndex)
    }
  } catch {}
  return $shape
}

function Assert-FrozenShapeMatch($shape, $target) {
  if ($null -eq $shape) { throw "Frozen selection shape was not found." }
  $tolerance = 3.0
  try {
    if ($null -ne $target.left -and [Math]::Abs([double]$shape.Left - [double]$target.left) -gt $tolerance) { throw "left mismatch" }
    if ($null -ne $target.top -and [Math]::Abs([double]$shape.Top - [double]$target.top) -gt $tolerance) { throw "top mismatch" }
    if ($null -ne $target.width -and [Math]::Abs([double]$shape.Width - [double]$target.width) -gt $tolerance) { throw "width mismatch" }
    if ($null -ne $target.height -and [Math]::Abs([double]$shape.Height - [double]$target.height) -gt $tolerance) { throw "height mismatch" }
  } catch {
    throw ("Frozen selection shape fingerprint check failed: " + [string]$_.Exception.Message)
  }
}

function Add-ChangedProperty($changed, [string]$name) {
  if ($null -ne $changed -and ![string]::IsNullOrWhiteSpace($name)) {
    [void]$changed.Add($name)
  }
}

function Apply-ParagraphFormat($textRange, $action, $changed) {
  if ($null -eq $textRange) { return }
  $pf = $null
  try { $pf = $textRange.ParagraphFormat } catch {}
  if ($null -eq $pf) { return }

  if ($null -ne $action.lineSpacing) {
    try {
      $beforeRule = [int]$pf.LineRuleWithin
      $beforeValue = [double]$pf.SpaceWithin
      $pf.LineRuleWithin = -1
      $pf.SpaceWithin = [double]$action.lineSpacing
      if ($beforeRule -ne -1 -or [Math]::Abs($beforeValue - [double]$action.lineSpacing) -gt 0.01) {
        Add-ChangedProperty $changed "lineSpacing"
      }
    } catch {}
  }
  if ($null -ne $action.spaceWithin) {
    try {
      $beforeRule = [int]$pf.LineRuleWithin
      $beforeValue = [double]$pf.SpaceWithin
      $pf.LineRuleWithin = -1
      $pf.SpaceWithin = [double]$action.spaceWithin
      if ($beforeRule -ne -1 -or [Math]::Abs($beforeValue - [double]$action.spaceWithin) -gt 0.01) {
        Add-ChangedProperty $changed "spaceWithin"
      }
    } catch {}
  }
  if ($null -ne $action.spaceBefore) {
    try {
      $beforeRule = [int]$pf.LineRuleBefore
      $beforeValue = [double]$pf.SpaceBefore
      $pf.LineRuleBefore = -1
      $pf.SpaceBefore = [double]$action.spaceBefore
      if ($beforeRule -ne -1 -or [Math]::Abs($beforeValue - [double]$action.spaceBefore) -gt 0.01) {
        Add-ChangedProperty $changed "spaceBefore"
      }
    } catch {}
  }
  if ($null -ne $action.spaceAfter) {
    try {
      $beforeRule = [int]$pf.LineRuleAfter
      $beforeValue = [double]$pf.SpaceAfter
      $pf.LineRuleAfter = -1
      $pf.SpaceAfter = [double]$action.spaceAfter
      if ($beforeRule -ne -1 -or [Math]::Abs($beforeValue - [double]$action.spaceAfter) -gt 0.01) {
        Add-ChangedProperty $changed "spaceAfter"
      }
    } catch {}
  }
}

function Apply-FormatToShape($shape, $action) {
  $changed = New-Object System.Collections.Generic.List[string]
  if ($null -ne $action.left) { $before = [double]$shape.Left; $shape.Left = [double]$action.left; if ([Math]::Abs($before - [double]$action.left) -gt 0.05) { Add-ChangedProperty $changed "left" } }
  if ($null -ne $action.top) { $before = [double]$shape.Top; $shape.Top = [double]$action.top; if ([Math]::Abs($before - [double]$action.top) -gt 0.05) { Add-ChangedProperty $changed "top" } }
  if ($null -ne $action.width) { $before = [double]$shape.Width; $shape.Width = [double]$action.width; if ([Math]::Abs($before - [double]$action.width) -gt 0.05) { Add-ChangedProperty $changed "width" } }
  if ($null -ne $action.height) { $before = [double]$shape.Height; $shape.Height = [double]$action.height; if ([Math]::Abs($before - [double]$action.height) -gt 0.05) { Add-ChangedProperty $changed "height" } }
  if ($shape.HasTextFrame -eq -1) {
    if ($null -ne $action.fontSize) {
      $before = [double]$shape.TextFrame.TextRange.Font.Size
      $shape.TextFrame.TextRange.Font.Size = [double]$action.fontSize
      if ([Math]::Abs($before - [double]$action.fontSize) -gt 0.05) { Add-ChangedProperty $changed "fontSize" }
    }
    if ($null -ne $action.bold) {
      $before = [int]$shape.TextFrame.TextRange.Font.Bold
      $shape.TextFrame.TextRange.Font.Bold = if ([bool]$action.bold) { -1 } else { 0 }
      $afterBold = if ([bool]$action.bold) { -1 } else { 0 }
      if ($before -ne $afterBold) { Add-ChangedProperty $changed "bold" }
    }
    Apply-ParagraphFormat $shape.TextFrame.TextRange $action $changed
    if ($null -ne $action.autofit -and [bool]$action.autofit) {
      try {
        $before = [int]$shape.TextFrame.AutoSize
        $shape.TextFrame.AutoSize = 1
        if ($before -ne 1) { Add-ChangedProperty $changed "autofit" }
      } catch {}
    }
  }
  if ($null -ne $action.fillRgb) {
    try {
      $before = [int]$shape.Fill.ForeColor.RGB
      $shape.Fill.ForeColor.RGB = [int]$action.fillRgb
      if ($before -ne [int]$action.fillRgb) { Add-ChangedProperty $changed "fillRgb" }
    } catch {}
  }
  return @($changed)
}

function Format-FrozenSelection($ppt, $action, $presOverride = $null) {
  $pres = $ppt.ActivePresentation
  if ($null -ne $presOverride) { $pres = $presOverride }
  $frozen = $action.frozenSelection
  if ($null -eq $frozen -or $null -eq $frozen.shapes) {
    throw "Frozen selection target is missing."
  }
  $slide = Get-SlideByFrozenSelection $pres $frozen
  $targets = New-Object System.Collections.Generic.List[object]
  foreach ($target in @($frozen.shapes)) {
    $shape = Find-FrozenShape $slide $target
    Assert-FrozenShapeMatch $shape $target
    $changedProperties = @(Apply-FormatToShape $shape $action)
    [void]$targets.Add(@{
      target = [string]$shape.Name
      shapeId = [int]$shape.Id
      slide = [int]$slide.SlideIndex
      changed = $changedProperties.Count
      changedProperties = $changedProperties
    })
  }
  return $targets
}

function Add-TableSlide($ppt, [int]$afterIndex, [string]$title, $columns, $rows, [string]$notes, $presOverride = $null) {
  $spec = @{ kind = "content"; title = $title; bullets = @("표 데이터 요약"); notes = $notes }
  $slide = (Add-Slides $ppt $afterIndex @($spec) $presOverride)[0]
  $pres = $ppt.ActivePresentation
  if ($null -ne $presOverride) { $pres = $presOverride }
  $slideObj = $pres.Slides.Item($slide.slide)
  $colCount = @($columns).Count
  $rowCount = @($rows).Count
  if ($colCount -lt 1 -or $rowCount -lt 1) { return $slideObj }
  $tableShape = $slideObj.Shapes.AddTable($rowCount + 1, $colCount, 70, 134, 580, 260)
  for ($c = 1; $c -le $colCount; $c++) {
    $cell = $tableShape.Table.Cell(1, $c).Shape.TextFrame.TextRange
    $cell.Text = [string]@($columns)[$c - 1]
    $cell.Font.Bold = -1
    $cell.Font.Size = 12
  }
  for ($r = 1; $r -le $rowCount; $r++) {
    $row = @(@($rows)[$r - 1])
    for ($c = 1; $c -le $colCount; $c++) {
      $value = ""
      if ($row.Count -ge $c) { $value = [string]$row[$c - 1] }
      $cell = $tableShape.Table.Cell($r + 1, $c).Shape.TextFrame.TextRange
      $cell.Text = $value
      $cell.Font.Size = 11
    }
  }
  return $slideObj
}

function Add-BarChartSlide($ppt, [int]$afterIndex, [string]$title, $items, [string]$message, [string]$notes, $presOverride = $null) {
  $spec = @{ kind = "content"; title = $title; subtitle = $message; bullets = @(); notes = $notes }
  $result = (Add-Slides $ppt $afterIndex @($spec) $presOverride)[0]
  $pres = $ppt.ActivePresentation
  if ($null -ne $presOverride) { $pres = $presOverride }
  $slide = $pres.Slides.Item($result.slide)
  $list = @($items)
  if ($list.Count -eq 0) { return $slide }
  $max = 0.0
  foreach ($item in $list) {
    try {
      $v = [double]$item.value
      if ($v -gt $max) { $max = $v }
    } catch {}
  }
  if ($max -le 0) { $max = 1.0 }
  $top = 150
  $barLeft = 210
  $barMax = 420
  $rowH = 36
  for ($i = 0; $i -lt $list.Count; $i++) {
    $item = $list[$i]
    $y = $top + ($i * $rowH)
    $value = 0.0
    try { $value = [double]$item.value } catch {}
    $width = [Math]::Max(10, ($value / $max) * $barMax)
    Add-TextBox $slide ([string]$item.label) 76 $y 120 24 12 $false (Rgb 47 52 65) | Out-Null
    Add-Rect $slide $barLeft ($y + 4) $width 18 (Rgb 86 69 160) -1 0 | Out-Null
    Add-TextBox $slide ([string]$item.value) ($barLeft + $width + 8) $y 80 24 12 $false (Rgb 47 52 65) | Out-Null
  }
  return $slide
}

function Apply-ActionPlan($ppt, [string]$jsonText, $presOverride = $null) {
  $plan = $jsonText | ConvertFrom-Json
  if ($null -eq $plan.actions) { throw "Action plan has no actions." }
  $pres = $ppt.ActivePresentation
  if ($null -ne $presOverride) { $pres = $presOverride }
  if ($null -eq $presOverride) {
    try { $ppt.StartNewUndoEntry() | Out-Null } catch {}
  }
  $results = New-Object System.Collections.Generic.List[object]
  foreach ($action in @($plan.actions)) {
    $type = [string]$action.type
    $slideIndex = 0
    try { $slideIndex = [int]$action.slide } catch {}
    if ($slideIndex -lt 1 -or $slideIndex -gt $pres.Slides.Count) {
      if ($null -eq $presOverride) {
        $slideIndex = (Get-ActiveSlide $ppt).SlideIndex
      } else {
        $slideIndex = 1
      }
    }
    $slide = $pres.Slides.Item($slideIndex)

    if ($type -eq "replace_deck") {
      $created = Replace-Deck $ppt $action.slides $presOverride
      foreach ($item in $created) { [void]$results.Add($item) }
    } elseif ($type -eq "add_slides") {
      $after = $slideIndex
      try { $after = [int]$action.after } catch {}
      $created = Add-Slides $ppt $after $action.slides $presOverride
      foreach ($item in $created) { [void]$results.Add($item) }
    } elseif ($type -eq "set_title") {
      $target = Set-SlideTitle $slide ([string]$action.text)
      [void]$results.Add(@{ type = $type; slide = $slideIndex; target = $target })
    } elseif ($type -eq "set_body") {
      $target = Set-SlideBody $slide ([string]$action.text)
      [void]$results.Add(@{ type = $type; slide = $slideIndex; target = $target })
    } elseif ($type -eq "set_notes") {
      $target = Set-SpeakerNotes $slide ([string]$action.text)
      [void]$results.Add(@{ type = $type; slide = $slideIndex; target = $target })
    } elseif ($type -eq "replace_text") {
      $changed = Replace-SlideText $slide ([string]$action.find) ([string]$action.replace)
      [void]$results.Add(@{ type = $type; slide = $slideIndex; changed = $changed })
    } elseif ($type -eq "format_selection") {
      if ($null -ne $action.frozenSelection) {
        $targets = Format-FrozenSelection $ppt $action $presOverride
      } else {
      if ($null -ne $presOverride) { throw "선택 영역 서식 preview는 아직 shadow deck에서 지원하지 않습니다." }
        $targets = Format-Selection $ppt $action
      }
      foreach ($target in $targets) {
        [void]$results.Add(@{
          type = $type
          slide = $target.slide
          target = $target.target
          changed = $target.changed
          changedProperties = $target.changedProperties
        })
      }
    } elseif ($type -eq "add_table_slide") {
      $after = $slideIndex
      try { $after = [int]$action.after } catch {}
      $newSlide = Add-TableSlide $ppt $after ([string]$action.title) $action.columns $action.rows ([string]$action.notes) $presOverride
      [void]$results.Add(@{ type = $type; slide = $newSlide.SlideIndex; target = "table" })
    } elseif ($type -eq "add_bar_chart_slide") {
      $after = $slideIndex
      try { $after = [int]$action.after } catch {}
      $newSlide = Add-BarChartSlide $ppt $after ([string]$action.title) $action.items ([string]$action.message) ([string]$action.notes) $presOverride
      [void]$results.Add(@{ type = $type; slide = $newSlide.SlideIndex; target = "bar_chart_shapes" })
    } else {
      [void]$results.Add(@{ type = $type; skipped = $true; reason = "unsupported" })
    }
  }
  return @{ ok = $true; summary = [string]$plan.assistantMessage; results = $results }
}

function Save-PresentationCopy($ppt, [string]$path) {
  if ([string]::IsNullOrWhiteSpace($path)) { throw "backup path is empty." }
  $dir = [System.IO.Path]::GetDirectoryName($path)
  if (![System.IO.Directory]::Exists($dir)) {
    [System.IO.Directory]::CreateDirectory($dir) | Out-Null
  }
  $ppt.ActivePresentation.SaveCopyAs($path)
  return @{ ok = $true; path = $path }
}

function Get-AffectedSlideIndexes($pres, $affectedSlideIds, $actions) {
  $indexes = New-Object System.Collections.Generic.List[int]
  $seen = @{}
  foreach ($id in @($affectedSlideIds)) {
    try {
      $slide = Get-SlideById $pres ([int]$id)
      if ($null -ne $slide -and !$seen.ContainsKey([string]$slide.SlideIndex)) {
        [void]$indexes.Add([int]$slide.SlideIndex)
        $seen[[string]$slide.SlideIndex] = $true
      }
    } catch {}
  }
  foreach ($action in @($actions)) {
    $idx = 0
    try { $idx = [int]$action.slide } catch {}
    if ($idx -lt 1) {
      try { $idx = [int]$action.after } catch {}
    }
    if ($idx -ge 1 -and $idx -le $pres.Slides.Count -and !$seen.ContainsKey([string]$idx)) {
      [void]$indexes.Add($idx)
      $seen[[string]$idx] = $true
    }
  }
  if ($indexes.Count -eq 0) {
    try {
      $active = Get-ActiveSlide (Get-PowerPointApp)
      [void]$indexes.Add([int]$active.SlideIndex)
    } catch {
      [void]$indexes.Add(1)
    }
  }
  return $indexes
}

function Preview-ActionPlan($ppt, [string]$jsonText) {
  $stage = "parse"
  $payload = $jsonText | ConvertFrom-Json
  if ([string]::IsNullOrWhiteSpace([string]$payload.shadowPath)) { throw "shadowPath is required." }
  if ([string]::IsNullOrWhiteSpace([string]$payload.assetDir)) { throw "assetDir is required." }
  $livePres = $ppt.ActivePresentation
  $assetDir = [string]$payload.assetDir
  if (![System.IO.Directory]::Exists($assetDir)) {
    [System.IO.Directory]::CreateDirectory($assetDir) | Out-Null
  }

  $actions = @($payload.actions)
  $indexes = Get-AffectedSlideIndexes $livePres $payload.affectedSlideIds $actions
  $slides = New-Object System.Collections.Generic.List[object]
  try {
    $stage = "export-before"
    foreach ($idx in $indexes) {
      if ($idx -ge 1 -and $idx -le $livePres.Slides.Count) {
        $slide = $livePres.Slides.Item($idx)
        $before = Join-Path $assetDir ("before-{0}.png" -f $slide.SlideID)
        Export-SlideImage $slide $before | Out-Null
        [void]$slides.Add(@{
          slideId = $slide.SlideID
          slideIndex = $slide.SlideIndex
          beforeImage = $before
          afterImage = $null
          textDiff = @()
        })
      }
    }

    $stage = "save-shadow"
    $shadowPath = [string]$payload.shadowPath
    $shadowDir = [System.IO.Path]::GetDirectoryName($shadowPath)
    if (![System.IO.Directory]::Exists($shadowDir)) {
      [System.IO.Directory]::CreateDirectory($shadowDir) | Out-Null
    }
    $livePres.SaveCopyAs($shadowPath)

    $stage = "open-shadow"
    $shadow = $null
    $openMode = "hidden"
    try {
      $shadow = $ppt.Presentations.Open($shadowPath, 0, 0, 0)
    } catch {
      $openMode = "visible_fallback"
      $shadow = $ppt.Presentations.Open($shadowPath)
    }
    $stage = "apply-shadow"
    $shadowSlideCountBefore = $shadow.Slides.Count
    $plan = @{ assistantMessage = [string]$payload.assistantMessage; actions = $actions } | ConvertTo-Json -Compress -Depth 12
    $result = Apply-ActionPlan $ppt $plan $shadow

    $stage = "export-after-existing"
    for ($slideRecordIndex = 0; $slideRecordIndex -lt $slides.Count; $slideRecordIndex++) {
      $item = $slides.Item($slideRecordIndex)
      $shadowSlide = $null
      $itemSlideId = $item["slideId"]
      $itemSlideIndex = $item["slideIndex"]
      if ($null -ne $itemSlideId) {
        $stage = "export-after-existing-find-id"
        $shadowSlide = Get-SlideById $shadow ([int]$itemSlideId)
      }
      if ($null -eq $shadowSlide -and $itemSlideIndex -ge 1 -and $itemSlideIndex -le $shadow.Slides.Count) {
        $stage = "export-after-existing-find-index"
        $shadowSlide = $shadow.Slides.Item([int]$itemSlideIndex)
      }
      if ($null -ne $shadowSlide) {
        $stage = "export-after-existing-path"
        $after = Join-Path $assetDir ("after-{0}.png" -f $itemSlideId)
        $stage = "export-after-existing-export"
        Export-SlideImage $shadowSlide $after | Out-Null
        $stage = "export-after-existing-record"
        $item["afterImage"] = $after
      }
    }

    $stage = "export-after-new"
    for ($idx = ($shadowSlideCountBefore + 1); $idx -le $shadow.Slides.Count; $idx++) {
      $stage = "export-after-new-get-slide"
      $slide = $shadow.Slides.Item($idx)
      $stage = "export-after-new-path"
      $after = Join-Path $assetDir ("after-new-{0}.png" -f $idx)
      $stage = "export-after-new-export"
      Export-SlideImage $slide $after | Out-Null
      $stage = "export-after-new-record"
      [void]$slides.Add(@{
        slideId = $slide.SlideID
        slideIndex = $slide.SlideIndex
        beforeImage = $null
        afterImage = $after
        textDiff = @()
      })
    }

    return @{ ok = $true; shadowDeckPath = $shadowPath; openMode = $openMode; slides = $slides; result = $result }
  } catch {
    throw ("preview stage " + $stage + ": " + [string]$_.Exception.Message)
  } finally {
    if ($null -ne $shadow) {
      try { $shadow.Close() } catch {}
    }
  }
}

function Open-PresentationCopy($ppt, [string]$path) {
  if ([string]::IsNullOrWhiteSpace($path) -or !(Test-Path -LiteralPath $path)) {
    throw "열 backup 파일을 찾지 못했습니다."
  }
  $opened = $ppt.Presentations.Open($path, 0, 0, -1)
  return @{ ok = $true; name = [string]$opened.Name; path = [string]$opened.FullName }
}

function Save-ActivePresentation($ppt) {
  $pres = $ppt.ActivePresentation
  $pres.Save()
  return @{ ok = $true; name = [string]$pres.Name; path = [string]$pres.FullName }
}

try {
  $ppt = Get-PowerPointApp
  if ($ppt.Presentations.Count -lt 1) {
    throw "열린 PowerPoint 프레젠테이션이 없습니다."
  }

  if ($Action -eq "context") {
    Write-Json (Get-Context $ppt)
    exit 0
  }

  if ($Action -eq "deck-text") {
    Write-Json @{ ok = $true; text = Get-DeckText $ppt }
    exit 0
  }

  if ($Action -eq "selection-context") {
    $slide = Get-ActiveSlide $ppt
    Write-Json @{
      ok = $true
      presentationName = [string]$ppt.ActivePresentation.Name
      presentationFullName = [string]$ppt.ActivePresentation.FullName
      slideCount = $ppt.ActivePresentation.Slides.Count
      slideIndex = $slide.SlideIndex
      slideId = $slide.SlideID
      selection = Get-SelectionContext $ppt
    }
    exit 0
  }

  $inputText = ""
  if ($InputPath -and (Test-Path -LiteralPath $InputPath)) {
    $inputText = Get-Content -Raw -Encoding UTF8 -LiteralPath $InputPath
  }

  if ($Action -eq "apply-json") {
    if ([string]::IsNullOrWhiteSpace($inputText)) { throw "실행할 계획 JSON이 비어 있습니다." }
    Write-Json (Apply-ActionPlan $ppt $inputText)
    exit 0
  }

  if ($Action -eq "save-copy") {
    if ([string]::IsNullOrWhiteSpace($inputText)) { throw "backup 요청 JSON이 비어 있습니다." }
    $payload = $inputText | ConvertFrom-Json
    Write-Json (Save-PresentationCopy $ppt ([string]$payload.backupPath))
    exit 0
  }

  if ($Action -eq "preview-json") {
    if ([string]::IsNullOrWhiteSpace($inputText)) { throw "preview 요청 JSON이 비어 있습니다." }
    Write-Json (Preview-ActionPlan $ppt $inputText)
    exit 0
  }

  if ($Action -eq "shape-map") {
    Write-Json (Get-ShapeMap $ppt $inputText)
    exit 0
  }

  if ($Action -eq "save-active") {
    Write-Json (Save-ActivePresentation $ppt)
    exit 0
  }

  if ($Action -eq "open-presentation") {
    if ([string]::IsNullOrWhiteSpace($inputText)) { throw "open 요청 JSON이 비어 있습니다." }
    $payload = $inputText | ConvertFrom-Json
    Write-Json (Open-PresentationCopy $ppt ([string]$payload.path))
    exit 0
  }

  throw "알 수 없는 작업입니다: $Action"
} catch {
  Write-Json @{ ok = $false; error = [string]$_.Exception.Message }
  exit 1
}
