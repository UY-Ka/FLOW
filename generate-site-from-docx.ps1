$ErrorActionPreference = 'Stop'

$docPath = 'c:\Users\podun\Yandex.Disk\Обучение пользователей информационных систем\Методическое пособие.docx'
$outDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$mediaDir = Join-Path $outDir 'assets\docx-media'

New-Item -ItemType Directory -Force -Path $mediaDir | Out-Null

Add-Type -AssemblyName System.IO.Compression
$fs = [System.IO.File]::Open($docPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
$zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Read)

try {
  $readXml = {
    param($entryName)
    $entry = $zip.GetEntry($entryName)
    $reader = [System.IO.StreamReader]::new($entry.Open())
    try { return [xml]$reader.ReadToEnd() }
    finally { $reader.Close() }
  }

  $documentXml = & $readXml 'word/document.xml'
  $relsXml = & $readXml 'word/_rels/document.xml.rels'

  $rels = @{}
  foreach ($rel in $relsXml.Relationships.Relationship) {
    if ($rel.Type -like '*image') {
      $rels[$rel.Id] = $rel.Target
    }
  }

  foreach ($target in $rels.Values) {
    $entryName = 'word/' + $target.TrimStart('/')
    $entry = $zip.GetEntry($entryName)
    if ($entry) {
      $dest = Join-Path $mediaDir ([System.IO.Path]::GetFileName($target))
      $inStream = $entry.Open()
      $outStream = [System.IO.File]::Create($dest)
      try { $inStream.CopyTo($outStream) }
      finally {
        $outStream.Close()
        $inStream.Close()
      }
    }
  }

  $ns = [System.Xml.XmlNamespaceManager]::new($documentXml.NameTable)
  $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
  $ns.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

  function Get-ParagraphItem($p) {
    $text = (($p.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join '').Trim()
    $images = @()
    foreach ($attr in $p.SelectNodes('.//@r:embed', $ns)) {
      if ($rels.ContainsKey($attr.Value)) {
        $images += ('assets/docx-media/' + [System.IO.Path]::GetFileName($rels[$attr.Value]))
      }
    }
    if ($text -or $images.Count -gt 0) {
      return [pscustomobject]@{ Type = 'paragraph'; Text = $text; Images = $images }
    }
    return $null
  }

  $paragraphs = @()
  $body = $documentXml.SelectSingleNode('//w:body', $ns)
  foreach ($node in $body.ChildNodes) {
    if ($node.LocalName -eq 'p') {
      $item = Get-ParagraphItem $node
      if ($item) { $paragraphs += $item }
    }
    elseif ($node.LocalName -eq 'tbl') {
      $rows = @()
      foreach ($tr in $node.SelectNodes('./w:tr', $ns)) {
        $cells = @()
        foreach ($tc in $tr.SelectNodes('./w:tc', $ns)) {
          $cellParts = @()
          foreach ($p in $tc.SelectNodes('.//w:p', $ns)) {
            $cellText = (($p.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join '').Trim()
            if ($cellText) { $cellParts += $cellText }
          }
          $cells += ($cellParts -join "`n")
        }
        if ($cells.Count -gt 0) { $rows += ,$cells }
      }
      if ($rows.Count -gt 0) {
        $paragraphs += [pscustomobject]@{ Type = 'table'; Text = ''; Images = @(); Rows = $rows }
      }
    }
  }
}
finally {
  $zip.Dispose()
  $fs.Close()
}

function Enc($value) {
  return [System.Net.WebUtility]::HtmlEncode([string]$value)
}

function Nav($active) {
  $links = @(
    @{ Key = 'home'; Href = 'index.html'; Text = 'Главная'; Icon = '⌂' },
    @{ Key = 'lectures'; Href = 'lectures.html'; Text = 'Лекции'; Icon = 'Л' },
    @{ Key = 'labs'; Href = 'labs.html'; Text = 'Практические'; Icon = 'П' },
    @{ Key = 'questions'; Href = 'questions.html'; Text = 'Вопросы'; Icon = '?' }
  )
  $html = @"
    <aside class="sidebar">
      <div class="brand">course.</div>
      <div class="menu-title">Навигация</div>
"@
  foreach ($link in $links) {
    $class = if ($link.Key -eq $active) { 'menu-link active' } else { 'menu-link' }
    $html += "`n      <a class=""$class"" href=""$($link.Href)""><span class=""icon"">$($link.Icon)</span>$($link.Text)</a>"
  }
  $html += "`n    </aside>"
  return $html
}

function HtmlShell($title, $body) {
  return @"
<!doctype html>
<html lang="ru">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>$(Enc $title)</title>
  <link rel="stylesheet" href="styles.css">
</head>
<body>
$body
</body>
</html>
"@
}

function ContentHtml($items, $skipFirstTitle) {
  $html = ''
  $skipped = -not $skipFirstTitle
  $pendingImages = @()

  foreach ($item in $items) {
    if ($item.Type -eq 'table') {
      $html += "`n        <div class=""table-wrap"">`n          <table class=""doc-table"">"
      for ($rowIndex = 0; $rowIndex -lt $item.Rows.Count; $rowIndex++) {
        $tag = if ($rowIndex -eq 0) { 'th' } else { 'td' }
        $html += "`n            <tr>"
        foreach ($cell in $item.Rows[$rowIndex]) {
          $encoded = (Enc $cell) -replace "(`r`n|`n|`r)", '<br>'
          $html += "<$tag>$encoded</$tag>"
        }
        $html += "</tr>"
      }
      $html += "`n          </table>`n        </div>"
      continue
    }

    foreach ($image in $item.Images) {
      $pendingImages += $image
    }

    $text = $item.Text
    if (-not $text) { continue }
    if (-not $skipped) {
      $skipped = $true
      continue
    }

    if ($text -match '^Рисунок\s+\d+') {
      $html += "`n        <figure class=""doc-figure"">"
      if ($pendingImages.Count -gt 0) {
        $src = Enc $pendingImages[0]
        $alt = Enc $text
        $html += "`n          <img src=""$src"" alt=""$alt"">"
        if ($pendingImages.Count -gt 1) { $pendingImages = $pendingImages[1..($pendingImages.Count - 1)] } else { $pendingImages = @() }
      }
      $html += "`n          <figcaption>$(Enc $text)</figcaption>`n        </figure>"
    }
    elseif ($text -match '^\d+\.\s') {
      $html += "`n        <h3>$(Enc $text)</h3>"
    }
    elseif ($text -match '^(Теоретическая часть|Практическая часть|Принципы обработки данных|Основные понятия баз данных|Архитектура баз данных)$') {
      $html += "`n        <h3>$(Enc $text)</h3>"
    }
    elseif ($text -match '^Цель работы:') {
      $html += "`n        <p class=""goal"">$(Enc $text)</p>"
    }
    elseif ($text -match '^\+\+\+') {
      $clean = $text -replace '^\+\+\+\s*', ''
      $html += "`n        <p class=""correct-answer"">$(Enc $clean)</p>"
    }
    elseif ($text -match '\+\+\+') {
      $clean = $text -replace '\s*\+\+\+\s*', ' '
      $html += "`n        <p class=""correct-answer"">$(Enc $clean)</p>"
    }
    else {
      $html += "`n        <p>$(Enc $text)</p>"
    }
  }

  foreach ($image in $pendingImages) {
    $src = Enc $image
    $html += "`n        <figure class=""doc-figure""><img src=""$src"" alt=""Иллюстрация из методического пособия""></figure>"
  }

  return $html
}

$starts = @()
for ($i = 0; $i -lt $paragraphs.Count; $i++) {
  $text = $paragraphs[$i].Text
  if ($text -match '^(Лекция\s+[1-7]\..+|Практическое задание\s+[1-8]\..+|Тестовые вопросы для самоконтроля)$' -and $text -notmatch '\d$') {
    $starts += [pscustomobject]@{ Index = $i; Title = $text }
  }
}

$sections = @()
for ($i = 0; $i -lt $starts.Count; $i++) {
  $from = $starts[$i].Index
  $to = if ($i -lt $starts.Count - 1) { $starts[$i + 1].Index - 1 } else { $paragraphs.Count - 1 }
  $sections += [pscustomobject]@{
    Title = $starts[$i].Title
    Items = $paragraphs[$from..$to]
  }
}

$lectures = $sections | Where-Object { $_.Title -match '^Лекция' }
$labs = $sections | Where-Object { $_.Title -match '^Практическое задание' }
$questions = $sections | Where-Object { $_.Title -match '^Тестовые вопросы' } | Select-Object -First 1

$annotation = 'Курс посвящён проектированию баз данных: от теоретических основ организации данных и архитектуры СУБД до ER-моделирования, нормализации, SQL-запросов и разработки пользовательских форм в Microsoft Access.'

$lectureList = ''
for ($i = 0; $i -lt $lectures.Count; $i++) {
  $num = $i + 1
  $name = ($lectures[$i].Title -replace '^Лекция\s+\d+\.\s*', '')
  $lectureList += "`n          <div class=""list-item""><span class=""index-badge"">$num</span><a class=""list-link"" href=""lecture-$num.html"">$(Enc $name)</a></div>"
}

$labList = ''
for ($i = 0; $i -lt $labs.Count; $i++) {
  $num = $i + 1
  $labList += "`n          <div class=""list-item""><span class=""index-badge"">$num</span><a class=""list-link"" href=""lab-$num.html"">$(Enc $($labs[$i].Title))</a></div>"
}

$homeBody = @"
  <div class="dashboard">
$(Nav 'home')
    <main class="main-panel">
      <div class="topbar">
        <h1>Проектирование баз данных</h1>
      </div>

      <section class="annotation">
        <strong>Аннотация:</strong>
        $(Enc $annotation)
      </section>

      <section>
        <h2>Лекции</h2>
        <div class="list">$lectureList
        </div>
        <div class="inline-nav">
          <a class="chip-btn secondary" href="lectures.html">Открыть все лекции</a>
        </div>
      </section>

      <section>
        <h2>Практические задания</h2>
        <div class="list">$labList
        </div>
        <div class="inline-nav">
          <a class="chip-btn secondary" href="labs.html">Открыть все практические</a>
        </div>
      </section>
    </main>
  </div>
"@
Set-Content -LiteralPath (Join-Path $outDir 'index.html') -Encoding utf8 -Value (HtmlShell 'Курс по базам данных' $homeBody)

$lecturesBody = @"
  <div class="dashboard">
$(Nav 'lectures')
    <main class="main-panel">
      <div class="topbar">
        <h1>Все лекции</h1>
      </div>
      <section>
        <div class="list">$lectureList
        </div>
      </section>
    </main>
  </div>
"@
Set-Content -LiteralPath (Join-Path $outDir 'lectures.html') -Encoding utf8 -Value (HtmlShell 'Лекции' $lecturesBody)

$labsBody = @"
  <div class="dashboard">
$(Nav 'labs')
    <main class="main-panel">
      <div class="topbar">
        <h1>Все практические задания</h1>
      </div>
      <section>
        <div class="list">$labList
        </div>
      </section>
    </main>
  </div>
"@
Set-Content -LiteralPath (Join-Path $outDir 'labs.html') -Encoding utf8 -Value (HtmlShell 'Практические задания' $labsBody)

for ($i = 0; $i -lt $lectures.Count; $i++) {
  $num = $i + 1
  $title = $lectures[$i].Title
  $name = $title -replace '^Лекция\s+\d+\.\s*', ''
  $nextLab = if ($num -le $labs.Count) { "`n        <a class=""button-link"" href=""lab-$num.html"">Перейти к практическому заданию $num</a>" } else { '' }
  $content = ContentHtml $lectures[$i].Items $true
  $body = @"
  <main class="page">
    <h1>Лекция $num</h1>
    <h2>$(Enc $name)</h2>
    <section class="content-card">$content
      <div class="button-row">$nextLab
        <a class="button-link" href="questions.html">Вопросы для самоконтроля</a>
      </div>
    </section>

    <a class="back-link" href="lectures.html">Назад к лекциям</a>
  </main>
"@
  Set-Content -LiteralPath (Join-Path $outDir "lecture-$num.html") -Encoding utf8 -Value (HtmlShell $title $body)
}

for ($i = 0; $i -lt $labs.Count; $i++) {
  $num = $i + 1
  $title = $labs[$i].Title
  $name = $title -replace '^Практическое задание\s+\d+\.\s*', ''
  $content = ContentHtml $labs[$i].Items $true
  $lectureBack = if ($num -le $lectures.Count) { "lecture-$num.html" } else { 'labs.html' }
  $backText = if ($num -le $lectures.Count) { "Назад к лекции $num" } else { 'Назад к практическим' }
  $body = @"
  <main class="page">
    <h1>Практическое задание $num</h1>
    <h2>$(Enc $name)</h2>
    <section class="content-card">$content
      <div class="button-row">
        <a class="button-link" href="questions.html">Вопросы для самоконтроля</a>
      </div>
    </section>

    <a class="back-link" href="$lectureBack">$backText</a>
  </main>
"@
  Set-Content -LiteralPath (Join-Path $outDir "lab-$num.html") -Encoding utf8 -Value (HtmlShell $title $body)
}

if ($questions) {
  $content = ContentHtml $questions.Items $true
  $body = @"
  <main class="page">
    <h1>Вопросы для самоконтроля</h1>
    <section class="content-card question-block">$content
    </section>

    <a class="back-link" href="index.html">На главную</a>
  </main>
"@
  Set-Content -LiteralPath (Join-Path $outDir 'questions.html') -Encoding utf8 -Value (HtmlShell 'Вопросы для самоконтроля' $body)
}

Write-Host "Generated $($lectures.Count) lectures, $($labs.Count) labs and media files in assets/docx-media."


