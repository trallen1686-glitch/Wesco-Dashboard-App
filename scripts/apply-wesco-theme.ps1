[CmdletBinding()]
param(
    [string]$Root = (Join-Path $PSScriptRoot '..'),
    [switch]$Preview,
    [switch]$Remove
)

$ErrorActionPreference = 'Stop'
$themeStart = '<!-- WESCO-DARK-THEME:START -->'
$themeEnd = '<!-- WESCO-DARK-THEME:END -->'

$themeBlock = @'
<!-- WESCO-DARK-THEME:START -->
<style>
@import url('https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600;700&display=swap');

:root {
  --wesco-orange: #F36A00;
  --wesco-orange-light: #FF9A4D;
  --wesco-page: #080D10;
  --wesco-panel: rgba(9, 14, 16, 0.74);
  --wesco-panel-solid: rgba(5, 10, 12, 0.88);
  --wesco-border: rgba(255, 255, 255, 0.14);
  --wesco-text: #FFFFFF;
  --wesco-text-soft: rgba(255, 255, 255, 0.82);
  --wesco-font: 'Open Sans', sans-serif;
  --page-bg: var(--wesco-page) !important;
  --card-bg: var(--wesco-panel) !important;
  --border: var(--wesco-border) !important;
  --ink: var(--wesco-text) !important;
  --ink-soft: var(--wesco-text-soft) !important;
  --blue: var(--wesco-orange) !important;
  --blue-bright: var(--wesco-orange) !important;
  --maroon: var(--wesco-orange) !important;
}

*, *::before, *::after { box-sizing: border-box; }

html, body {
  min-height: 100%;
  margin: 0;
  color: var(--wesco-text) !important;
  font-family: var(--wesco-font) !important;
  background-color: var(--wesco-page) !important;
  background-image:
    linear-gradient(115deg, rgba(6,10,12,.96) 0%, rgba(9,14,16,.82) 42%, rgba(9,14,16,.94) 100%),
    radial-gradient(circle at 72% 18%, rgba(243,106,0,.10), transparent 32%),
    repeating-linear-gradient(90deg, transparent 0 79px, rgba(255,255,255,.018) 80px),
    repeating-linear-gradient(0deg, transparent 0 79px, rgba(255,255,255,.014) 80px) !important;
  background-attachment: fixed !important;
}

body, button, input, textarea, select, option, table {
  font-family: var(--wesco-font) !important;
}

body > main, body > .widget, body > .container, body > .wrapper,
.page-shell, .dashboard-shell {
  width: 100% !important;
  max-width: 1280px !important;
  margin: 0 auto !important;
  padding: 18px !important;
}

.card, .panel, .widget-card, .dashboard-card, .person-card,
.tab-tile, .tile, .quick-tool-chip, .modal-content, .dropdown-menu,
.field, .form-section, .attach-item, .bid-card, .list-card, .stat-card {
  color: var(--wesco-text) !important;
  background: var(--wesco-panel) !important;
  border-color: var(--wesco-border) !important;
  border-radius: 4px !important;
  box-shadow: 0 12px 30px rgba(0,0,0,.30) !important;
  backdrop-filter: blur(8px);
  -webkit-backdrop-filter: blur(8px);
}

header, .header, .page-header, .dashboard-header,
.starter-header, .messages-header {
  color: var(--wesco-text) !important;
  background: rgba(4,8,10,.62) !important;
  border-color: var(--wesco-border) !important;
  border-bottom: 2px solid var(--wesco-orange) !important;
}

h1, h2, h3, h4, h5, h6, .title, .person-name,
.starter-name, .tab-label, .card-title {
  color: var(--wesco-text) !important;
}

p, label, small, .subtitle, .helper-text, .muted, .contact-preview {
  color: var(--wesco-text-soft) !important;
}

a, .back-link, .section-heading, .role-section-title, .section-label {
  color: var(--wesco-orange) !important;
}

.back-link {
  background: transparent !important;
  border: none !important;
  box-shadow: none !important;
}

.card b, .tile b, .tool-card b {
  color: var(--wesco-orange-light) !important;
}

.tab-desc, .tile-description, .card-description, .page-description {
  color: var(--wesco-orange-light) !important;
}

button, .btn, input[type='button'], input[type='submit'] {
  color: #FFFFFF !important;
  background: var(--wesco-orange) !important;
  border: 1px solid var(--wesco-orange) !important;
  border-radius: 3px !important;
}

button:hover, .btn:hover, input[type='button']:hover, input[type='submit']:hover {
  background: #FF7A14 !important;
}

input:not([type='button']):not([type='submit']), textarea, select {
  color: var(--wesco-text) !important;
  background: rgba(3,7,9,.78) !important;
  border: 1px solid var(--wesco-border) !important;
}

input::placeholder, textarea::placeholder {
  color: rgba(255,255,255,.48) !important;
}

table, th, td {
  color: var(--wesco-text) !important;
  border-color: var(--wesco-border) !important;
}

th {
  color: var(--wesco-orange-light) !important;
  background: rgba(4,8,10,.66) !important;
}

tr:hover td { background: rgba(243,106,0,.08) !important; }

/* Keep initials and job descriptions, but remove their small boxes. */
.avatar, .initials, .staff-initials {
  color: var(--wesco-orange-light) !important;
  background: transparent !important;
  border: none !important;
  box-shadow: none !important;
}

.tag, .role-tag, .job-title, .job-description {
  color: var(--wesco-orange-light) !important;
  background: transparent !important;
  border: none !important;
  box-shadow: none !important;
}

/* Preserve all Wesco logo colors. */
img.logo, .logo img, .brand img, .branding img {
  filter: none !important;
  opacity: 1 !important;
  mix-blend-mode: normal !important;
}

@media (max-width: 600px) {
  body > main, body > .widget, body > .container, body > .wrapper,
  .page-shell, .dashboard-shell { padding: 8px !important; }
}
</style>
<!-- WESCO-DARK-THEME:END -->
'@

$resolvedRoot = (Resolve-Path -LiteralPath $Root).Path
$ignoredFolders = '[\\/](\.git|\.github|node_modules|dist|build|coverage|vendor)[\\/]'
$markerPattern = '(?s)' + [regex]::Escape($themeStart) + '.*?' + [regex]::Escape($themeEnd) + "`r?`n?"
$markerRegex = [regex]::new($markerPattern)
$headRegex = [regex]::new('</head>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
$byteSafeEncoding = [System.Text.Encoding]::GetEncoding(28591)

$files = Get-ChildItem -LiteralPath $resolvedRoot -Recurse -File -Filter '*.html' |
    Where-Object { $_.FullName -notmatch $ignoredFolders }

$changed = 0
$skipped = 0

foreach ($file in $files) {
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    $content = $byteSafeEncoding.GetString($bytes)
    $newContent = $content
    # Use LF inside the managed block so Git diffs stay consistent across platforms.
    $fileThemeBlock = [regex]::Replace($themeBlock, "`r`n|`r|`n", "`n")

    if ($Remove) {
        if ($markerRegex.IsMatch($content)) {
            $newContent = $markerRegex.Replace($content, '', 1)
        }
    }
    elseif ($markerRegex.IsMatch($content)) {
        $newContent = $markerRegex.Replace($content, $fileThemeBlock + "`n", 1)
    }
    elseif ($headRegex.IsMatch($content)) {
        $newContent = $headRegex.Replace($content, $fileThemeBlock + "`n</head>", 1)
    }
    else {
        Write-Warning "Skipped because </head> was not found: $($file.FullName)"
        $skipped++
        continue
    }

    if ($newContent -ceq $content) {
        continue
    }

    $changed++
    if ($Preview) {
        Write-Host "Would update: $($file.FullName)" -ForegroundColor Yellow
    }
    else {
        [System.IO.File]::WriteAllBytes($file.FullName, $byteSafeEncoding.GetBytes($newContent))
        Write-Host "Updated: $($file.FullName)" -ForegroundColor Green
    }
}

if ($Preview) {
    Write-Host "Preview: $changed file(s) would change; $skipped skipped."
}
else {
    Write-Host "Complete: $changed file(s) changed; $skipped skipped."
}
