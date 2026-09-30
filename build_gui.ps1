# Build PPT-Rendering-Deck GUI into a single-folder exe.
# Pure ASCII on purpose: no encoding pitfalls for PS 5.1 / 7.
param(
  [string]$Root = $PSScriptRoot
)
$ErrorActionPreference = 'Stop'
$py      = (Get-Command python -ErrorAction SilentlyContinue).Source
if (-not $py) { throw 'Python not found on PATH (install Python 3.10+ and add it to PATH)' }
$pyi     = Join-Path (Split-Path $py) 'Scripts\pyinstaller.exe'
$proj    = $Root
$assets  = Join-Path $proj 'assets'
$entry   = Join-Path $proj 'ppt_gui.py'
$ico     = Join-Path $proj 'pptdeck.ico'
$verfile = Join-Path $proj 'version_info.txt'

# 1) regenerate PNG icons from the .ico (Tk 8.6 reads PNG natively -> no Pillow at runtime)
New-Item -ItemType Directory -Force -Path $assets | Out-Null
$gen = @'
import sys
from PIL import Image
src, out = sys.argv[1], sys.argv[2]
im = Image.open(src)
for name, size in (('icon32.png', (32, 32)), ('icon256.png', (256, 256))):
    try:
        frame = im.ico.getimage(size)
    except Exception:
        frame = im.convert('RGBA').resize(size)
    frame.convert('RGBA').save(out + '\\' + name)
    print('wrote', name, frame.size)
'@
$genPath = Join-Path $proj '_gen_icons.py'
Set-Content -LiteralPath $genPath -Value $gen -Encoding UTF8
& $py -X utf8 $genPath $ico $assets
if ($LASTEXITCODE -ne 0) { throw 'icon generation failed' }

# 1b) engine dependency (pptxgenjs) must be present for the packaged app to render
if (-not (Test-Path (Join-Path $proj 'engine\node_modules\pptxgenjs'))) {
  Write-Host 'installing engine dependency (pptxgenjs)...'
  Push-Location (Join-Path $proj 'engine')
  npm install pptxgenjs --no-audit --no-fund
  Pop-Location
}
# 2) build onedir (fast startup, per doc 07) windowed GUI exe
& $pyi --noconfirm --clean --windowed --log-level WARN `
  --name 'PPT渲染台' `
  --icon $ico `
  --version-file $verfile `
  --add-data "$assets;assets" `
  --collect-all tkinterdnd2 `
  --distpath (Join-Path $proj 'dist') `
  --workpath (Join-Path $proj 'build') `
  --specpath $proj `
  $entry
if ($LASTEXITCODE -ne 0) { throw 'pyinstaller failed' }

# 3) engine goes next to the exe (the app prefers app_dir()/engine), docs go with it
$distDir = Join-Path $proj 'dist\PPT渲染台'
robocopy (Join-Path $proj 'engine') (Join-Path $distDir 'engine') /MIR /NFL /NDL /NJH /NJS /NP | Out-Null
Get-ChildItem -LiteralPath $proj -Filter '*.md' -File | Copy-Item -Destination $distDir -Force

$exe = Join-Path $proj 'dist\PPT渲染台\PPT渲染台.exe'
Write-Host ''
Write-Host ('BUILD OK -> ' + $exe)
Write-Host ('size: ' + [math]::Round((Get-Item -LiteralPath $exe).Length / 1MB, 2) + ' MB')
