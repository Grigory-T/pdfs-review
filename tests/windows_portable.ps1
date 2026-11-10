param([Parameter(Mandatory=$true)][string]$Zip,
      [Parameter(Mandatory=$true)][string]$Fixtures,
      [string]$Results)
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
$utf8=New-Object Text.UTF8Encoding($false)
$base=Join-Path (Split-Path ([IO.Path]::GetFullPath($Zip)) -Parent) ('acceptance-'+[Guid]::NewGuid().ToString('N'))
$root=Join-Path $base ('tool '+[char]0x00e9)
$inputDir=Join-Path $base ('input '+[char]0x044f+' 100% ! # & (data)')
$outside=Join-Path $base 'outside-canary'
New-Item -ItemType Directory -Force -Path $root,$inputDir,$outside|Out-Null
Expand-Archive -LiteralPath $Zip -DestinationPath $root
Expand-Archive -LiteralPath $Fixtures -DestinationPath $inputDir
$inputBefore=@{}
Get-ChildItem -LiteralPath $inputDir -File | ForEach-Object {
    $inputBefore[$_.Name]=(Get-FileHash -LiteralPath $_.FullName).Hash+':'+$_.LastWriteTimeUtc.Ticks
}
foreach($name in @('TEMP','TMP','APPDATA','LOCALAPPDATA','PYTHONPYCACHEPREFIX')) {
    [Environment]::SetEnvironmentVariable($name,$outside,'Process')
}
# Prove that installed Python/uv/7-Zip and inherited uv indexes are not required.
$env:PATH=$env:SystemRoot+'\System32;'+$env:SystemRoot+';'+$env:SystemRoot+'\System32\WindowsPowerShell\v1.0'
$env:UV_INDEX='https://invalid.invalid/simple'
$env:UV_PYTHON_INSTALL_DIR=$outside
$records=New-Object Collections.Generic.List[object]

function Set-Setting([string]$Name,[string]$Value,[switch]$Unicode) {
    $path=Join-Path $root 'pdfs-review.txt'
    $text=[IO.File]::ReadAllText($path)
    $evaluator=[Text.RegularExpressions.MatchEvaluator]{param($m) $Name+' = '+$Value}
    $text=[regex]::Replace($text,('(?m)^'+[regex]::Escape($Name)+'\s*=.*$'),$evaluator)
    $encoding=if($Unicode){[Text.Encoding]::Unicode}else{$utf8}
    [IO.File]::WriteAllText($path,$text,$encoding)
}
function Run-Tool([string]$Case,[int]$Expected,[switch]$Alias) {
    $old=$ErrorActionPreference
    try {
        $ErrorActionPreference='Continue'
        $entry=if($Alias){'RUN.bat'}else{'pdfs-review.bat'}
        $lines=& cmd.exe /d /c ('"'+(Join-Path $root $entry)+'" --no-pause') 2>&1
        $code=$LASTEXITCODE
    } finally { $ErrorActionPreference=$old }
    $script:output=($lines|Out-String)
    if($code -ne $Expected){throw "$Case expected $Expected, got $code`n$output"}
    if(@(Get-ChildItem -LiteralPath $outside -Force).Count){throw 'Wrote to outside profile/temp canary'}
    $temp=Join-Path $root 'support/state/temp'
    if(Test-Path -LiteralPath $temp){
        if(@(Get-ChildItem -LiteralPath $temp -Filter 'work-*').Count){throw 'Temporary scan work leaked'}
        if(@(Get-ChildItem -LiteralPath $temp -Filter 'openpyxl.*').Count){throw 'Worksheet XML leaked'}
    }
    $records.Add([ordered]@{case=$Case;exit_code=$code})
    Write-Host ($Case+': PASS')
}
function Latest-Summary {
    $path=Get-ChildItem -LiteralPath $root -Filter 'pdfs-review-*-summary.json'|Sort-Object LastWriteTimeUtc -Descending|Select-Object -First 1
    if(-not $path){throw 'Summary missing'}
    return (Get-Content -LiteralPath $path.FullName -Raw -Encoding UTF8|ConvertFrom-Json)
}

Set-Setting 'folder' ('"'+$inputDir+'"') -Unicode
Run-Tool 'cold-setup-and-real-bat' 0
if($output -notmatch 'Setup: COMPLETE'){throw 'Cold setup did not run'}
$stats=Latest-Summary
if($stats.pdfs_found -ne 3 -or $stats.pages -ne 4 -or $stats.errors -ne 0 -or $stats.archives -ne 3){throw 'Mixed archive/PDF counts differ'}
if($stats.classes.digital -ne 4){throw 'Classification changed'}
if(Test-Path (Join-Path $root 'support/state/setup.jsonl')){throw 'Default setup created a journal'}
$firstReports=@(Get-ChildItem -LiteralPath $root -Filter '*.xlsx')
if($firstReports.Count -ne 1){throw 'Expected report beside BAT'}
$python=Join-Path $root 'support/state/venv/Scripts/python.exe'
$pins=Get-Content -LiteralPath (Join-Path $root 'support/pins.json') -Raw|ConvertFrom-Json
$archive=Join-Path (Join-Path (Join-Path $root 'support/state/downloads/python-mirror') $pins.python.release) $pins.python.archive
if((Get-FileHash -LiteralPath $archive).Hash.ToLowerInvariant() -ne $pins.python.sha256){throw 'Python archive not verified'}

$uv=Join-Path $root 'support/state/uv.exe'
Move-Item -LiteralPath $uv -Destination ($uv+'.saved')
$env:HTTP_PROXY='http://127.0.0.1:9'; $env:HTTPS_PROXY=$env:HTTP_PROXY; $env:ALL_PROXY=$env:HTTP_PROXY; $env:NO_PROXY=''
try {
    Run-Tool 'offline-restart-without-uv' 0
    if($output -notmatch 'Setup: SKIPPED' -or $output -match 'Download:|First-time setup'){throw 'Offline restart performed setup'}
} finally { Move-Item -LiteralPath ($uv+'.saved') -Destination $uv }
if(@(Get-ChildItem -LiteralPath $root -Filter '*.xlsx').Count -ne 2){throw 'Prior timestamped report not retained'}
Set-Setting 'text_threshold' 'invalid'
Run-Tool 'invalid-setting-before-scan' 1
if(@(Get-ChildItem -LiteralPath $root -Filter '*.xlsx').Count -ne 2){throw 'Invalid config created report'}
Set-Setting 'text_threshold' '40'
Run-Tool 'compatibility-bat-same-runtime' 0 -Alias
$lock=[IO.File]::Open((Join-Path $root 'support/state/workspace.lock'),'OpenOrCreate','ReadWrite','None')
try { Run-Tool 'concurrent-workspace-rejected' 10 } finally { $lock.Dispose() }
$extra=Join-Path $root 'support/pymupdf.py'
[IO.File]::WriteAllText($extra,'raise SystemExit(77)',$utf8)
try { Run-Tool 'unlisted-import-shadow-rejected' 10 } finally { Remove-Item -LiteralPath $extra }
$technical=Join-Path $root 'support/pdfs_review.py'
$bytes=[IO.File]::ReadAllBytes($technical)
[IO.File]::AppendAllText($technical,"`n# changed",$utf8)
try { Run-Tool 'modified-source-rejected' 10 } finally { [IO.File]::WriteAllBytes($technical,$bytes) }
Set-Setting 'folder' (Join-Path $root 'support')
Run-Tool 'support-input-rejected' 1
Set-Setting 'folder' $root
Run-Tool 'self-scan-excludes-support' 0
if((Latest-Summary).pdfs_found -ne 0 -or (Latest-Summary).archives -ne 0){throw 'Support was scanned'}
$broken=Join-Path $base 'broken-input'
New-Item -ItemType Directory -Path $broken|Out-Null
[IO.File]::WriteAllText((Join-Path $broken 'broken.pdf'),'Synthetic invalid PDF',$utf8)
Set-Setting 'folder' $broken
Run-Tool 'unreadable-pdf-produces-usable-partial' 2
Set-Setting 'folder' $inputDir
Set-Setting 'diagnostics' 'yes'
Run-Tool 'opt-in-neutral-setup-journal' 0
$journal=[IO.File]::ReadAllText((Join-Path $root 'support/state/setup.jsonl'))
foreach($private in @($root,$inputDir,$env:USERNAME,'loose.PDF')) {
    if($private -and $journal.IndexOf($private,[StringComparison]::OrdinalIgnoreCase) -ge 0){throw 'Setup journal contains private data'}
}
Set-Setting 'diagnostics' 'no'
$oldRoot=$root
$root=Join-Path $base 'relocated tool'
Copy-Item -LiteralPath $oldRoot -Destination $root -Recurse
Run-Tool 'offline-relocation-from-retained-cache' 0
if($output -notmatch 'Folder moved:'){throw 'Relocation was not detected'}
foreach($file in Get-ChildItem -LiteralPath $inputDir -File) {
    if($inputBefore[$file.Name] -ne ((Get-FileHash -LiteralPath $file.FullName).Hash+':'+$file.LastWriteTimeUtc.Ticks)){throw 'Input contents/mtime changed'}
}
$receipt=[ordered]@{all_checks_passed=$true;zip_sha256=(Get-FileHash -LiteralPath $Zip).Hash.ToLowerInvariant();cases=$records.ToArray();input_preserved=$true;outside_canary_untouched=$true}
if($Results){[IO.File]::WriteAllText($Results,($receipt|ConvertTo-Json -Depth 5),$utf8)}
Write-Host 'WINDOWS PORTABLE ACCEPTANCE: SUCCESS'
