# =============================================================================
# sync_vba_to_books.ps1
# =============================================================================
# src/vba/*.bas と src/ribbon/customUI14.xml を、
#   1. src/src.xlsm（編集元ブック）
#   2. releases/在庫自動調整.xlam（配布用アドイン）
# の両方へ反映し、コンパイル確認・smoke テスト・取込結果の diff 検証まで行う。
#
# 使い方（リポジトリのどこからでも可。Windows PowerShell 5.1 / pwsh 7 両対応）:
#   pwsh -File tools/sync_vba_to_books.ps1
#   pwsh -File tools/sync_vba_to_books.ps1 -SkipRibbon          # モジュールのみ反映
#   pwsh -File tools/sync_vba_to_books.ps1 -CompileTimeoutSec 180
#
# 前提:
#   - Excel（デスクトップ版）がインストールされていること
#   - Excel の「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が有効なこと
#     （docs/開発ガイド.md 2.1 参照）
#   - EXCEL.EXE が起動していないこと（起動中なら中止する）
#
# 注意:
#   - このファイルは UTF-8 (BOM 付き) で保存すること。BOM が無いと
#     Windows PowerShell 5.1 が日本語リテラルを誤読して動かなくなる。
#   - コンパイルエラーがあると VBE がモーダルダイアログで固まるため、
#     ウォッチドッグ（別プロセス）が CompileTimeoutSec 秒後に EXCEL を強制終了する。
#   - 反映前に両ブックを %TEMP%\koten-macro-backup\ へバックアップする。
# =============================================================================

[CmdletBinding()]
param(
    # リポジトリのルート。省略時はこのスクリプトの親フォルダの1つ上を使う。
    [string]$RepoRoot = '',
    # リボン XML の zip 差し替えを行わない場合に指定。
    [switch]$SkipRibbon,
    # コンパイル確認のウォッチドッグ時間（秒）。
    [int]$CompileTimeoutSec = 90
)

$ErrorActionPreference = 'Stop'

# PowerShell 5.1 では param 既定値の段階で $PSScriptRoot が使えないため、本体で解決する。
if ([string]::IsNullOrEmpty($RepoRoot)) {
    $scriptDir = $PSScriptRoot
    if ([string]::IsNullOrEmpty($scriptDir)) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
    $RepoRoot = Split-Path -Parent $scriptDir
}

$vbaDir     = Join-Path $RepoRoot 'src\vba'
$ribbonPath = Join-Path $RepoRoot 'src\ribbon\customUI14.xml'
$books      = @(
    (Join-Path $RepoRoot 'src\src.xlsm'),
    (Join-Path $RepoRoot 'releases\在庫自動調整.xlam')
)

function Fail([string]$message) {
    Write-Host "NG: $message" -ForegroundColor Red
    exit 1
}

# -----------------------------------------------------------------------------
# 0. 事前チェック
# -----------------------------------------------------------------------------
if (-not (Test-Path $vbaDir))     { Fail "src/vba が見つかりません: $vbaDir" }
if (-not (Test-Path $ribbonPath)) { Fail "リボン定義が見つかりません: $ribbonPath" }
foreach ($book in $books) {
    if (-not (Test-Path $book)) { Fail "対象ブックが見つかりません: $book" }
}

$basFiles = @(Get-ChildItem -Path $vbaDir -Filter '*.bas' | Sort-Object Name)
if ($basFiles.Count -eq 0) { Fail "src/vba に .bas がありません。" }
$moduleNames = @($basFiles | ForEach-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Name) })

if (Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue) {
    Fail "EXCEL.EXE が起動中です。すべての Excel を閉じてから再実行してください。"
}

# VBE 自動化の許可（AccessVBOM）を確認する。Office のバージョンキーは環境で変わるため走査する。
$accessVbomOk = $false
foreach ($officeKey in (Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue)) {
    $securityKey = Join-Path $officeKey.PSPath 'Excel\Security'
    if (Test-Path $securityKey) {
        $value = (Get-ItemProperty -Path $securityKey -Name 'AccessVBOM' -ErrorAction SilentlyContinue).AccessVBOM
        if ($value -eq 1) { $accessVbomOk = $true }
    }
}
if (-not $accessVbomOk) {
    Fail "「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が無効です。docs/開発ガイド.md 2.1 を参照してください。"
}

# -----------------------------------------------------------------------------
# 1. バックアップ
# -----------------------------------------------------------------------------
$backupDir = Join-Path $env:TEMP ("koten-macro-backup\" + (Get-Date -Format 'yyyyMMdd_HHmmss'))
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
foreach ($book in $books) {
    Copy-Item -Path $book -Destination $backupDir
}
Write-Host "バックアップ: $backupDir"

# -----------------------------------------------------------------------------
# 2. リボン XML の zip 差し替え（Excel で開く前に行う）
# -----------------------------------------------------------------------------
Add-Type -AssemblyName System.IO.Compression.FileSystem

$ribbonEntryName = 'customUI/customUI14.xml'
$ribbonBytes = [System.IO.File]::ReadAllBytes($ribbonPath)

if (-not $SkipRibbon) {
    foreach ($book in $books) {
        $zip = [System.IO.Compression.ZipFile]::Open($book, [System.IO.Compression.ZipArchiveMode]::Update)
        try {
            $entry = $zip.GetEntry($ribbonEntryName)
            if ($null -ne $entry) { $entry.Delete() }
            $entry = $zip.CreateEntry($ribbonEntryName)
            $stream = $entry.Open()
            try { $stream.Write($ribbonBytes, 0, $ribbonBytes.Length) } finally { $stream.Dispose() }
        }
        finally {
            $zip.Dispose()
        }
        Write-Host "リボン反映: $(Split-Path -Leaf $book)"
    }
} else {
    Write-Host "リボン反映: スキップ（-SkipRibbon）"
}

# -----------------------------------------------------------------------------
# 3. Excel COM でモジュール入替・コンパイル・smoke・検証
# -----------------------------------------------------------------------------
$sentinelDir = Join-Path $env:TEMP 'koten-macro-sync'
New-Item -ItemType Directory -Force -Path $sentinelDir | Out-Null
$exportDir = Join-Path $sentinelDir ("export_" + (Get-Date -Format 'yyyyMMdd_HHmmss'))
New-Item -ItemType Directory -Force -Path $exportDir | Out-Null

$excel = $null
$failures = @()

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    # Workbook_Open などの自動マクロを走らせない（AutomationSecurity は使わない。
    # ForceDisable にするとブック内マクロの実行自体が禁止され、smoke テストの
    # Application.Run まで失敗するため）。
    $excel.EnableEvents = $false

    foreach ($book in $books) {
        $bookName = Split-Path -Leaf $book
        Write-Host "--- $bookName ---"

        $wb = $excel.Workbooks.Open($book, 0, $false)
        try {
            $vbproj = $wb.VBProject

            # 3-1. 標準モジュールを src/vba と同期する（vbext_ct_StdModule = 1）。
            #      同名モジュールの入替に加え、src/vba に無い mod* の残留モジュールも削除する。
            $existingComponents = @()
            foreach ($comp in $vbproj.VBComponents) { $existingComponents += $comp }
            foreach ($comp in $existingComponents) {
                if ($comp.Type -eq 1 -and $comp.Name -like 'mod*') {
                    if ($moduleNames -notcontains $comp.Name) {
                        Write-Host "残留モジュール削除: $($comp.Name)"
                    }
                    $vbproj.VBComponents.Remove($comp)
                }
            }
            foreach ($bas in $basFiles) {
                [void]$vbproj.VBComponents.Import($bas.FullName)
            }
            Write-Host "モジュール取込: $($basFiles.Count) 本"

            # 3-2. プロジェクト全体をコンパイル（VBE CommandBar コントロール ID 578）
            #      コンパイルエラー時はモーダルで固まるため、ウォッチドッグを先に仕掛ける。
            $sentinel = Join-Path $sentinelDir ("compile_done_" + [Guid]::NewGuid().ToString('N'))
            $watchdog = Start-Process -PassThru -WindowStyle Hidden powershell -ArgumentList @(
                '-NoProfile', '-Command',
                "for (`$i = 0; `$i -lt $CompileTimeoutSec; `$i++) { if (Test-Path '$sentinel') { exit 0 }; Start-Sleep 1 }; Stop-Process -Name EXCEL -Force -ErrorAction SilentlyContinue"
            )
            try {
                $excel.VBE.ActiveVBProject = $vbproj
                $compileControl = $excel.VBE.CommandBars.FindControl(1, 578)  # msoControlButton, コンパイル
                if ($null -ne $compileControl -and $compileControl.Enabled) {
                    $compileControl.Execute()
                }
                New-Item -ItemType File -Path $sentinel -Force | Out-Null
                Write-Host "コンパイル: OK"
            }
            catch {
                New-Item -ItemType File -Path $sentinel -Force | Out-Null
                throw "コンパイル確認に失敗しました（コンパイルエラーの可能性。VBE でデバッグ→コンパイルを実行して確認してください）: $_"
            }
            finally {
                if (-not $watchdog.HasExited) { Stop-Process -Id $watchdog.Id -Force -ErrorAction SilentlyContinue }
            }

            # 3-3. smoke テスト。DebugCompileSmokeTest は Option Private Module のため、
            #      同一プロジェクト内に一時モジュールを作って呼び出す。
            $smokeModule = $vbproj.VBComponents.Add(1)
            try {
                $smokeModule.CodeModule.AddFromString("Public Function ZZSyncSmokeCall() As Boolean`r`nZZSyncSmokeCall = DebugCompileSmokeTest()`r`nEnd Function")
                $smokeResult = $excel.Run("'$($wb.Name)'!ZZSyncSmokeCall")
                if ($smokeResult -ne $true) { throw "DebugCompileSmokeTest が True を返しませんでした: $smokeResult" }
                Write-Host "smoke テスト: OK"
            }
            finally {
                $vbproj.VBComponents.Remove($smokeModule)
            }

            $wb.Save()

            # 3-4. 取込結果をエクスポートして、src/vba と byte 比較で検証する
            $bookExportDir = Join-Path $exportDir ($bookName -replace '\.', '_')
            New-Item -ItemType Directory -Force -Path $bookExportDir | Out-Null
            foreach ($bas in $basFiles) {
                $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($bas.Name)
                $exportPath = Join-Path $bookExportDir $bas.Name
                $vbproj.VBComponents.Item($moduleName).Export($exportPath)
                $srcBytes = [System.IO.File]::ReadAllBytes($bas.FullName)
                $expBytes = [System.IO.File]::ReadAllBytes($exportPath)
                if (-not [System.Linq.Enumerable]::SequenceEqual([byte[]]$srcBytes, [byte[]]$expBytes)) {
                    # VBA は大文字小文字を区別しないため、大文字小文字だけの差は動作に影響しない。
                    # ただし放置すると diff が読みにくくなるので、原因ヒント付きで警告する
                    # （docs/開発ガイド.md トラブルシューティングの「識別子リケース」参照）。
                    $srcText = [System.Text.Encoding]::GetEncoding(932).GetString($srcBytes)
                    $expText = [System.Text.Encoding]::GetEncoding(932).GetString($expBytes)
                    if ([string]::Equals($srcText, $expText, [System.StringComparison]::OrdinalIgnoreCase)) {
                        Write-Host "警告: $($bas.Name) は大文字小文字だけが src/vba と異なります（VBE の識別子リケース）。" -ForegroundColor Yellow
                    } else {
                        $failures += "${bookName} / $($bas.Name): エクスポート結果が src/vba と一致しません（$exportPath）"
                    }
                }
            }

            # 取込後の mod* 標準モジュール一覧が src/vba と完全一致するかも確認する
            $bookModules = @()
            foreach ($comp in $vbproj.VBComponents) {
                if ($comp.Type -eq 1 -and $comp.Name -like 'mod*') { $bookModules += $comp.Name }
            }
            $unexpected = @($bookModules | Where-Object { $moduleNames -notcontains $_ })
            $missing    = @($moduleNames | Where-Object { $bookModules -notcontains $_ })
            if ($unexpected.Count -gt 0) { $failures += "${bookName}: src/vba に無いモジュールが残っています: $($unexpected -join ', ')" }
            if ($missing.Count -gt 0)    { $failures += "${bookName}: 取込漏れモジュール: $($missing -join ', ')" }
            Write-Host "モジュール diff 検証: 完了"
        }
        finally {
            # コンパイルエラーでウォッチドッグが EXCEL を強制終了した場合、ここでの Close は
            # RPC エラーになる。本来の診断メッセージ（throw 済み）を握りつぶさないよう例外は無視する。
            try { $wb.Close($false) } catch { }
        }

        # 3-5. リボンが保存後も維持されているか byte 比較で検証する
        $zip = [System.IO.Compression.ZipFile]::OpenRead($book)
        try {
            $entry = $zip.GetEntry($ribbonEntryName)
            if ($null -eq $entry) {
                $failures += "${bookName}: zip 内に $ribbonEntryName がありません"
            } else {
                $ms = New-Object System.IO.MemoryStream
                $stream = $entry.Open()
                try { $stream.CopyTo($ms) } finally { $stream.Dispose() }
                if (-not [System.Linq.Enumerable]::SequenceEqual([byte[]]$ribbonBytes, [byte[]]$ms.ToArray())) {
                    $failures += "${bookName}: zip 内のリボンが src/ribbon/customUI14.xml と一致しません"
                }
            }
        }
        finally {
            $zip.Dispose()
        }
        Write-Host "リボン検証: 完了"
    }
}
finally {
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    # COM の後始末に失敗して EXCEL が残った場合の保険。
    Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

# -----------------------------------------------------------------------------
# 4. 結果
# -----------------------------------------------------------------------------
if ($failures.Count -gt 0) {
    Write-Host ''
    $failures | ForEach-Object { Write-Host "NG: $_" -ForegroundColor Red }
    exit 1
}

Write-Host ''
Write-Host "PASS: すべてのモジュールとリボンが両ブックへ反映され、コンパイル・smoke・diff 検証を通過しました。" -ForegroundColor Green
Write-Host "バックアップ: $backupDir"
exit 0
