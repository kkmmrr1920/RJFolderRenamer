# ============================================================
# RJFolderRenamer.ps1
# DLsiteからタイトルを取得し、フォルダ名のRJ番号をタイトルに置換する
# ============================================================

[CmdletBinding()]
param(
    # 処理対象のベースディレクトリ。未指定ならカレントを使う（既存挙動）
    [string]$BaseDir = (Get-Location).Path,
    # DLsiteへの連続アクセスを避けるための待機秒数
    [int]$DelaySeconds = 1,
    # HTTPタイムアウト（応答が無い場合に処理全体が止まるのを防ぐ）
    [int]$TimeoutSec = 30
)

# PowerShell 5.x で Invoke-WebRequest が IE エンジン未起動警告を出すのを抑止
$PSDefaultParameterValues['Invoke-WebRequest:UseBasicParsing'] = $true

# PowerShell 5.x の既定エンコーディング(Shift-JIS)では日本語フォルダ名やコンソール出力が
# 文字化けするため UTF-8 に統一する
$OutputEncoding = [Text.Encoding]::UTF8
[Console]::OutputEncoding = [Text.Encoding]::UTF8
[Console]::InputEncoding  = [Text.Encoding]::UTF8
chcp 65001 | Out-Null

# HtmlDecode 用。ループ内で都度ロードすると警告が出るので最初に一度だけ読む
Add-Type -AssemblyName System.Web


# ------------------------------------------------------------
# 移動先に同名フォルダが既に存在する場合、中身をコピーして統合する
# コピー失敗時に元フォルダを消さないよう try/catch で保護する
# ------------------------------------------------------------
function Merge-Folders {
    param(
        [string]$Source,
        [string]$Destination
    )

    try {
        $items = Get-ChildItem -LiteralPath $Source -Recurse -Force -ErrorAction Stop
        foreach ($item in $items) {
            # Source 起点の相対パスを Destination に貼り直す
            $relative = $item.FullName.Substring($Source.Length).TrimStart('\')
            $destPath = Join-Path $Destination $relative

            if ($item.PSIsContainer) {
                if (-not (Test-Path -LiteralPath $destPath)) {
                    New-Item -ItemType Directory -Path $destPath -ErrorAction Stop | Out-Null
                }
            } else {
                # 同名ファイルは上書き（既存挙動維持）
                Copy-Item -LiteralPath $item.FullName -Destination $destPath -Force -ErrorAction Stop
            }
        }
        # 全コピー成功後にのみ元フォルダを削除する
        Remove-Item -LiteralPath $Source -Recurse -Force -ErrorAction Stop
    }
    catch {
        # 途中失敗時は元フォルダを残す（データ損失防止）
        throw "統合失敗: $($_.Exception.Message)"
    }
}

# ------------------------------------------------------------
# DLsite maniax の作品ページからタイトルを取得する
# 取得不可なら $null を返す（呼び出し側でスキップ判定）
# ------------------------------------------------------------
function Get-DlsiteTitle {
    param([string]$RJ)

    $url = "https://www.dlsite.com/maniax/work/=/product_id/$RJ"
    $response = Invoke-WebRequest -Uri $url `
        -Headers @{ "User-Agent" = "Mozilla/5.0" } `
        -TimeoutSec $TimeoutSec `
        -ErrorAction Stop

    # <h1 id="work_name"> の中身を抜き出す。子タグ(<span>等)が混ざる場合があるので
    # まずタグを剥がしてから HtmlDecode する
    if ($response.Content -match '<h1[^>]*id="work_name"[^>]*>([\s\S]*?)</h1>') {
        $plain = ($matches[1] -replace '<[^>]+>', '').Trim()
        return [System.Web.HttpUtility]::HtmlDecode($plain).Trim()
    }
    return $null
}


# ============================================================
# メイン処理
# ============================================================

# 名前に RJ数字 を含むサブフォルダだけを対象にする
$folders = Get-ChildItem -LiteralPath $BaseDir -Directory |
    Where-Object { $_.Name -match 'RJ\d+' }

if ($folders.Count -eq 0) {
    Write-Host "対象フォルダが見つかりませんでした。"
    return
}

Write-Host "対象フォルダ数: $($folders.Count)"
Write-Host ("-" * 50)

foreach ($folder in $folders) {
    if ($folder.Name -notmatch '(RJ\d+)') { continue }
    $rj = $matches[1]

    try {
        $title = Get-DlsiteTitle -RJ $rj
        if (-not $title) {
            Write-Host "⚠️ タイトルが取得できません: $rj"
            continue
        }

        # Windows のファイル名禁止文字を _ に置換
        $safeTitle = $title -replace '[\\/:*?"<>|]', '_'
        $newName   = $folder.Name -replace [Regex]::Escape($rj), $safeTitle

        # 既にリネーム済みのフォルダを再処理しないためのガード
        if ($newName -eq $folder.Name) {
            Write-Host "⏭️ スキップ（変更なし）: $($folder.Name)"
            continue
        }

        $newFullPath = Join-Path $BaseDir $newName

        if (Test-Path -LiteralPath $newFullPath) {
            # 同名フォルダ衝突時は中身を統合してから元を削除
            Write-Host "🔀 統合中（同名フォルダあり）: $($folder.Name) → $newName"
            Merge-Folders -Source $folder.FullName -Destination $newFullPath
            Write-Host "✅ 統合完了: $($folder.Name) → $newName"
        } else {
            # LiteralPath: [] や日本語等の特殊文字を含むパスを正確に扱う
            Rename-Item -LiteralPath $folder.FullName -NewName $newName -ErrorAction Stop
            Write-Host "✅ リネーム完了: $($folder.Name) → $newName"
        }
    }
    catch {
        Write-Host "⚠️ 失敗: $rj → $($_.Exception.Message)"
    }

    # DLsite への連続アクセスによる負荷/ブロックを避けるため待機
    Start-Sleep -Seconds $DelaySeconds
}

Write-Host ("-" * 50)
Write-Host "処理完了"
