# ============================================================
# RJFolderRenamer.ps1
# DLSITEからタイトルを取得し、フォルダ名のRJ番号をタイトルに置換する
# ============================================================

# Invoke-WebRequest の警告を抑制する（PowerShell 5.x 向け）
$PSDefaultParameterValues['Invoke-WebRequest:UseBasicParsing'] = $true

# コンソールの入出力エンコーディングを UTF-8 に統一する
# PowerShell 5.x はデフォルトが Shift-JIS のため、日本語パスの取り扱いや
# 画面出力が文字化けする。chcp と OutputEncoding の両方を設定することで解消する。
$OutputEncoding = [Text.Encoding]::UTF8
[Console]::OutputEncoding = [Text.Encoding]::UTF8
[Console]::InputEncoding  = [Text.Encoding]::UTF8
chcp 65001 | Out-Null

# System.Web を事前にロード（HTMLデコードに使用）
# ループ内で毎回呼ぶと警告が出るため、ここで一度だけ読み込む
Add-Type -AssemblyName System.Web

# ============================================================
# フォルダ統合ヘルパー関数
# 移動先に同名フォルダが既に存在する場合、中身を上書き統合する
# ============================================================
function Merge-Folders {
    param (
        [string]$Source,      # 移動元フォルダのフルパス
        [string]$Destination  # 移動先フォルダのフルパス
    )

    # 移動元フォルダ内のすべてのアイテムを取得（再帰的に）
    $items = Get-ChildItem -LiteralPath $Source -Recurse -Force

    foreach ($item in $items) {
        # 移動元からの相対パスを計算して移動先のパスを構築する
        $relativePath = $item.FullName.Substring($Source.Length).TrimStart('\')
        $destPath = Join-Path $Destination $relativePath

        if ($item.PSIsContainer) {
            # サブフォルダが存在しない場合は作成する
            if (-not (Test-Path -LiteralPath $destPath)) {
                New-Item -ItemType Directory -Path $destPath | Out-Null
            }
        } else {
            # ファイルを移動先へコピー（同名ファイルは上書き）
            Copy-Item -LiteralPath $item.FullName -Destination $destPath -Force
        }
    }

    # 移動元フォルダを削除する（統合完了後）
    Remove-Item -LiteralPath $Source -Recurse -Force
}

# ============================================================
# メイン処理
# ============================================================

# 作業ディレクトリをスクリプト実行時のカレントフォルダに設定する
$baseDir = Get-Location

# カレントフォルダ直下のサブフォルダのうち、名前に "RJ数字" を含むものだけを抽出する
$folders = Get-ChildItem -Path $baseDir -Directory | Where-Object { $_.Name -match 'RJ\d+' }

# 対象フォルダが見つからない場合は早期終了する
if ($folders.Count -eq 0) {
    Write-Host "対象フォルダが見つかりませんでした。"
    exit
}

Write-Host "対象フォルダ数: $($folders.Count)"
Write-Host ("-" * 50)

# 対象フォルダごとに処理を繰り返す
foreach ($folder in $folders) {

    # フォルダ名から RJ番号（例: RJ123456）を正規表現で取り出す
    if ($folder.Name -match '(RJ\d+)') {
        $rj = $matches[1]

        # DLsite maniax の作品ページURLを組み立てる（RJ番号は maniax カテゴリ固定）
        $url = "https://www.dlsite.com/maniax/work/=/product_id/$rj"

        try {
            # DLsite ページを取得する
            $response = Invoke-WebRequest -Uri $url -Headers @{ "User-Agent" = "Mozilla/5.0" } -ErrorAction Stop
            $html = $response.Content
            # HTML内から <h1 id="work_name">...</h1> を抽出する
            # 子タグが含まれる場合に備えて子タグを除去してからテキストを取り出す
            if ($html -match '<h1[^>]*id="work_name"[^>]*>([\s\S]*?)</h1>') {
                $raw   = $matches[1]
                # 子タグ（<span> など）を除去してプレーンテキストにする
                $decoded = ($raw -replace '<[^>]+>', '').Trim()
                # HTMLエンティティ（&#039; → ' など）をデコードする
                $title = [System.Web.HttpUtility]::HtmlDecode($decoded).Trim()
            } else {
                Write-Host "⚠️ タイトルが取得できません: $rj"
                continue
            }

            # Windowsのファイル名に使えない文字（\/:*?"<>|）を _ に置換する
            $safeTitle = ($title -replace '[\\/:*?"<>|]', '_')

            # フォルダ名の RJ番号部分をタイトルに置換して新しいフォルダ名を作る
            $newName = $folder.Name -replace [Regex]::Escape($rj), $safeTitle

            # 変更がない場合はスキップする（既にリネーム済みのフォルダを誤って処理しない）
            if ($newName -eq $folder.Name) {
                Write-Host "⏭️ スキップ（変更なし）: $($folder.Name)"
                continue
            }

            # リネーム先のフルパスを組み立てる
            $newFullPath = Join-Path $baseDir $newName

            if (Test-Path $newFullPath) {
                # ─────────────────────────────────────────────
                # 同名フォルダが既に存在する場合 → 統合（マージ）する
                # 移動元の中身を移動先へコピーし、移動元を削除する
                # ─────────────────────────────────────────────
                Write-Host "🔀 統合中（同名フォルダあり）: $($folder.Name) → $newName"
                Merge-Folders -Source $folder.FullName -Destination $newFullPath
                Write-Host "✅ 統合完了: $($folder.Name) → $newName"
            } else {
                # 同名フォルダがない場合は通常のリネームを行う
                # LiteralPath を使うことで [] や日本語などの特殊文字を含むパスを正確に扱う
                Rename-Item -LiteralPath $folder.FullName -NewName $newName -ErrorAction Stop
                Write-Host "✅ リネーム完了: $($folder.Name) → $newName"
            }
        }
        catch {
            # 予期しないエラーが発生した場合はエラーメッセージを表示して次へ進む
            Write-Host "⚠️ 失敗: $rj → $($_.Exception.Message)"
        }

        # DLsiteサーバーへの負荷を避けるため、1秒待機してから次のフォルダを処理する
        Start-Sleep -Seconds 1
    }
}

Write-Host ("-" * 50)
Write-Host "処理完了"