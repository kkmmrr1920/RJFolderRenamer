# 出力を UTF-8 に設定（日本語文字化け防止）
[Console]::OutputEncoding = [Text.Encoding]::UTF8

# 作業ディレクトリを取得（例：現在のフォルダ）
$baseDir = Get-Location

# フォルダ一覧を取得し、名前に RJ数字 が含まれるフォルダだけを抽出
$folders = Get-ChildItem -Path $baseDir -Directory | Where-Object { $_.Name -match 'RJ\d+' }

# 対象フォルダごとに処理
foreach ($folder in $folders) {
    # フォルダ名から RJ番号を取り出す
    if ($folder.Name -match '(RJ\d+)') {
        $rj = $matches[1]

        # DLsite の作品ページ URL を作成
        $url = "https://www.dlsite.com/maniax/work/=/product_id/$rj"

        try {
            # DLsite ページを取得
            $response = Invoke-WebRequest -Uri $url -Headers @{ "User-Agent" = "Mozilla/5.0" }
            $html = $response.Content

            # HTML 内から <h1 id="work_name">作品名</h1> を抜き出す
            if ($html -match '<h1[^>]*id="work_name"[^>]*>(.*?)</h1>') {
                # タイトルを HTML デコード（&#039; → ' など）して整形
            } else {
                Write-Host "タイトルが取得できません: $rj"
                continue
            }
            
            # Windows で使えない文字（\/:*?"<>|）を _ に置換
            $safeTitle = ($title -replace '[\\/:*?"<>|]', '_')

            # フォルダ名の RJ番号部分を作品タイトルに置換
            $newName = $folder.Name -replace [Regex]::Escape($rj), $safeTitle

            # フォルダ名を変更
            Rename-Item -Path $folder.FullName -NewName $newName -ErrorAction Stop
            Write-Host "✅ $($folder.Name) → $newName"
        }
        catch {
            # エラー発生時の処理
            Write-Host "⚠️ 失敗: $rj $($_.Exception.Message)"
        }
    }
}
