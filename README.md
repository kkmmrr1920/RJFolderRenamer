# RJFolderRenamer

RJ番号（例: `RJ123456`）で管理されているフォルダを、DLsite から作品タイトルを取得して自動的にリネームする PowerShell スクリプトです。  
Windows 環境での禁止文字や重複対策も行っています。


## 特徴
- RJ番号（例: `RJ123456`）を検出
- DLsite から作品タイトルを自動取得
- Windows で使えない文字を `_` に置換

## 使い方

### 1. スクリプトをダウンロード
- このリポジトリの [Releases](../../releases) ページから `RJFolderRenamer.ps1` をダウンロードしてください。

### 2. フォルダへ配置
- RJ番号（例: `RJ123456`）が含まれる対象フォルダと **同じ場所** に `RJFolderRenamer.ps1` を置きます。

例:
📂 作業フォルダ
┣ 📂 RJ123456
┣ 📂 RJ654321
┗ 📄 script.ps1

### 3. PowerShell で実行
PowerShellを起動し
対象フォルダ（`RJFolderRenamer.ps1` がある場所）に移動して実行します。

例：PS F:\Downloads> ./RJFolderRenamer.ps1
