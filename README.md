# RJFolderRenamer

RJ番号で管理されているフォルダを、DLsite から作品タイトルを取得して自動的にリネームする PowerShell スクリプトです。  
Windows 環境での日本語ファイル名や禁止文字への対応も行います。

## 特徴
- RJ番号（例: `RJ123456`）を検出
- DLsite から作品タイトルを自動取得
- Windows で使えない文字を `_` に置換
