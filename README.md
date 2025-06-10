# vba_logger
VBA 用のロガー、イミディエイトウィンドウ、エクセルシート出力に対応

## 使用方法

```bash
# コードのエクスポート
cscript vbac.wsf decombine
# コードのインポート
cscript vbac.wsf combine
```

## コードファイル作成時の注意点
・**改行コード**
CRLF
・**文字エンコード**
UTF-8

## ライセンス

このプロジェクトはMITライセンスの下で公開されています。

### 第三者ライブラリ

このプロジェクトは以下の第三者ライブラリを使用しています：

#### vbac.wsf (Ariawase Library)
- **ライセンス**: MIT License
- **著作権**: Copyright (c) 2011 igeta
- **プロジェクトページ**: https://github.com/vbaidiot/ariawase
- **変更内容**: @Folderアノテーションサポート機能を追加（フォルダ階層の自動構築）

完全なライセンステキストは、vbac.wsfファイル内に記載されています。