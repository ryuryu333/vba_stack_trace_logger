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



Singleton管理の自動化/隠蔽:
ResetMyLogger の手動呼び出しをなくし、MyLogger.Initialize の呼び出し時や、アドインのロード時 (例: Auto_Open や Workbook_Open イベントなど、状況に応じて) に自動的に状態をリセットする、あるいは常にクリーンなインスタンスを返すように設計変更を検討します。
依存性逆転の積極的活用:
循環参照のある箇所では、抽象（インターフェースやイベント）に依存するように変更します。例えば、Logger_Controller が IStackTraceNotificationReceiver のようなインターフェースに依存し、Logger_StackTraceController がそれを実装してイベントを通知する形などが考えられます。
設定フローの再考:
Logger_Facade.Initialize() が Logger_ConfigBuilder を返し、Logger_ConfigBuilder.Build() が Logger_ConfigStruct を返却。その構造体を Logger_Facade の別のメソッド（例: ApplyConfig(config As Logger_ConfigStruct)）に渡すか、あるいは Facade 自身が Build を呼び出した ConfigBuilder を引数に取るメソッドを持つなど、Facadeが主体的に設定を完了するフローを検討します。