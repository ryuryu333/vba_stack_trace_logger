# VBA Stack Trace Logger

VBAでスタックトレース付きのログ出力を簡単に実現できるアドインです。

```:Sample
[2025-06-21 00:02:44.459][Trace][MyModule.SubProc < MyModule.MainProc] >> Enter MyModule.SubProc
[2025-06-21 00:02:44.461][DEBUG][MyModule.SubProc < MyModule.MainProc] Hi!
[2025-06-21 00:02:44.462][Trace][MyModule.SubProc < MyModule.MainProc] << Exit MyModule.SubProc
```

## WIP 本ロガーは開発途中です WIP
main ブランチは正常に動く状態にしています。
ただし、現在はプレリリース状態であり、以下の作業が進行中です。

- ドキュメント整備（英語/日本語対応、APIリファレンス、クラス図などのUML）
- 単体テスト構築
  - 上記に伴うクラス設計・コードの見直し
- エラーハンドリングの徹底（現状は最低限です）
- その他 機能改修（優先度低）

もしもバグを発見した場合は Issue にて連絡ください。

## 概要・利点
- 関数の呼び出し履歴（スタックトレース）を自動で記録
  - プロシージャ開始・終了ログを自動出力
  - 通常のログ出力に呼び出し記録を併記
- ログレベル（Info, Warning, Errorなど）のタグ付けが可能
- イミディエイトウィンドウやExcelシートへの出力に対応
- Debug.Print と同じ使用感で利用可能

## インストール方法
1. [binフォルダ](vba_stack_trace_logger/bin)から`VbaStackTraceLogger.xlam`をダウンロード
2. 任意のフォルダ（例: `C:\Users\YourName\Documents\vba_addins\`）に保存
3. Excelで「開発」タブ→「Visual Basic」→「ツール」→「参照設定」→「参照(B)」から`VbaStackTraceLogger.xlam`を追加

## クイックスタート
1. 標準モジュールを挿入し、以下のサンプルコードを貼り付けてください。

```vba:MyModule
Option Explicit
Private Const MODULE_NAME As String = "MyModule"

Sub CheckLogger()
    ' === Initialize ===
    myLogger.StartConfiguration _
        .EnableStackTrace _
        .EnableWriteToExcelSheet _
        .SetOutputExcelSheet(ActiveSheet) _
        .Build
    ' === Use ===
    myLogger.Log "Start"
    MainProc
    myLogger.Log "End"
    ' === Terminate ===
    myLogger.Terminate
End Sub

Sub MainProc()
    Const PROC_NAME As String = "MainProc": Dim scopeGuard As Variant: Set scopeGuard = myLogger.UsingTracer(MODULE_NAME, PROC_NAME)
    
    myLogger.Log "Hi!", LogTag_Debug
End Sub
```

2. マクロを実行し、イミディエイトウィンドウにログが出力されることを確認してください。

---

より詳細な使い方・設定例は `docs/USER_GUIDE.md` をご覧ください。
アドインの拡張方法や内部構造については `docs/DEVELOPER_GUIDE.md` を参照してください。


## ライセンス
このプロジェクトはMITライセンスの下で公開されています。

### 第三者ライブラリ
このプロジェクトは以下の第三者ライブラリを使用しています：

#### vbac.wsf (Ariawase Library)
- **ライセンス**: MIT License
- **著作権**: Copyright (c) 2011 igeta
- **プロジェクトページ**: https://github.com/vbaidiot/ariawase
- **変更内容**: @Folderアノテーションサポート機能を追加（フォルダ階層の自動構築）
