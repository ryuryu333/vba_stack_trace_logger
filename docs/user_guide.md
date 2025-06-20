# VBA Stack Trace Logger

作業中です...  

## ロガーの機能へアクセスする
`MyLogger.` と打ち込む事で、ロガーの全ての機能にアクセス可能です。  

`.` を記入すると VBE にてインテリセンスが表示されるので、使用したい機能を選択してください。  

### Note インテリセンスを頼ろう
VBE にて `Ctrl` + `Space` を押すと、インテリセンスを表示できます。  
本ロガーではインテリンス機能に対応したコードを作成しており、`MyLogger` 以外のワードは暗記しなくても良いようにしています。  

## ロガーの挙動をカスタマイズする
ロガーの設定を行います。  
初期化処理を兼ねているので、ロガーを利用する前に必ず実行してください。  

```vba
' デフォルトの設定を利用する場合
MyLogger.StartConfiguration.Build

' カスタマイズする場合
MyLogger.StartConfiguration.SomeSetting.Build
```

デフォルトの設定は最低限の構成にしています。
- スタックトレース機能は無効化
- ログ出力先は VBE イミディエイトウインドウのみ

以下のプロシージャをメソッドチェーンとして使用することで、各設定を反映させることができます。

| プロシージャ名                    | 説明                                                                 |
|---------------------------|----------------------------------------------------------------------|
| `DisableLogging`          | ログ出力を無効化します。                                            |
| `EnableTagFiltering`      | 指定されたタグのログ出力を無効化します。                           |
| `EnableStackTrace`        | スタックトレース機能を有効化します。                               |
| `DisableWriteToImmediate` | VBE のイミディエイトウィンドウへのログ出力を無効化します。         |
| `EnableWriteToExcelSheet` | 指定されたエクセルシートへのログ出力を有効化します。               |

**EnableTagFiltering** を使用した場合、次のメソッドチェーンは `Add` と `Apply` に限定されます。  
`Add(LogTag_Debug)` のようにログ出力を無効化したいタグを指定します。  
複数のタグを指定したい場合は、`Add` を繰り返します。  
指定が完了したら `Apply` を使用してください。  

**EnableWriteToExcelSheet** を使用した場合、次のメソッドチェーンは `SetOutputExcelSheet` に限定されます。  
`SetOutputExcelSheet(ActiveSheet)` のように `Worksheet` 型の値を引数に渡してください。  
指定されたシートにログ出力が行われます。  
ロガー初期化処理でシートの内容は全てクリアされるので、ログ専用のシートを準備してください。  

全ての設定を指定し終えたら、`Build` を使用します。

### ログを出力する

```vba
' 通常
MyLogger.Log "Your Message"

' タグを指定する場合（未指定だと INFO）
MyLogger.Log "Your Message", LogTag_Debug
```

## ロガーの終了処理を行う

```
myLogger.Terminate
```
