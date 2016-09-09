
# Labs.Command

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

クライアントとホスト間でメッセージを渡すのに使用される一般的なコマンドです。

```
class Command
```


## プロパティ


|**名前**|**説明**|
|:-----|:-----|
| `public var type: string`|The type of the command.|
| `public var commandData: any`|Optional data associated with the command.|

## メソッド




### コンストラクター

 `function constructor(type: string, commandData?: any)`

説明

 **パラメーター**


|||
|:-----|:-----|
| `type`|The type of the command.|
| `commandData`|Optional data associated with the command.|
