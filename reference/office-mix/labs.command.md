
# <a name="labs.command"></a>Labs.Command

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

クライアントとホスト間でメッセージを渡すのに使用される一般的なコマンドです。

```
class Command
```


## <a name="properties"></a>プロパティ


|**名前**|**説明**|
|:-----|:-----|
| `public var type: string`|コマンドの型。|
| `public var commandData: any`|コマンドに関連付けられている省略可能なデータ。|

## <a name="methods"></a>メソッド




### <a name="constructor"></a>コンストラクター

 `function constructor(type: string, commandData?: any)`

説明

 **パラメーター**


|||
|:-----|:-----|
| `type`|コマンドの型。|
| `commandData`|コマンドに関連付けられている省略可能なデータ。|
