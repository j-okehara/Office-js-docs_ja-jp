
# <a name="labs.core.ilabcallback"></a>Labs.Core.ILabCallback

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

Labs.js コールバック メソッドを処理するインターフェイス。

```
interface ILabCallback<T>
```


## <a name="callback-signature"></a>コールバック シグネチャ

 `(err: any, data: T): void`

 **コールバック パラメーター**


|||
|:-----|:-----|
| _err_|エラーが発生していない場合は、**null** です。エラーが発生した場合は、**null** 以外です。|
| _data_|コールバックで返されるデータ。|
