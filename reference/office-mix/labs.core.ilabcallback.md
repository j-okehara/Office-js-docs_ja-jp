
# Labs.Core.ILabCallback

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

The interface for handling Labs.js callback methods.

```
interface ILabCallback<T>
```


## Callback signature

 `(err: any, data: T): void`

 **Callback parameters**


|||
|:-----|:-----|
| _err_|エラーが発生していない場合は、**null** です。エラーが発生した場合は、**null** 以外です。|
| _data_|The data returned with the callback.|
