
# Labs.IMessageHandler

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

イベント ハンドラーを定義できるようにするインターフェイスです。

```
interface IMessageHandler(origin: Window, data: any, callback: Labs.Core.ILabCallback<any>): void
```


## 

 **パラメーター**


|||
|:-----|:-----|
| `origin`|The lab window from which the message originated.|
| `data`|The contents of the message.|
| `callback`|Callback function that fires once the message is received.|
