
# <a name="labs.imessagehandler"></a>Labs.IMessageHandler

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

イベント ハンドラーを定義できるようにするインターフェイス。

```
interface IMessageHandler(origin: Window, data: any, callback: Labs.Core.ILabCallback<any>): void
```


## 

 **Parameters**


|||
|:-----|:-----|
| `origin`|メッセージの送信元であるラボのウィンドウ。|
| `data`|The contents of the message.|
| `callback`|Callback function that fires once the message is received.|
