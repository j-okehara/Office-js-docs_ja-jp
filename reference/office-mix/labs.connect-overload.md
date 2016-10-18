
# <a name="labs.connect-(overload)"></a>Labs.connect (overload)

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

ホストとの接続を初期化します。

```
function connect(labHost: Core.ILabHost, callback: Core.ILabCallback<Core.IConnectionResponse>)
```


## <a name="parameters"></a>パラメーター


|||
|:-----|:-----|
| _labHost_|省略可能。接続先の [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md)インスタンスです。ホストを指定しない場合は、[Labs.DefaultHostBuilder](../../reference/office-mix/labs.defaulthostbuilder.md) を使用して構築されます。|
| _callback_|接続が確立すると起動するコールバック。|

## <a name="return-value"></a>戻り値

ホストに接続を戻します。

