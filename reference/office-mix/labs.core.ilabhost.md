
# <a name="labs.core.ilabhost"></a>Labs.Core.ILabHost

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

Labs.js をホストに接続するための抽象レイヤーを提供します。

```
interface ILabHost
```


## <a name="methods"></a>メソッド


### <a name="getsupportedversions"></a>getSupportedVersions

 `getSupportedVersions(): Core.ILabHostVersionInfo[]`

ラボ ホストでサポートされているバージョンを取得します。

 **パラメーター**

なし。


### <a name="connect"></a>connect

 `connect(versions: Core.ILabHostVersionInfo[], callback: Core.ILabCallback<Core.IConnectionResponse>)`

ホストとの接続を初期化します。

 **パラメーター**


|||
|:-----|:-----|
| _versions_|クライアントが使用できるホストのバージョンをリストします。|
| _callback_|接続が完了すると起動するコールバック関数。|

### <a name="disconnect"></a>disconnect

 `disconnect(callback: Core.ILabCallback<void>)`

ホストとの通信を終了します。

 **パラメーター**


|||
|:-----|:-----|
| _completionStatus_|切断時のラボの状態。|
| _callback_|切断が完了すると起動するコールバック関数。|

### <a name="on"></a>on

 `on(handler: (string: any, any: any): void)`

ホストからのメッセージを処理するために、イベント ハンドラーを追加します。解決された約束は、ホストに返されます。

 **パラメーター**


|||
|:-----|:-----|
| _handler_|イベント ハンドラー。|

### <a name="sendmessage"></a>sendMessage

 `sendMessage(type: string, options: Core.IMessage, callback: Core.ILabCallback<Core.IMessageResponse>)`

ホストにメッセージを送信します。

 **パラメーター**


|||
|:-----|:-----|
| _type_|送信されるメッセージの種類。|
| _options_|メッセージ オプション。|
| _callback_|メッセージが受信されると起動するコールバック関数。|

### <a name="create"></a>create

 `create(options: Core.ILabCreationOptions, callback: Core.ILabCallback<void>)`

ラボを作成します。ホスト情報を格納し、構成とその他の要素を格納するための場所を確保します。

 **パラメーター**


|||
|:-----|:-----|
| _options_|作成処理の一環として渡されるオプション。|
| _callback_|ラボが作成されると起動するコールバック関数。|

### <a name="getconfiguration"></a>getConfiguration

 `getConfiguration(callback: Core.ILabCallback<Core.IConfiguration>)`

現在のラボ構成をホストから取得します。

 **パラメーター**


|||
|:-----|:-----|
| _callback_|構成情報を取得するコールバック関数。|

### <a name="setconfiguration"></a>setConfiguration

 `setConfiguration(configuration: Core.IConfiguration, callback: Core.ILabCallback<void>)`

ホスト上に新しいラボ構成を設定します。

 **パラメーター**


|||
|:-----|:-----|
| _configuration_|設定されるラボ構成。|
| _callback_|構成が設定されると起動するコールバック関数。|

### <a name="getconfigurationinstance"></a>getConfigurationInstance

 `getConfigurationInstance(callback: Core.ILabCallback<Core.IConfigurationInstance>)`

ラボのインスタンス構成を取得します。

 **パラメーター**


|||
|:-----|:-----|
| _callback_|構成インスタンスが取得されると起動するコールバック関数。|

### <a name="getstate"></a>getState

 `getState(callback: Core.ILabCallback<any>)`

特定のユーザー用のラボの現在の状態を取得します。

 **パラメーター**


|||
|:-----|:-----|
| _completionStatus_|現在のラボの状態を返すコールバック関数。|

### <a name="setstate"></a>setState

 `setState(state: any, callback: Core.ILabCallback<void>)`

特定のユーザー用のラボの状態を設定します。

 **パラメーター**


|||
|:-----|:-----|
| _state_|ラボの状態。|
| _callback_|状態が設定されると起動するコールバック関数。|

### <a name="takeaction"></a>takeAction

 `takeAction(type: string, options: Core.IActionOptions, callback: Core.ILabCallback<Core.IAction>)`

アクション時の試行を取得します。

 **パラメーター**


|||
|:-----|:-----|
| _type_|アクションのタイプ。|
| _options_|アクションで提供されるオプション。|
| _callback_|最後に実行したアクションを返すコールバック関数。|

### <a name="takeaction"></a>takeAction

 `takeAction(type: string, options: Core.IActionOptions, result: Core.IActionResult, callback: Core.ILabCallback<Core.IAction>)`

完了したアクションを取得します。

 **パラメーター**


|||
|:-----|:-----|
| _type_|アクションのタイプ。|
| _options_|アクションで提供されるオプション。|
| _result_|アクションの結果。|
| _callback_|最後に実行したアクションを返すコールバック関数。|

### <a name="getactions"></a>getActions

 `getActions(type: string, options: Core.IGetActionOptions, callback: Core.ILabCallback<Core.IAction[]>)`

アクション時の試行を取得します。

 **パラメーター**


|||
|:-----|:-----|
| _type_|get アクションのタイプ。|
| _options_|get アクションで提供されるオプション。|
| _callback_|完了したアクションのリストを返すコールバック関数。|
