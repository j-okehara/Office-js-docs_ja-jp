
# Labs.Core.ILabHost

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

Provides an abstraction layer for connecting Labs.js to the host.

```
interface ILabHost
```


## メソッド


### getSupportedVersions

 `getSupportedVersions(): Core.ILabHostVersionInfo[]`

Retrieves the versions supported by the lab host.

 **パラメーター**

なし。


### connect

 `connect(versions: Core.ILabHostVersionInfo[], callback: Core.ILabCallback<Core.IConnectionResponse>)`

ホストとの接続を初期化します。

 **パラメーター**


|||
|:-----|:-----|
| _versions_|List of host versions that the client can make use of.|
| _callback_|Callback function that fires when the connection is complete.|

### disconnect

 `disconnect(callback: Core.ILabCallback<void>)`

Terminates communication with the host.

 **パラメーター**


|||
|:-----|:-----|
| _completionStatus_|Status of the lab at the time of the disconnection.|
| _callback_|Callback function that fires when the disconnect is complete.|

### on

 `on(handler: (string: any, any: any): void)`

ホストからのメッセージを処理するために、イベント ハンドラーを追加します。 解決された約束は、ホストに返されます。

 **パラメーター**


|||
|:-----|:-----|
| _handler_|The event handler.|

### sendMessage

 `sendMessage(type: string, options: Core.IMessage, callback: Core.ILabCallback<Core.IMessageResponse>)`

Sends a message to the host.

 **パラメーター**


|||
|:-----|:-----|
| _type_|The type of message being sent.|
| _オプション_|Message options.|
| _callback_|Callback function that fires once the message is received.|

### create

 `create(options: Core.ILabCreationOptions, callback: Core.ILabCallback<void>)`

ラボを作成します。 ホスト情報を格納し、構成とその他の要素を格納するための場所を確保します。

 **パラメーター**


|||
|:-----|:-----|
| _オプション_|Options passed as part of the create operation.|
| _callback_|Callback function that fires once the lab has been created.|

### getConfiguration

 `getConfiguration(callback: Core.ILabCallback<Core.IConfiguration>)`

Retrieves the current lab configuration from the host.

 **パラメーター**


|||
|:-----|:-----|
| _callback_|Callback function to retrieve the configuration information.|

### setConfiguration

 `setConfiguration(configuration: Core.IConfiguration, callback: Core.ILabCallback<void>)`

Sets a new lab configuration on the host.

 **パラメーター**


|||
|:-----|:-----|
| _構成_|The lab configuration that is set.|
| _callback_|Callback function that fires once the configuration is set.|

### getConfigurationInstance

 `getConfigurationInstance(callback: Core.ILabCallback<Core.IConfigurationInstance>)`

Retrieves the instance configuration for the lab.

 **パラメーター**


|||
|:-----|:-----|
| _callback_|Callback function that fires once the configuration instance has been retrieved.|

### getState

 `getState(callback: Core.ILabCallback<any>)`

Retrieves the current state of the lab for a given user.

 **パラメーター**


|||
|:-----|:-----|
| _completionStatus_|Callback function that returns the current lab state.|

### setState

 `setState(state: any, callback: Core.ILabCallback<void>)`

Sets the state of the lab for a given user.

 **パラメーター**


|||
|:-----|:-----|
| _state_|The lab state.|
| _callback_|Callback function that fires when state has been set.|

### takeAction

 `takeAction(type: string, options: Core.IActionOptions, callback: Core.ILabCallback<Core.IAction>)`

Takes an attempt at an action.

 **パラメーター**


|||
|:-----|:-----|
| _type_|Type of action.|
| _オプション_|Options provided with the action.|
| _callback_|Callback function that returns the final executed action.|

### takeAction

 `takeAction(type: string, options: Core.IActionOptions, result: Core.IActionResult, callback: Core.ILabCallback<Core.IAction>)`

Takes an action that has already been completed.

 **パラメーター**


|||
|:-----|:-----|
| _type_|Type of action.|
| _オプション_|Options provided with the action.|
| _result_|アクションの結果。|
| _callback_|Callback function that returns the final executed action.|

### getActions

 `getActions(type: string, options: Core.IGetActionOptions, callback: Core.ILabCallback<Core.IAction[]>)`

Takes an attempt at an action.

 **パラメーター**


|||
|:-----|:-----|
| _type_|Type of get action.|
| _オプション_|Options provided with the get action.|
| _callback_|Callback function that returns the list of completed actions.|
