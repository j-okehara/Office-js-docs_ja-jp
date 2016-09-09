
# Labs.LabEditor

 _**適用対象:** Office ???? | Office ???? | Office Mix | PowerPoint_

The  **LabEditor** object allows you to edit a given lab as well as get and set configuration data associated with the lab.

```
class LabEditor
```


## メソッド


### getConfiguration

 `public function getConfiguration(callback: Labs.Core.ILabCallback<Labs.Core.IConfiguration>): void`

Retrieves the current lab configuration.

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _callback_|Callback function that is fired once the configuration has been retrieved.|

### setConfiguration

 `public function getConfiguration(callback: Labs.Core.ILabCallback<Labs.Core.IConfiguration>): void`

Sets a new lab configuration.

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _configuration_|The configuration to set.|
| _callback_|Callback function that is fired once the configuration has been set.|

### done

 `public function done(callback: Labs.Core.ILabCallback<void>): void`

Indicates that the user has finished editing the lab.

 **パラメーター**


|**名前**|**説明**|
|:-----|:-----|
| _callback_|Callback function that is fired once the lab editor has finished.|
