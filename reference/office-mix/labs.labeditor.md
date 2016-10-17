
# <a name="labs.labeditor"></a>Labs.LabEditor

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

**LabEditor** オブジェクトを使うと、指定されたラボの編集に加えて、ラボに関連付けられている構成データを取得し、設定できます。

```
class LabEditor
```


## <a name="methods"></a>メソッド


### <a name="getconfiguration"></a>getConfiguration

 `public function getConfiguration(callback: Labs.Core.ILabCallback<Labs.Core.IConfiguration>): void`

現在のラボの構成を取得します。

 **Parameters**


|**名前**|**説明**|
|:-----|:-----|
| _callback_|構成が取得されると起動するコールバック関数。|

### <a name="setconfiguration"></a>setConfiguration

 `public function getConfiguration(callback: Labs.Core.ILabCallback<Labs.Core.IConfiguration>): void`

新しいラボ構成を設定します。

 **Parameters**


|**名前**|**説明**|
|:-----|:-----|
| _configuration_|設定する構成。|
| _callback_|構成が設定されると起動するコールバック関数。|

### <a name="done"></a>done

 `public function done(callback: Labs.Core.ILabCallback<void>): void`

ユーザーがラボの編集を完了したことを示します。

 **Parameters**


|**名前**|**説明**|
|:-----|:-----|
| _callback_|ラボのエディターが完了すると起動するコールバック関数。|
