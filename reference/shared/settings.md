
# <a name="settings-object"></a>Settings オブジェクト
ホスト ドキュメントに名前/値のペアとして格納される、作業ウィンドウ アドインまたはコンテンツ アドインのカスタム設定を表します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、PowerPoint、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|設定値|
|**最終変更バージョン**|1.1|

```
Office.context.document.settings
```


## <a name="members"></a>メンバー


**メソッド**

|||
|:-----|:-----|
|名前|説明|
|[addHandlerAsync](../../reference/shared/settings.addhandlerasync.md)|**settingsChanged** イベントのイベント ハンドラーを追加します。|
|[get](../../reference/shared/settings.get.md)|指定された設定を取得します。|
|[refreshAsync](../../reference/shared/settings.refreshasync.md)|ドキュメントに保存されている設定をすべて読み取り、メモリ上にあるアドインのコピーに対してこれらの設定を更新します。|
|[remove](../../reference/shared/settings.remove.md)|指定された設定を削除します。|
|[removeHandlerAsync](../../reference/shared/settings.removehandlerasync.md)|**settingsChanged** イベントのイベント ハンドラーを削除します。|
|[saveAsync](../../reference/shared/settings.saveasync.md)|設定を保存します。|
|[set](../../reference/shared/settings.set.md)|指定された設定を行うかまたは作成します。|

**イベント**


|**名前**|**説明**|
|:-----|:-----|
|[settingsChanged](../../reference/shared/settings.settingschangedevent.md)|設定が変更されるときに発生します。|

## <a name="remarks"></a>注釈

**Settings** オブジェクトのメソッドを使用して作成される設定は、アドイン単位およびドキュメント単位で保存されます。つまり、これらの設定は、それを作成したアドインでのみ、かつ設定が保存されているドキュメントからのみ使用できます。

設定の名前は  **string** ですが、値には **string**、 **number**、 **boolean**、 **null**、 **object**、または  **array** を指定できます。

**Settings** オブジェクトは [Document](../../reference/shared/document.md) オブジェクトの一部として自動的に読み込まれます。Settings オブジェクトを使用するには、アドインがアクティブになったときに Document オブジェクトの [settings](../../reference/shared/document.settings.md) プロパティを呼び出します。開発者は、設定を削除または追加した後 [saveAsync](../../reference/shared/settings.saveasync.md) メソッドを呼び出してその設定をドキュメントに保存する必要があります。


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このオブジェクトは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのオブジェクトをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|設定値|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴

|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.1|**addHandlerAsync** メソッドと **removeHandlerAsync** メソッドについては、Access 用コンテンツ アドインにおけるイベントのイベント ハンドラーの追加と削除のサポートが追加されました。**get**、**refreshAsync**、**remove**、**saveAsync**、**set** メソッドについては、Access 用コンテンツ アドインにおけるカスタム設定のサポートが追加されました。|
|1.0|導入|