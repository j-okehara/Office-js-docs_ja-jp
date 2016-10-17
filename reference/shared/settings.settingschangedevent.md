

# <a name="settings.settingschanged-event"></a>Settings.settingsChanged イベント
[Settings.saveAsync](../../reference/shared/settings.saveasync.md) メソッドによって、設定プロパティ バッグのメモリ内コピーがドキュメントに保存された場合に発生します。

|||
|:-----|:-----|
|**ホスト:**|Excel |
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|設定値|
|**最終変更バージョン**|1.0|

```js
Office.EventType.SettingsChanged
```


## <a name="remarks"></a>注釈

**settingsChanged** イベントのイベント ハンドラーを追加するには、[Settings](../../reference/shared/settings.addhandlerasync.md) オブジェクトの **addHandlerAsync** メソッドを使用します。

**settingsChanged** イベントは、アドインのスクリプトが **Settings.saveAsync** メソッドを呼び出して、設定のメモリ内コピーをドキュメント ファイルに保持した場合にのみ発生します。**settingsChanged** イベントは、[Settings.set](../../reference/shared/settings.set.md) または [Settings.remove](../../reference/shared/settings.remove.md) メソッドが呼び出された場合にはトリガーされません。

**settingsChanged** イベントは、アドインが共有 (共同編集) ドキュメントで使用されていて、2 人以上のユーザーが設定を同時に保存しようとした場合に、競合の可能性を処理できるように設計されています。


 >**重要**  アドインが Excel クライアントで実行されている場合、アドインのコードで **settingsChanged** イベントのハンドラーを登録できますが、このイベントが発生するのは、アドインが Excel Online で開かれているスプレッドシートと共に読み込まれ、_なおかつ_ 複数のユーザーがこのスプレッドシートで作業している (共同編集) 場合のみです。そのため、**settingsChanged** イベントが効率的にサポートされるのは、共同編集シナリオの Excel Online 内のみです。


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このイベントは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションがこのイベントをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。



||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||Y||

|||
|:-----|:-----|
|**要件セットに指定できるもの**|設定値|
|**最小限のアクセス許可レベル**|[制限あり](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴




|**バージョン**|**変更内容**|
|:-----|:-----|
|1.0|導入|
