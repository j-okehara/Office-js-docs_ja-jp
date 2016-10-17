
# <a name="document.settings-property"></a>Document.settings プロパティ
現在のドキュメントのコンテンツ アプリまたは作業ウィンドウ アプリの保存されているカスタム設定を表すオブジェクトを取得します。

|||
|:-----------------|:--------------------------------|
| ホスト:           | Access、Excel、PowerPoint、Word |
| 最終変更バージョン: | 1.1                             |

```js
var _settings = Office.context.document.settings;
```

## <a name="return-value"></a>戻り値

[Settings](./settings.md) オブジェクト。

## <a name="support-details"></a>サポートの詳細

次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。

**サポートされるホスト (プラットフォーム別)**

|             | Windows デスクトップ版 Office | Office Online (ブラウザー) | Office for iPad |
|:------------|:---------------------------|:---------------------------|:----------------|
| Access      |                            | Y                          |                 |
| Excel       | Y                          | Y                          | Y               |
| PowerPoint  | Y                          | Y                          | Y               |
| Word        | Y                          | Y                          | Y               |

|||
|:--------------------------|:-----|
| 最小限のアクセス許可レベル  | [制限あり](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
| アドインの種類:             | コンテンツ、作業ウィンドウ
| ライブラリ:                  | Office.js
| 名前空間:                | Office

## <a name="support-history"></a>サポート履歴

| 変更内容 | 1.1 |
|:--------|:--------|
| 1.1     |Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。
| 1.1     |Access 用コンテンツのアドインのサポートが追加されました。
| 1.0     |導入
