
# <a name="officetheme.controlforegroundcolor-property"></a>officeTheme.controlForegroundColor プロパティ
Office テーマのコントロールの前景色を取得します。

 **重要:**現在、この API は、Windows デスクトップの [Office 2016 プレビュー](https://products.office.com/en-us/office-2016-preview)の Excel、Outlook、PowerPoint、および Word でのみ機能します。



|||
|:-----|:-----|
|**ホスト:**|Excel、Outlook、PowerPoint、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|セットには指定できない|
|**追加されたバージョン**|1.3|

```js
var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;
```


## <a name="return-value"></a>戻り値

16 進数の色の組み合わせ


## <a name="remarks"></a>解説

返される色は、 **[ファイル]**  >  **[Office アカウント]**  >  **[Office テーマ]** UI でユーザーが選択した Office テーマの値に関連付けられています。これは、Office ホスト アプリケーション全体に適用されます。


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|**デバイス用 OWA**|
|:-----|:-----|:-----|:-----|:-----|
|**Excel**|Y||||
|**Outlook**|Y||||
|**PowerPoint**|Y||||
|**Word**|Y||||



|||
|:-----|:-----|
|**最小限のアクセス許可レベル**|[制限あり](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ、Outlook|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



|**バージョン**|**変更内容**|
|:-----|:-----|
|1.3|導入|
