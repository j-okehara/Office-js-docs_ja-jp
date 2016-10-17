
# <a name="context.officetheme-property"></a>Context.officeTheme プロパティ
Office テーマの色のプロパティにアクセスできるようにします。

 **重要:**現在、この API は、Windows デスクトップの [Office 2016 プレビュー](https://products.office.com/en-us/office-2016-preview)の Excel、Outlook、PowerPoint、および Word でのみ機能します。


|||
|:-----|:-----|
|**ホスト:**|Excel、Outlook、PowerPoint、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|セットには指定できない|
|**追加されたバージョン**|1.3|



```js
Office.context.officeTheme
```


## <a name="members"></a>メンバー


**プロパティ**

|||
|:-----|:-----|
|名前|説明|
|[bodyBackgroundColor ](../../reference/shared/office.context.bodybackgroundcolor.md)|Office テーマの本文の背景色を取得します。|
|[bodyForegroundColor](../../reference/shared/office.context.bodyforegroundcolor.md)|Office テーマの本文の前景色を取得します。|
|[controlBackgroundColor](../../reference/shared/office.context.controlbackgroundcolor.md)|Office テーマのコントロールの背景色を取得します。|
|[controlForegroundColor](../../reference/shared/office.context.controlforegroundcolor.md)|Office テーマのコントロールの前景色を取得します。|

## <a name="remarks"></a>解説

Office テーマの色を使用すると、**[ファイル]**  >  **[Office アカウント]**  >  **[Office テーマ]** UI によってユーザーが選択した現在の Office テーマに合わせてアドインの配色を調整できます。このテーマは Office ホスト アプリケーション全体に適用されます。Office テーマの色を使用することは、Outlook アドインと作業ウィンドウ アドインに適しています。


## <a name="example"></a>例


```js
function applyOfficeTheme(){
    // Get office theme colors.
    var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
    var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
    var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
    var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

    // Apply body background color to a CSS class.
    $('.body').css('background-color', bodyBackgroundColor);
}
```


## <a name="support-details"></a>サポートの詳細



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
