
# <a name="context.roamingsettings-property"></a>Context.roamingSettings プロパティ
ユーザーのメールボックスに保存されている、Outlook アドインのカスタム設定または状態を表すオブジェクトを取得します。

|||
|:-----|:-----|
|**ホスト:**|Outlook|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|メールボックス|
|**最終変更バージョン**|1.0|

```
var appSettings = office.context.roamingSettings;
```


## <a name="return-value"></a>戻り値


  [RoamingSettings](http://msdn.microsoft.com/library/cf21bb08-7274-4ad6-ae9e-b2c12f92abc9%28Office.15%29.aspx) オブジェクト。


## <a name="remarks"></a>注釈

**RoamingSettings** オブジェクトを使用すると、ユーザーのメールボックスに保存されている、Outlook アドインのデータの保存やアクセスを実行できます。そのため、Outlook アドインは、このメールボックスへのアクセスに使用するどのホスト クライアント アプリケーションから実行されても、このデータを使用できます。


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|
|**Outlook**|Y|Y|Y|

|||
|:-----|:-----|
|**要件セットに指定できるもの**|メールボックス|
|**最小限のアクセス許可レベル**|[制限あり](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|Outlook|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



****


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.0|導入|
