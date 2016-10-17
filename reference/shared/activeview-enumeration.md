
# <a name="activeview-enumeration"></a>ActiveView 列挙
ユーザーがドキュメントを編集できるかどうかなど、ドキュメントのアクティブなビューの状態を指定します。

|||
|:-----|:-----|
|**導入された Office.js バージョン**|1.1|

|||
|:-----|:-----|
|**ホスト:**|PowerPoint|
|**追加されたバージョン**|1.1|



```
Office.ActiveView
```


## <a name="members"></a>メンバー


**値**


|**列挙**|**値**|**説明**|
|:-----|:-----|:-----|
|Office.ActiveView.Read|"read"|ホスト アプリケーションのアクティブ ビューでは、ドキュメントのコンテンツの読み取りのみが許可されます。|
|Office.ActiveView.Edit|"edit"|ホスト アプリケーションのアクティブ ビューで、ドキュメントのコンテンツを編集できます。|

## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、この列挙は、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションがこの列挙をサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|

|||
|:-----|:-----|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



****


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad で PowerPoint のサポートが追加されました。|
|1.1|導入|
