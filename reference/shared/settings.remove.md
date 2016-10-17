

# <a name="settings.remove-method"></a>Settings.remove メソッド
指定された設定を削除します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、PowerPoint、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|設定値|
|**最終変更バージョン**|1.1|

```js
Office.context.document.settings.remove(name);
```


## <a name="parameters"></a>パラメーター



_name_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;型:  **string**

&nbsp;&nbsp;&nbsp;&nbsp;削除する設定の名前 (大文字と小文字を区別)。

    



## <a name="remarks"></a>解説

 **null** は設定として有効な値です。したがって、 **null** を設定に割り当ててもその設定が設定プロパティ バッグから削除されるわけではありません。


 >**重要**: **Settings.remove** メソッドは、設定プロパティ バッグのメモリ内コピーに対してのみ動作します。指定した設定の削除をドキュメントに保存するには、**Settings.remove** メソッドの呼び出し後アドインを閉じるまでの間に [Settings.saveAsync](../../reference/shared/settings.saveasync.md) メソッドを呼び出す必要があります。


## <a name="example"></a>例




```js
function removeMySetting() {
    Office.context.document.settings.remove('mySetting');
}
```




## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

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
|**最小限のアクセス許可レベル**|[制限あり](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴




|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|PowerPoint Online のサポートが追加されました。|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.1|Access 用コンテンツ アドインにおけるカスタム設定の作成のサポートが追加されました。|
|1.0|導入|
