

# <a name="settings.set-method"></a>Settings.set メソッド
指定された設定を行うかまたは作成します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、PowerPoint、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|設定値|
|**最終変更バージョン**|1.1|

```js
Office.context.document.settings.set(name, value);
```


## <a name="parameters"></a>パラメーター



_name_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;型:  **string**

&nbsp;&nbsp;&nbsp;&nbsp;設定または作成する設定の名前 (大文字と小文字を区別します)。

    
_value_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;型: **string**、**number**、**boolean**、**null**、**object**、**array**

&nbsp;&nbsp;&nbsp;&nbsp;格納する値を指定します。
    

## <a name="remarks"></a>解説

**set**メソッドは、指定された名前の設定がまだ存在しない場合に新しい設定を作成するか、または設定プロパティ バッグのメモリ内コピーにある指定された名前の既存の設定に設定します。[Settings.saveAsync](../../reference/shared/settings.saveasync.md) メソッドを呼び出した後で、その値はそのデータ型のシリアル化された JSON 表現としてドキュメントに格納されます。各アドインの設定に最大 2 MB を使用できます。


 >**重要**:  **Settings.set** メソッドは、設定プロパティ バッグのメモリ内コピーに対してのみ動作します。再度ドキュメントを開いたときにも設定に対する追加や変更がアドインに反映されるようにするには、**Settings.set** メソッドの呼び出し後アドインを閉じるまでの間に **Settings.saveAsync** メソッドを呼び出して、ドキュメントに設定を保存する必要があります。


## <a name="example"></a>例




```js
function setMySetting() {
    Office.context.document.settings.set('mySetting', 'mySetting value');
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
|**Word**|Y|Y|Y|

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
|1.1|Access 用コンテンツ アドインにおけるカスタム設定のサポートが追加されました。|
|1.0|導入|
