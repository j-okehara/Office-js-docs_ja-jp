
# <a name="binding-object"></a>Binding オブジェクト
ドキュメントのセクションへのバインドを表す抽象クラス。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Word|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|MatrixBinding, TableBinding, TextBinding|
|**TableBinding の最終変更**|1.1|

```js
Office.context.document.bindings.getByIdAsync(id);
```

## <a name="members"></a>メンバー


**オブジェクト**


|**名前**|**説明**|
|:-----|:-----|
|[MatrixBinding](../../reference/shared/binding.matrixbinding.md)|行と列の 2 次元でバインドを表現します。|
|[TableBinding](../../reference/shared/binding.tablebinding.md)|バインドを行と列の 2 次元で、必要に応じてヘッダーと共に表します。|
|[TextBinding](../../reference/shared/binding.textbinding.md)|ドキュメント内のバインドされているテキスト選択を表します。|

**プロパティ**


|**名前**|**説明**|
|:-----|:-----|
|[document](../../reference/shared/binding.document.md)|バインドに関連付けられた **Document** オブジェクトを取得します。|
|[id](../../reference/shared/binding.id.md)|オブジェクトの識別子を取得します。|
|[type](../../reference/shared/binding.type.md)|バインドの種類を取得します。|

**メソッド**


|**名前**|**説明**|
|:-----|:-----|
|[addHandlerAsync](../../reference/shared/binding.addhandlerasync.md)|指定されたイベントの種類のバインドにハンドラーを追加します。|
|[getDataAsync](../../reference/shared/binding.getdataasync.md)|バインド内に含まれるデータを返します。|
|[removeHandlerAsync](../../reference/shared/binding.removehandlerasync.md)|指定されたイベントの種類のバインドから、指定されたハンドラーを削除します。|
|[setDataAsync](../../reference/shared/binding.setdataasync.md)|指定されたバインド オブジェクトで表されるドキュメントのバインド セクションにデータを書き込みます。|
|[TableBinding.setFormatsAsync](../../reference/shared/binding.tablebinding.setformatsasync.md)|バインド テーブル内の指定のアイテムとデータの書式を設定または更新します。|

**イベント**


|**名前**|**説明**|
|:-----|:-----|
|[bindingDataChanged](../../reference/shared/binding.bindingdatachangedevent.md)|バインド内でデータが変更されるときに発生します。|
|[bindingSelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md)|バインド内で選択が変更されるときに発生します。|

## <a name="remarks"></a>注釈

**Binding** オブジェクトは、種類にかかわらず、すべてのバインドが所有する機能を公開します。

**Binding** オブジェクトが直接呼び出されることはありません。このオブジェクトは、バインドの種類 ([MatrixBinding](../../reference/shared/binding.matrixbinding.md)、[TableBinding](../../reference/shared/binding.tablebinding.md)、または [TextBinding](../../reference/shared/binding.textbinding.md)) を表すオブジェクトの抽象親クラスです。これら 3 つのオブジェクトはすべて、**Binding** オブジェクトから **getDataAsync** および **setDataAsync** メソッドを継承して、バインド内のデータを操作できます。また、**id** および **type** プロパティを継承して、これらのプロパティ値をクエリすることもできます。さらに、**MatrixBinding** および **TableBinding** オブジェクトは、行数と列数をカウントする機能など、マトリックスおよびテーブル固有の機能も公開します。


## <a name="support-details"></a>サポートの詳細


**Binding** オブジェクトの各 API メンバーのサポートは、Office のホスト アプリケーションの間で異なります。各メンバーのホストのサポート情報のトピックにある「サポートの詳細」セクションを参照してください。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


|||
|:-----|:-----|
|**要件セットに指定できるもの**|MatrixBinding, TableBinding, TextBinding|
|**アドインの種類**|コンテンツ、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|
