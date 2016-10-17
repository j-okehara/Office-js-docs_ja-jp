

# <a name="office-object"></a>Office オブジェクト
アドインのインスタンスを表します。これは、API の最上位レベルのオブジェクトへのアクセスを提供します。

|||
|:-----|:-----|
|**ホスト:**|Access、Excel、Outlook、PowerPoint、Project、Word|
|**最終変更バージョン**|1.1|

```js
Office
```


## <a name="members"></a>メンバー


**プロパティ**

|||
|:-----|:-----|
|名前|説明|
|[context](../../reference/shared/office.context.md)|アドインのランタイム環境を表す Context オブジェクトを取得し、API の最上位レベルのオブジェクトへのアクセスを提供します。|
|[cast.item](../../reference/shared/office.cast.item.md)|新規作成モードまたは閲覧モードのメッセージおよび予定に固有の Visual Studio の IntelliSense を提供します。 <br/><br/><blockquote>**メモ**  Visual Studio で Outlook アドインを開発するデザイン時にのみ適用できます。 </blockquote>|

**メソッド**

|||
|:-----|:-----|
|名前|説明|
|[select](../../reference/shared/office.select.md)|渡されるセレクター文字列に基づくバインドを返す promise を作成します。|
|[useShortNamespace](../../reference/shared/office.useshortnamespace.md)|**Microsoft.Office.WebExtension** という完全な名前空間に対して **Office** というエイリアスを使用するかどうかを切り替えます。|

**イベント**

|||
|:-----|:-----|
|名前|説明|
|[initialize](../../reference/shared/office.initialize.md)|ランタイム環境が読み込まれ、アプリケーションやホストされたドキュメントを対話操作するアドインの準備ができたときに発生します。|

## <a name="remarks"></a>注釈

**Office** オブジェクトを使用すると、開発者は Initialize イベントに対してコールバック関数を実装できます。また、Office オブジェクトは、[Context](../../reference/shared/context.md) オブジェクトへのアクセスを提供します。


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このオブジェクトは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのオブジェクトをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|**Office for iPad**|**デバイス用 OWA**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**アドインの種類**|コンテンツ、Outlook、作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.1|Office for iPad で Excel、PowerPoint、および Word のサポートが追加されました。|
|1.1|<ul><li><a href="6c4b2c16-d4fb-4ecf-b72c-1e33b205daaf.htm">context</a> で、Access 用コンテンツ アドインにおけるランタイム コンテンツの取得のサポートが追加されました。</p></li><li><p><a href="23aeb136-da1f-4127-a798-99dc27bc4dae.htm">select</a> で、Access 用コンテンツ アドインにおけるテーブル バインドの選択のサポートが追加されました。</li><li><a href="9a4d5c7d-fcc4-4e8f-bef2-f2a8d8b4ae00.htm">useShortNamespace</a> で、Access 用コンテンツ アドインのサポートが追加されました。</li><li><a href="727adf79-a0b5-48d2-99c7-6642c2c334fc.htm">initialize</a> で、Access 用コンテンツ アドインにおける初期化のサポートが追加されました。</li></ul>|
|1.0|導入|

