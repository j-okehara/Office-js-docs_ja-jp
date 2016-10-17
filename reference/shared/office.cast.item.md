
# <a name="office.cast.item-property"></a>Office.cast.item プロパティ
新規作成モードまたは閲覧モードのメッセージおよび予定に固有の IntelliSense を提供します。

|||
|:-----|:-----|
|**ホスト:**|Outlook|
|**[要件セット](../../docs/overview/specify-office-hosts-and-api-requirements.md)に指定できるもの**|メールボックス|
|**最終変更バージョン**|1.0|



|||
|:-----|:-----|
|**適用可能な Outlook のモード**|Visual Studio のデザイン時のみ|

```js
Office.cast.item.toAppointmentCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toAppointmentRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toAppointment(Office.context.mailbox.item);
```

```js
Office.cast.item.toItemCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toItemRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessageCompose(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessageRead(Office.context.mailbox.item);
```

```js
Office.cast.item.toMessage(Office.context.mailbox.item);
```


## <a name="return-value"></a>戻り値

Outlook アドイン用の適切な IntelliSense を選択できるようにする一連のメソッド。


## <a name="remarks"></a>注釈

このプロパティとそのメソッドは、Outlook アドインを Visual Studio 上で開発する場合にのみ IntelliSense をサポートします。他の開発ツールには効果がありません。

**Office.cast.item** のメソッドは、**Office.context.mailbox.item** プロパティに固有の IntelliSense を提供するために Visual Studio でデザイン時に使用されます。たとえば、**toAppointmentCompose** メソッドを使用するときには、作成モードに適用される **Appointment** のメソッドおよびプロパティのみが IntelliSense により表示されます。

実行時に  **Office.cast.item** のメソッドは Outlook アドインに影響を与えません。


## <a name="example"></a>例

次の例では **toMessageCompose** メソッドを使用して **Office.context.mailbox.item** プロパティをキャストして、新規作成モードで **Message** オブジェクトの IntelliSense のみが表示されるようにします。キャストした後、 `message` 変数には、新規作成モードで使用できるメソッドおよびプロパティの IntelliSense のみが表示されます。


```js
var message = Office.cast.item.toMessageCompose(Office.context.mailbox.item);

```


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、このメソッドは、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションでこのメソッドをサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。

||Windows デスクトップ版 Office|Office Online (ブラウザー)|Outlook for Mac|
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



|**バージョン**|**変更内容**|
|:-----|:-----|
|1.0|導入|
