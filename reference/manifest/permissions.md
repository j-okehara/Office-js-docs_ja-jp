
# <a name="permissions-element"></a>Permissions 要素
Office アドインの API アクセスのレベルを指定します。最小特権の原則に基づいてアクセス許可を要求する必要があります。

 **アドインの種類:**コンテンツ、作業ウィンドウ、メール


## <a name="syntax:"></a>構文:

コンテンツ アドインおよび作業ウィンドウ アドインの場合:


```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

メール アドインの場合:




```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```


## <a name="contained-in:"></a>次に含まれる:

 _[OfficeApp](../../reference/manifest/officeapp.md)_


## <a name="remarks"></a>注釈

詳細については、「[コンテンツ アドインおよび作業ウィンドウ アドインでの API 使用のアクセス許可を要求する](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)」と「[Outlook アドインのアクセス許可について](../../docs/outlook/understanding-outlook-add-in-permissions.md)」をご覧ください。

