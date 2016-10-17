
# <a name="projectprojectfields-enumeration"></a>ProjectProjectFields 列挙型
**[getProjectFieldAsync](../../reference/shared/projectdocument.getprojectfieldasync.md)** メソッドのパラメーターとして使用できるプロジェクト フィールドを指定します。

|||
|:-----|:-----|
|**ホスト:**|Project|
|**追加されたバージョン**|1.0|

```
ProjectProjectFields={
    CurrencyDigits: 0, 
    CurrencySymbol: 1, 
    CurrencySymbolPosition: 2, 
    DurationUnits: 3,
    GUID: 4, 
    Finish: 5, 
    Start: 6, 
    ReadOnly: 7, 
    VERSION: 8, 
    WorkUnits: 9, 
    ProjectServerUrl: 10, 
    WSSUrl: 11, 
    WSSList: 12
}
```


## <a name="members"></a>メンバー


****


|**メンバー**|**説明**|
|:-----|:-----|
|**CurrencyDigits**|通貨の小数点以下の桁数。|
|**CurrencySymbol**|通貨記号。|
|**CurrencySymbolPosition**|通貨記号の位置: 指定しない = -1、値の前、スペースなし ($0) = 0、値の後、スペースなし (0$) = 1、値の前、スペースあり ($ 0) = 2、値の後、スペースあり (0 $) = 3。|
|**GUID**|プロジェクトの GUID。|
|**Finish**|プロジェクトの終了日。|
|**Start**|プロジェクトの開始日。|
|**ReadOnly**|プロジェクトが読み取り専用かどうかを指定します。|
|**VERSION**|プロジェクトのバージョン。|
|**WorkUnits**|プロジェクトの作業単位 (日数、時間数など)。|
|**ProjectServerUrl**|Project Server に保存されるプロジェクトの Project Web App の URL。|
|**WSSUrl**|SharePoint リストと同期されるプロジェクトの SharePoint URL。|
|**WSSList**|タスク リストと同期されるプロジェクトの SharePoint リストの名前。|

## <a name="remarks"></a>注釈

**ProjectProjectFields** 定数は、**[getProjectFieldAsync](../../reference/shared/projectdocument.getprojectfieldasync.md)** メソッドのパラメーターとして使用できます。


## <a name="support-details"></a>サポートの詳細


次の表で、大文字 Y は、この列挙は、対応する Office ホスト アプリケーションでサポートされていることを示します。空のセルは、Office ホスト アプリケーションがこの列挙をサポートしないことを示します。

Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」をご覧ください。


**サポートされるホスト (プラットフォーム別)**


||**Windows デスクトップ版 Office**|**Office Online (ブラウザー)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**アドインの種類**|作業ウィンドウ|
|**ライブラリ**|Office.js|
|**名前空間**|Office|

## <a name="support-history"></a>サポート履歴



****


|**バージョン**|**変更内容**|
|:-----|:-----|
|1.0|導入|

## <a name="see-also"></a>関連項目



#### <a name="other-resources"></a>その他の技術情報


[getProjectFieldAsync メソッド](../../reference/shared/projectdocument.getprojectfieldasync.md)
