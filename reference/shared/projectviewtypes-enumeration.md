
# <a name="projectviewtypes-enumeration"></a>ProjectViewTypes 列挙型
**[getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)** メソッドで認識できるビューの種類を指定します。

|||
|:-----|:-----|
|**ホスト:**|Project|
|**追加されたバージョン**|1.0|

```
ProjectViewTypes={
    Gantt           : 1, 
    NetworkDiagram  : 2, 
    TaskDiagram     : 3, 
    TaskForm        : 4, 
    TaskSheet       : 5, 
    ResourceForm    : 6, 
    ResourceSheet   : 7, 
    ResourceGraph   : 8, 
    TeamPlanner     : 9, 
    TaskDetails     : 10, 
    TaskNameForm    : 11, 
    ResourceNames   : 12, 
    Calendar        : 13, 
    TaskUsage       : 14, 
    ResourceUsage   : 15, 
    Timeline        : 16
}
```


## <a name="members"></a>メンバー


****


|**メンバー**|**説明**|
|:-----|:-----|
|**Gantt**|ガント チャート ビュー。|
|**NetworkDiagram**|ネットワーク ダイアグラム ビュー。|
|**TaskDiagram**|タスク ダイアグラム ビュー。|
|**TaskForm**|タスク フォーム ビュー。|
|**TaskSheet**|タスク シート ビュー。|
|**ResourceForm**|リソース フォーム ビュー。|
|**ResourceSheet**|リソース シート ビュー。|
|**ResourceForm**|リソース フォーム ビュー。|
|**ResourceGraph**|リソース グラフ ビュー。|
|**TeamPlanner**|チーム プランナー ビュー。|
|**TaskDetails**|タスクの詳細ビュー。|
|**TaskNameForm**|タスク名フォームの簡易ビュー。|
|**ResourceNames**|リソース名ビュー。|
|**Calendar**|カレンダー ビュー。|
|**TaskUsage**|タスク配分状況ビュー。|
|**ResourceUsage**|リソース配分状況ビュー。|
|**TimeLine**|タイムライン ビュー。|

## <a name="remarks"></a>注釈

**[getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)** メソッドは、アクティブ ビューに対応する **ProjectViewTypes** の定数値と名前を返します。


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


[getSelectedViewAsync メソッド](../../reference/shared/projectdocument.getselectedviewasync.md)
