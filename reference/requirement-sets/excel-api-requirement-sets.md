# <a name="excel-javascript-api-requirement-sets"></a>Excel JavaScript API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のホストと API の要件を指定する](../../docs/overview/specify-office-hosts-and-api-requirements.md)」をご覧ください。

Excel アドインは、Office 2016 for Windows、Office for iPad、Office for Mac、Office Online など、複数のバージョンの Office で機能します。次の表は、Excel の要件セット、その要件セットをサポートする Office ホスト アプリケーション、それらのアプリケーションのビルド バージョンまたはビルド番号の一覧です。

|  要件セット  |  Office 2016 for Windows*  |  Office 2016 for iPad  |  Office 2016 for Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| ExcelApi 1.5 **Beta**  | バージョン 1702 (ビルド TBD) 以降| 近日公開 |  近日公開| 近日公開 | 近日公開|
| ExcelApi 1.4 **ベータ版** | バージョン 1702 (ビルド TBD) 以降| 近日公開 |  近日公開| 近日公開 | 近日公開|
| ExcelApi 1.3  | バージョン 1608 (ビルド 7369.2055) 以降| 1.27 以降 |  15.27 以降| 2016 年 9 月 | バージョン 1608 (ビルド 7601.6800) 以降|
| ExcelApi 1.2  | バージョン 1601 (ビルド 6741.2088) 以降 | 1.21 以降 | 15.22 以降| 2016 年 1 月 ||
| ExcelApi 1.1  | バージョン 1509 (ビルド 4266.1001) 以降 | 1.19 以降 | 15.20 以降| 2016 年 1 月 ||

> **注**:MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。このバージョンには、ExcelApi 1.1 の要件セットのみが含まれています。

バージョン、ビルド番号、および Office Online Server の詳細については以下を参照してください。

- [Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号](https://technet.microsoft.com/en-us/library/mt592918.aspx)
- [使用している Office のバージョンを確認する方法](https://support.office.com/en-us/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-US&rs=en-US&ad=US&fromAR=1)
- [Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。](https://technet.microsoft.com/en-us/library/mt592918.aspx#Anchor_1)
- [Office Online Server 概要](https://technet.microsoft.com/en-us/library/jj219437(v=office.16).aspx)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット
共通 API の要件セットについて詳しくは、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」を参照してください。

## <a name="whats-new-in-excel-javascript-api-14"></a>Excel JavaScript API 1.4 の新機能
要件セット 1.3 の Excel JavaScript API に新しく追加された点は次のとおりです。

### <a name="named-item-add-and-new-properties"></a>名前付きアイテムの追加と新しいプロパティ

新しいプロパティ
* `comment`
* `scope` ワークシートまたはブックの対象になるアイテム
* `worksheet` 名前付きアイテムの対象になるワークシートを返します。

新しいメソッド
* `add(name: string, reference: Range or string, comment: string)` は、新しい名前を指定したスコープのコレクションに追加します。
* `addFormulaLocal(name: string, formula: string, comment: string)` は、ユーザーのロケールを数式に使用して、新しい名前を指定したスコープのコレクションに追加します。

### <a name="settings-api-in-in-excel-namespace"></a>Excel の名前空間での Setting API

[Setting](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_1.4_OpenSpec/reference/excel/setting.md) オブジェクトは、ドキュメントに永続化されている設定のキーと値のペアを表します。ここでは、Excel の名前空間に設定関連の API を追加しました。これは純粋な新機能は提供しませんが、これにより約束ベースのバッチ API 構文を維持することが容易になり、Excel 関連タスクの共通 API に対する依存を減らすことができます。

API には、キーを使用して設定エントリを取得するための `getItem()` と、指定したキーと値の設定のペアをワークブックに追加するための `add()` が含まれています。

### <a name="others"></a>その他

* テーブルの列名を設定します (以前のバージョンでは読み取りのみ可能)。
* テーブルの列をテーブルの末尾に追加します (以前のバージョンでは末尾以外の任意の場所のみ可能)。
* 一度に複数の行をテーブルに追加します (以前のバージョンでは一度に 1 行のみ可能)。
* `range.getColumnsAfter(count: number)` および `range.getColumnsBefore(count: number)` を使用して、現在の Range オブジェクトの左右にある特定の数の列を取得します。
* アイテムまたは null オブジェクト関数。この機能により、キーを使用してオブジェクトを取得できます。オブジェクトが存在しない場合、返されたオブジェクトの isNullObject プロパティは true になります。これにより、開発者は例外処理を通じてオブジェクトを処理する必要なしに、オブジェクトが存在するかどうかを確認することができます。ワークシート、名前付きアイテム、バインド、グラフの系列などで使用できます。

`worksheet.GetItemOrNullObject()`

### <a name="suspend-calculation"></a>計算の中断
次の "context.sync()" が呼び出されるまで、計算を中断します (application.suspendCalculationUntilNextSync())。設定されると、依存関係が確実に伝達されるようにブックを再計算するのは開発者の責任です。

さらに、ダーティのセルを再計算していない F9 の再計算バグを修正しました。

|オブジェクト| 新機能| 説明|要件セット|
|:----|:----|:----|:----|
|[application](../excel/application.md)|_メソッド_ > [suspendCalculationUntilNextSync()](../excel/application.md#suspendcalculationuntilnextsync)|次の "context.sync()" が呼び出されるまで、計算を中断します。設定されると、依存関係が確実に伝達されるようにブックを再計算するのは開発者の責任です。|1.4|
|[bindingCollection](../excel/bindingcollection.md)|_メソッド_ > [getCount()](../excel/bindingcollection.md#getcount)|コレクション内にあるバインドの数を取得します。|1.4|
|[bindingCollection](../excel/bindingcollection.md)|_メソッド_ > [getItemOrNullObject(id: string)](../excel/bindingcollection.md#getitemornullobjectid-string)|ID によってバインド オブジェクトを取得します。バインディング オブジェクトが存在しない場合は null オブジェクトを返します。|1.4|
|[chartCollection](../excel/chartcollection.md)|_メソッド_ > [getCount()](../excel/chartcollection.md#getcount)|ワークシート上のグラフの数を返します。|1.4|
|[chartCollection](../excel/chartcollection.md)|_メソッド_ > [getItemOrNullObject(name: string)](../excel/chartcollection.md#getitemornullobjectname-string)|名前を使用してグラフを取得します。同じ名前の複数のグラフがある場合は、最初の 1 つが返されます。|1.4|
|[chartPointsCollection](../excel/chartpointscollection.md)|_メソッド_ > [getCount()](../excel/chartpointscollection.md#getcount)|系列内にあるグラフのポイントの数を取得します。|1.4|
|[chartSeriesCollection](../excel/chartseriescollection.md)|_メソッド_ > [getCount()](../excel/chartseriescollection.md#getcount)|コレクション内にあるデータ系列の数を取得します。|1.4|
|[namedItem](../excel/nameditem.md)|_プロパティ_ > comment|この名前に関連付けられているコメントを表します。|1.4|
|[namedItem](../excel/nameditem.md)|_プロパティ_ > scope|名前がブックを対象にしているのか、特定のワークシートを対象にしているのかを示します。読み取り専用です。使用可能な値は次のとおりです。Equal、Greater、GreaterEqual、Less、LessEqual、NotEqual。|1.4|
|[namedItem](../excel/nameditem.md)|_リレーションシップ_ > worksheet|名前付きのアイテムの対象になるワークシートを返します。アイテムがブックを対象にしている場合は、エラーをスローします。読み取り専用です。|1.4|
|[namedItem](../excel/nameditem.md)|_リレーションシップ_ > worksheetOrNullObject|名前付きのアイテムの対象になるワークシートを返します。アイテムがブックを対象にしている場合は、null オブジェクトを返します。読み取り専用です。|1.4|
|[namedItem](../excel/nameditem.md)|_メソッド_ > [delete()](../excel/nameditem.md#delete)|指定された名前を削除します。|1.4|
|[namedItem](../excel/nameditem.md)|_メソッド_ > [getRangeOrNullObject()](../excel/nameditem.md#getrangeornullobject)|名前に関連付けられている範囲オブジェクトを返します。名前付きアイテムの型が範囲でない場合は、null オブジェクトを返します。|1.4|
|[namedItemCollection](../excel/nameditemcollection.md)|_メソッド_ > [add(name: string, reference:Range または string, comment: string)](../excel/nameditemcollection.md#addname-string-reference-range-or-string-comment-string)|新しい名前を指定したスコープのコレクションに追加します。|1.4|
|[namedItemCollection](../excel/nameditemcollection.md)|_メソッド_ > [addFormulaLocal(name: string, formula: string, comment: string)](../excel/nameditemcollection.md#addformulalocalname-string-formula-string-comment-string)|ユーザーのロケールを数式に使用して、新しい名前を指定したスコープのコレクションに追加します。|1.4|
|[namedItemCollection](../excel/nameditemcollection.md)|_メソッド_ > [getCount()](../excel/nameditemcollection.md#getcount)|コレクション内の名前付きアイテムの数を取得します。|1.4|
|[namedItemCollection](../excel/nameditemcollection.md)|_メソッド_ > [getItemOrNullObject(name: string)](../excel/nameditemcollection.md#getitemornullobjectname-string)|名前を使用して、nameditem オブジェクトを取得します。nameditem オブジェクトが存在しない場合は null オブジェクトを返します。|1.4|
|[pivotTableCollection](../excel/pivottablecollection.md)|_メソッド_ > [getCount()](../excel/pivottablecollection.md#getcount)|コレクション内のピボット テーブルの数を取得します。|1.4|
|[pivotTableCollection](../excel/pivottablecollection.md)|_メソッド_ > [getItemOrNullObject(name: string)](../excel/pivottablecollection.md#getitemornullobjectname-string)|名前を使用してピボットテーブルを取得します。PivotTable が存在しない場合は null オブジェクトを返します。|1.4|
|[range](../excel/range.md)|_メソッド_ > [getIntersectionOrNullObject(anotherRange:Range or string)](../excel/range.md#getintersectionornullobjectanotherrange-range-or-string)|指定した範囲の長方形の交差部分を表す Range オブジェクトを取得します。交差部分が見つからない場合は、null オブジェクトを返します。|1.4|
|[range](../excel/range.md)|_メソッド_ > [getUsedRangeOrNullObject(valuesOnly: bool)](../excel/range.md#getusedrangeornullobjectvaluesonly-bool)|指定した範囲オブジェクトのうち使用されている範囲を返します。範囲内に使用済みのセルがない場合、この関数は null オブジェクトを返します。|1.4|
|[rangeViewCollection](../excel/rangeviewcollection.md)|_メソッド_ > [getCount()](../excel/rangeviewcollection.md#getcount)|コレクション内にある RangeView オブジェクトの数を取得します。|1.4|
|[setting](../excel/setting.md)|_プロパティ_ > key|Setting の ID を表すキーを返します。読み取り専用です。|1.4|
|[setting](../excel/setting.md)|_プロパティ_ > value|この設定に格納されている値を表します。|1.4|
|[setting](../excel/setting.md)|_メソッド_ > [delete()](../excel/setting.md#delete)|設定を削除します。|1.4|
|[settingCollection](../excel/settingcollection.md)|_プロパティ_ > items|setting オブジェクトのコレクション。読み取り専用です。|1.4|
|[settingCollection](../excel/settingcollection.md)|_メソッド_ > [add(key: string, value: (any)[])](../excel/settingcollection.md#addkey-string-value-any)|指定した設定をブックに設定または追加します。|1.4|
|[settingCollection](../excel/settingcollection.md)|_メソッド_ > [getCount()](../excel/settingcollection.md#getcount)|コレクション内にある Setting の数を取得します。|1.4|
|[settingCollection](../excel/settingcollection.md)|_メソッド_ > [getItem(key: string)](../excel/settingcollection.md#getitemkey-string)|キーから Setting エントリを取得します。|1.4|
|[settingCollection](../excel/settingcollection.md)|_メソッド_ > [getItemOrNullObject(key: string)](../excel/settingcollection.md#getitemornullobjectkey-string)|キーから Setting エントリを取得します。Setting が存在しない場合は null オブジェクトを返します。|1.4|
|[settingsChangedEventArgs](../excel/settingschangedeventargs.md)|_リレーションシップ_ > settings|SettingsChanged イベントが発生したバインドを表す Setting オブジェクトを取得します。|1.4|
|[tableCollection](../excel/tablecollection.md)|_メソッド_ > [getCount()](../excel/tablecollection.md#getcount)|コレクション内のテーブルの数を取得します。|1.4|
|[tableCollection](../excel/tablecollection.md)|_メソッド_ > [getItemOrNullObject(key: number or string)](../excel/tablecollection.md#getitemornullobjectkey-number-or-string)|名前または ID でテーブルを取得します。テーブルが存在しない場合は null オブジェクトを返します。|1.4|
|[tableColumnCollection](../excel/tablecolumncollection.md)|_メソッド_ > [getCount()](../excel/tablecolumncollection.md#getcount)|表の列数を取得します。|1.4|
|[tableColumnCollection](../excel/tablecolumncollection.md)|_メソッド_ > [getItemOrNullObject(key: number or string)](../excel/tablecolumncollection.md#getitemornullobjectkey-number-or-string)|名前または ID によって、列オブジェクトを取得します。列が存在しない場合は null オブジェクトを返します。|1.4|
|[tableRowCollection](../excel/tablerowcollection.md)|_メソッド_ > [getCount()](../excel/tablerowcollection.md#getcount)|表の行数を取得します。|1.4|
|[workbook](../excel/workbook.md)|_リレーションシップ_ > settings|ブックに関連付けられている Setting のコレクションを表します。読み取り専用です。|1.4|
|[worksheet](../excel/worksheet.md)|_リレーションシップ_ > names|現在のワークシートにスコープされている名前のコレクション。読み取り専用です。|1.4|
|[worksheet](../excel/worksheet.md)|_メソッド_ > [getUsedRangeOrNullObject(valuesOnly: bool)](../excel/worksheet.md#getusedrangeornullobjectvaluesonly-bool)|使用範囲とは、値または書式設定が割り当たっているすべてのセルを包含する最小の範囲です。ワークシート全体が空白の場合、この関数は null オブジェクトを返します。|1.4|
|[worksheetCollection](../excel/worksheetcollection.md)|_メソッド_ > [getCount(visibleOnly: bool)](../excel/worksheetcollection.md#getcountvisibleonly-bool)|コレクション内のワークシートの数を取得します。|1.4|
|[worksheetCollection](../excel/worksheetcollection.md)|_メソッド_ > [getItemOrNullObject(key: string)](../excel/worksheetcollection.md#getitemornullobjectkey-string)|名前または ID を使用して、ワークシート オブジェクトを取得します。ワークシートが存在しない場合は null オブジェクトを返します。|1.4|



## <a name="whats-new-in-excel-javascript-api-13"></a>Excel JavaScript API 1.3 の新機能
要件セット 1.3 の Excel JavaScript API に新しく追加された点は次のとおりです。

|オブジェクト| 新機能| 説明|要件セット|
|:----|:----|:----|:----|
|[binding](../excel/binding.md)|_メソッド_ > [delete()](../excel/binding.md#delete)|バインドを削除します。|1.3|
|[bindingCollection](../excel/bindingcollection.md)|_メソッド_ > [add(range:Range or string, bindingType: string, id: string)](../excel/bindingcollection.md#addrange-range-or-string-bindingtype-string-id-string)|特定の範囲に新しいバインドを追加します。|1.3|
|[bindingCollection](../excel/bindingcollection.md)|_メソッド_ > [addFromNamedItem(name: string, bindingType: string, id: string)](../excel/bindingcollection.md#addfromnameditemname-string-bindingtype-string-id-string)|ブック内の名前付きアイテムに基づいて新しいバインドを追加します。|1.3|
|[bindingCollection](../excel/bindingcollection.md)|_メソッド_ > [addFromSelection(bindingType: string, id: string)](../excel/bindingcollection.md#addfromselectionbindingtype-string-id-string)|現在の選択範囲に基づいて新しいバインドを追加します。|1.3|
|[bindingCollection](../excel/bindingcollection.md)|_メソッド_ > [getItemOrNull(id: string)](../excel/bindingcollection.md#getitemornullid-string)|ID を使用してバインド オブジェクトを取得します。バインド オブジェクトが存在しない場合、戻りオブジェクトの isNull プロパティは true になります。|1.3|
|[chartCollection](../excel/chartcollection.md)|_メソッド_ > [getItemOrNull(name: string)](../excel/chartcollection.md#getitemornullname-string)|グラフ名を使用してグラフを取得します。同じ名前の複数のグラフがある場合は、最初の 1 つが返されます。|1.3|
|[namedItemCollection](../excel/nameditemcollection.md)|_メソッド_ > [getItemOrNull(name: string)](../excel/nameditemcollection.md#getitemornullname-string)|nameditem オブジェクトを、名前を使用して取得します。nameditem オブジェクトが存在しない場合、返されたオブジェクトの isNull プロパティは true になります。|1.3|
|[pivotTable](../excel/pivottable.md)|_プロパティ_ > name|ピボットテーブルの名前。|1.3|
|[pivotTable](../excel/pivottable.md)|_リレーションシップ_ > worksheet|現在のピボットテーブルを含んでいるワークシート。読み取り専用。|1.3|
|[pivotTable](../excel/pivottable.md)|_メソッド_ > [refresh()](../excel/pivottable.md#refresh)|ピボットテーブルを更新します。|1.3|
|[pivotTableCollection](../excel/pivottablecollection.md)|_プロパティ_ > items|ピボットテーブル オブジェクトのコレクション。読み取り専用。|1.3|
|[pivotTableCollection](../excel/pivottablecollection.md)|_メソッド_ > [getItem(name: string)](../excel/pivottablecollection.md#getitemname-string)|名前を使用してピボットテーブルを取得します。|1.3|
|[pivotTableCollection](../excel/pivottablecollection.md)|_メソッド_ > [getItemOrNull(name: string)](../excel/pivottablecollection.md#getitemornullname-string)|名前を使用してピボットテーブルを取得します。ピボットテーブルが存在しない場合、戻りオブジェクトの isNull プロパティは true になります。|1.3|
|[range](../excel/range.md)|_メソッド_ > [getIntersectionOrNull(anotherRange:Range or string)](../excel/range.md#getintersectionornullanotherrange-range-or-string)|指定した範囲の長方形の交差部分を表す Range オブジェクトを取得します。交差部分が見つからない場合は、null オブジェクトを返します。|1.3|
|[range](../excel/range.md)|_メソッド_ > [getVisibleView()](../excel/range.md#getvisibleview)|現在の範囲の表示されている行を表します。|1.3|
|[rangeView](../excel/rangeview.md)|_プロパティ_ > cellAddresses|RangeView のセル アドレスを表します。読み取り専用。|1.3|
|[rangeView](../excel/rangeview.md)|_プロパティ_ > columnCount|表示されている列の数を返します。読み取り専用。|1.3|
|[rangeView](../excel/rangeview.md)|_プロパティ_ > formulas|A1 スタイル表記の数式を表します。|1.3|
|[rangeView](../excel/rangeview.md)|_プロパティ_ > formulasLocal|ユーザーの言語と数値書式ロケールで、A1 スタイル表記の数式を表します。たとえば、英語の数式 "=SUM(A1, introduced in 1.5" は、ドイツ語では "=SUMME(A1; 1,5)" になります。|1.3|
|[rangeView](../excel/rangeview.md)|_プロパティ_ > formulasR1C1|R1C1 スタイル表記の数式を表します。|1.3|
|[rangeView](../excel/rangeview.md)|_プロパティ_ > index|RangeView のインデックスを表す値を返します。読み取り専用。|1.3|
|[rangeView](../excel/rangeview.md)|_プロパティ_ > numberFormat|指定したセルの Excel の数値書式コードを表します。|1.3|
|[rangeView](../excel/rangeview.md)|_プロパティ_ > rowCount|表示されている行の数を返します。読み取り専用。|1.3|
|[rangeView](../excel/rangeview.md)|_プロパティ_ > text|指定した範囲のテキスト値。テキスト値は、セルの幅には依存しません。Excel UI で発生する # 記号による置換は、この API から返されるテキスト値には影響しません。読み取り専用です。|1.3|
|[rangeView](../excel/rangeview.md)|_プロパティ_ > valueTypes|各セルのデータの種類を表します。読み取り専用です。使用可能な値は次のとおりです。Unknown、Empty、String、Integer、Double、Boolean、Error。|1.3|
|[rangeView](../excel/rangeview.md)|_プロパティ_ > values|指定した範囲ビューの Raw 値を表します。返されるデータの型は、文字列、数値、ブール値のいずれかになります。エラーが含まれているセルは、エラー文字列を返します。|1.3|
|[rangeView](../excel/rangeview.md)|_リレーションシップ_ > rows|範囲に関連付けられている範囲ビューのコレクションを表します。読み取り専用。|1.3|
|[rangeView](../excel/rangeview.md)|_メソッド_ > [getRange()](../excel/rangeview.md#getrange)|現在の RangeView に関連付けられている親の範囲を取得します。|1.3|
|[rangeViewCollection](../excel/rangeviewcollection.md)|_プロパティ_ > items|rangeView オブジェクトのコレクション。読み取り専用。|1.3|
|[rangeViewCollection](../excel/rangeviewcollection.md)|_メソッド_ > [getItemAt(index: number)](../excel/rangeviewcollection.md#getitematindex-number)|RangeView のインデックスから RangeView の行番号を取得します。0 を起点とする番号になります。|1.3|
|[setting](../excel/setting.md)|_プロパティ_ > key|Setting の ID を表すキーを返します。読み取り専用。|1.3|
|[setting](../excel/setting.md)|_メソッド_ > [delete()](../excel/setting.md#delete)|設定を削除します。|1.3|
|[settingCollection](../excel/settingcollection.md)|_プロパティ_ > items|setting オブジェクトのコレクション。読み取り専用。|1.3|
|[settingCollection](../excel/settingcollection.md)|_メソッド_ > [getItem(key: string)](../excel/settingcollection.md#getitemkey-string)|キーから Setting エントリを取得します。|1.3|
|[settingCollection](../excel/settingcollection.md)|_メソッド_ > [getItemOrNull(key: string)](../excel/settingcollection.md#getitemornullkey-string)|キーから Setting エントリを取得します。Setting が存在しない場合、返されたオブジェクトの isNull プロパティは true になります。|1.3|
|[settingCollection](../excel/settingcollection.md)|_メソッド_ > [set(key: string, value: string)](../excel/settingcollection.md#setkey-string-value-string)|指定した設定をブックに設定または追加します。|1.3|
|[settingsChangedEventArgs](../excel/settingschangedeventargs.md)|_リレーションシップ_ > settingCollection|SettingsChanged イベントが発生したバインドを表す Setting オブジェクトを取得します。|1.3|
|[table](../excel/table.md)|_プロパティ_ > highlightFirstColumn|最初の列に特別な書式設定が含まれているかどうかを示します。|1.3|
|[table](../excel/table.md)|_プロパティ_ > highlightLastColumn|最後の列に特別な書式設定が含まれているかどうかを示します。|1.3|
|[table](../excel/table.md)|_プロパティ_ > showBandedColumns|テーブルを見やすくするため、奇数列を偶数列とは異なる方法で強調表示する書式設定にして、列を縞模様で表示するかどうかを示します。|1.3|
|[table](../excel/table.md)|_プロパティ_ > showBandedRows|テーブルを見やすくするため、奇数行を偶数行とは異なる方法で強調表示する書式設定にして、行を縞模様で表示するかどうかを示します。|1.3|
|[table](../excel/table.md)|_プロパティ_ > showFilterButton|フィルター ボタンを各列のヘッダーの上部に表示するかどうかを示します。これは、テーブルにヘッダー行が含まれている場合のみ設定できます。|1.3|
|[tableCollection](../excel/tablecollection.md)|_メソッド_ > [getItemOrNull(key: number or string)](../excel/tablecollection.md#getitemornullkey-number-or-string)|名前または ID を使用してテーブルを取得します。テーブルが存在しない場合、戻りオブジェクトの isNull プロパティは true になります。|1.3|
|[tableColumnCollection](../excel/tablecolumncollection.md)|_メソッド_ > [getItemOrNull(key: number or string)](../excel/tablecolumncollection.md#getitemornullkey-number-or-string)|名前または ID を使用して列オブジェクトを取得します。列が存在しない場合、返されたオブジェクトの isNull プロパティは true になります。|1.3|
|[workbook](../excel/workbook.md)|_リレーションシップ_ > pivotTables|ブックに関連付けられているピボットテーブルのコレクションを表します。読み取り専用。|1.3|
|[workbook](../excel/workbook.md)|_リレーションシップ_ > settings|ブックに関連付けられている Setting のコレクションを表します。読み取り専用。|1.3|
|[worksheet](../excel/worksheet.md)|_リレーションシップ_ > pivotTables|ワークシートの一部になっているピボットテーブルのコレクション。読み取り専用。|1.3|

## <a name="whats-new-in-excel-javascript-api-12"></a>Excel JavaScript API 1.2 の新機能
要件セット 1.2 の Excel JavaScript API に新たに追加された点は次のとおりです。

|オブジェクト| 新機能| 説明|要件セット|
|:----|:----|:----|:----|
|[chart](../excel/chart.md)|_プロパティ_ > id|コレクション内での位置を基にグラフを取得します。読み取り専用です。|1.2|
|[chart](../excel/chart.md)|_リレーションシップ_ > worksheet|現在のグラフを含んでいるワークシート。読み取り専用。|1.2|
|[chart](../excel/chart.md)|_メソッド_ > [getImage(height: number, width: number, fittingMode: string)](../excel/chart.md#getimageheight-number-width-number-fittingmode-string)|指定したサイズに合わせてグラフを拡大、縮小することで、グラフを Base64 でエンコードされた画像としてレンダリングします。|1.2|
|[filter](../excel/filter.md)|_リレーションシップ_ > criteria|指定した列に現在適用されているフィルターです。読み取り専用です。|1.2|
|[filter](../excel/filter.md)|_メソッド_ > [apply(criteria:FilterCriteria)](../excel/filter.md#applycriteria-filtercriteria)|指定した列に、指定されたフィルター条件を適用します。|1.2|
|[filter](../excel/filter.md)|_メソッド_ > [applyBottomItemsFilter(count: number)](../excel/filter.md#applybottomitemsfiltercount-number)|指定した数の要素の列に "下位アイテム" フィルターを適用します。|1.2|
|[filter](../excel/filter.md)|_メソッド_ > [applyBottomPercentFilter(percent: number)](../excel/filter.md#applybottompercentfilterpercent-number)|指定した割合の要素の列に "下位パーセント" フィルターを適用します。|1.2|
|[filter](../excel/filter.md)|_メソッド_ > [applyCellColorFilter(color: string)](../excel/filter.md#applycellcolorfiltercolor-string)|指定した色の列に "セルの色" フィルターを適用します。|1.2|
|[filter](../excel/filter.md)|_メソッド_ > [applyCustomFilter(criteria1: string, criteria2: string, oper: string)](../excel/filter.md#applycustomfiltercriteria1-string-criteria2-string-oper-string)|指定した条件の文字列の列に "アイコン" フィルターを適用します。|1.2|
|[filter](../excel/filter.md)|_メソッド_ > [applyDynamicFilter(criteria: string)](../excel/filter.md#applydynamicfiltercriteria-string)|列に "動的" フィルターを適用します。|1.2|
|[filter](../excel/filter.md)|_メソッド_ > [applyFontColorFilter(color: string)](../excel/filter.md#applyfontcolorfiltercolor-string)|指定した色の列に "フォントの色" フィルターを適用します。|1.2|
|[filter](../excel/filter.md)|_メソッド_ > [applyIconFilter(icon:Icon)](../excel/filter.md#applyiconfiltericon-icon)|指定したアイコンの列に "アイコン" フィルターを適用します。|1.2|
|[filter](../excel/filter.md)|_メソッド_ > [applyTopItemsFilter(count: number)](../excel/filter.md#applytopitemsfiltercount-number)|指定した数の要素の列に "上位アイテム" フィルターを適用します。|1.2|
|[filter](../excel/filter.md)|_メソッド_ > [applyTopPercentFilter(percent: number)](../excel/filter.md#applytoppercentfilterpercent-number)|指定した割合の要素の列に "上位パーセント" フィルターを適用します。|1.2|
|[filter](../excel/filter.md)|_メソッド_ > [applyValuesFilter(values: ()[])](../excel/filter.md#applyvaluesfiltervalues-)|指定した値の列に "値" フィルターを適用します。|1.2|
|[filter](../excel/filter.md)|_メソッド_ > [clear()](../excel/filter.md#clear)|指定した列のフィルターをクリアします。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_プロパティ_ > color|セルをフィルター処理するために使用する HTML カラー文字列。「CellColor」フィルターおよび「fontColor」フィルターと併用します。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_プロパティ_ > criterion1|データをフィルター処理するために使用する最初の条件。「カスタム」フィルター処理の場合には、演算子として使用されます。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_プロパティ_ > criterion2|データをフィルター処理するために使用する 2 番目の条件。「カスタム」フィルター処理の場合には、演算子としてのみ使用されます。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_プロパティ_ >dynamicCriteria|この列に適用する Excel.DynamicFilterCriteria の動的条件。「動的」フィルター処理で使用します。使用可能な値は次のいずれかです。Unknown、AboveAverage、AllDatesInPeriodApril、AllDatesInPeriodAugust、AllDatesInPeriodDecember、AllDatesInPeriodFebruray、AllDatesInPeriodJanuary、AllDatesInPeriodJuly、AllDatesInPeriodJune、AllDatesInPeriodMarch、AllDatesInPeriodMay、AllDatesInPeriodNovember、AllDatesInPeriodOctober、AllDatesInPeriodQuarter1、AllDatesInPeriodQuarter2、AllDatesInPeriodQuarter3、AllDatesInPeriodQuarter4、AllDatesInPeriodSeptember、BelowAverage、LastMonth、LastQuarter、LastWeek、LastYear、NextMonth、NextQuarter、NextWeek、NextYear、ThisMonth、ThisQuarter、ThisWeek、ThisYear、Today、Tomorrow、YearToDate、Yesterday。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_プロパティ_ > filterOn|値を表示したままにするかどうかを判別するために、フィルターで使用するプロパティ。使用可能な値は次のとおりです。BottomItems、BottomPercent、CellColor、Dynamic、FontColor、Values、TopItems、TopPercent、Icon、Custom。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_プロパティ_ > operator|"カスタム" フィルター処理を使用するときに、条件 1 と条件 2 と結合との使用する演算子。使用可能な値は次のとおりです。And、Or。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_プロパティ_ > values|"値" フィルター処理の一部として使用する値のセット。|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_リレーションシップ_ > icon|セルをフィルター処理するために使用するアイコン。「アイコン」フィルター処理で使用します。|1.2|
|[filterDatetime](../excel/filterdatetime.md)|_プロパティ_ > date|データのフィルター処理に使用する ISO8601 形式の日付です。|1.2|
|[filterDatetime](../excel/filterdatetime.md)|_プロパティ_ > specificity|データを保持するのに、日付をどの程度詳細に使用するか。たとえば、date が 2005-04-02 で "month" に設定した場合、フィルター操作では 2005 年 4 月の日付データを含むすべての行が保持されます。使用可能な値は次のとおりです。Year、Month、Day、Hour、Minute、Second。|1.2|
|[formatProtection](../excel/formatprotection.md)|_プロパティ_ > formulaHidden|Excel が範囲内のセルの数式を非表示にするかどうかを示します。null 値は、範囲全体に一様な数式非表示設定がないことを表します。|1.2|
|[formatProtection](../excel/formatprotection.md)|_プロパティ_ > locked|Excel がオブジェクト内のセルをロックするかどうかを示します。null 値は、範囲全体に一様なロック設定がないことを表します。|1.2|
|[icon](../excel/icon.md)|_プロパティ_ > index|指定したセット内のアイコンのインデックスを表します。|1.2|
|[icon](../excel/icon.md)|_プロパティ_ > set|アイコンがその一部であるセットを表します。使用可能な値は次のとおりです。Invalid、ThreeArrows、ThreeArrowsGray、ThreeFlags、ThreeTrafficLights1、ThreeTrafficLights2、ThreeSigns、ThreeSymbols、ThreeSymbols2、FourArrows、FourArrowsGray、FourRedToBlack、FourRating、FourTrafficLights、FiveArrows、FiveArrowsGray、FiveRating、FiveQuarters、ThreeStars、ThreeTriangles、FiveBoxes。|1.2|
|[range](../excel/range.md)|_プロパティ_ > columnHidden|現在の範囲のすべての列が非表示になっているかどうかを表します。|1.2|
|[range](../excel/range.md)|_プロパティ_ > formulasR1C1|R1C1 スタイル表記の数式を表します。|1.2|
|[range](../excel/range.md)|_プロパティ_ > hidden|現在の範囲のすべてのセルが非表示になっているかどうかを表します。読み取り専用です。|1.2|
|[range](../excel/range.md)|_プロパティ_ > rowHidden|現在の範囲のすべての行が非表示になっているかどうかを表します。|1.2|
|[range](../excel/range.md)|_リレーションシップ_ > sort|現在の範囲について、範囲の並べ替えを表します。読み取り専用。|1.2|
|[range](../excel/range.md)|_メソッド_ > [merge(across: bool)](../excel/range.md#mergeacross-bool)|範囲内のセルをワークシートの 1 つの領域に結合します。|1.2|
|[range](../excel/range.md)|_メソッド_ > [unmerge()](../excel/range.md#unmerge)|範囲内のセルを結合解除して別々のセルにします。|1.2|
|[rangeFormat](../excel/rangeformat.md)|_プロパティ_ > columnWidth|範囲内のすべての列の幅を取得または設定します。列の幅が均一でない場合は、null が返されます。|1.2|
|[rangeFormat](../excel/rangeformat.md)|_プロパティ_ > rowHeight|範囲内のすべての行の高さを取得または設定します。行の高さが均一でない場合は、null が返されます。|1.2|
|[rangeFormat](../excel/rangeformat.md)|_リレーションシップ_ > protection|範囲に対する書式保護オブジェクトを返します。読み取り専用です。|1.2|
|[rangeFormat](../excel/rangeformat.md)|_メソッド_ > [autofitColumns()](../excel/rangeformat.md#autofitcolumns)|現在の列のデータに基づいて、現在の範囲の列の幅を最適な幅に変更します。|1.2|
|[rangeFormat](../excel/rangeformat.md)|_メソッド_ > [autofitRows()](../excel/rangeformat.md#autofitrows)|現在の行のデータに基づいて、現在の範囲の行の高さを最適な高さに変更します。|1.2|
|[rangeReference](../excel/rangereference.md)|_プロパティ_ > address|現在の範囲の表示されている行を表します。|1.2|
|[rangeSort](../excel/rangesort.md)|_メソッド_ > [apply(fields:SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)](../excel/rangesort.md#applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string)|並べ替え操作を実行します。|1.2|
|[sortField](../excel/sortfield.md)|_プロパティ_ > ascending|昇順で並べ替えるかどうかを表します。|1.2|
|[sortField](../excel/sortfield.md)|_プロパティ_ > color|並べ替えがフォントまたはセルの色で行われる場合に、条件の対象となる色を表します。|1.2|
|[sortField](../excel/sortfield.md)|_プロパティ_ > dataOption|このフィールドのその他の並べ替えオプションを表します。使用可能な値は次のとおりです。Normal、TextAsNumber。|1.2|
|[sortField](../excel/sortfield.md)|_プロパティ_ > key|条件の対象とする列 (または行。並べ替えの方向によって異なります) を表します。最初の列 (または行) からのオフセットとして表します。|1.2|
|[sortField](../excel/sortfield.md)|_プロパティ_ > sortOn|この条件の並べ替えの種類を表します。使用可能な値は次のとおりです。Value、CellColor、FontColor、Icon。|1.2|
|[sortField](../excel/sortfield.md)|_リレーションシップ_ > icon|並べ替えがセルのアイコンで行われる場合に、条件の対象となるアイコンを表します。|1.2|
|[table](../excel/table.md)|_リレーションシップ_ > sort|テーブル内の並べ替えを表します。読み取り専用。|1.2|
|[table](../excel/table.md)|_リレーションシップ_ > worksheet|現在のテーブルを含んでいるワークシート。読み取り専用です。|1.2|
|[table](../excel/table.md)|_メソッド_ > [clearFilters()](../excel/table.md#clearfilters)|現在テーブルに適用されているすべてのフィルターをクリアします。|1.2|
|[table](../excel/table.md)|_メソッド_ > [convertToRange()](../excel/table.md#converttorange)|テーブルを通常の範囲のセルに変換します。すべてのデータが保持されます。|1.2|
|[table](../excel/table.md)|_メソッド_ > [reapplyFilters()](../excel/table.md#reapplyfilters)|現在テーブルに適用されているすべてのフィルターを再適用します。|1.2|
|[tableColumn](../excel/tablecolumn.md)|_リレーションシップ_ > filter|列に適用されるフィルターを取得します。読み取り専用です。|1.2|
|[tableSort](../excel/tablesort.md)|_プロパティ_ > matchCase|大文字小文字の区別が、テーブルの最後の並べ替え操作に影響を与えたかどうかを表します。読み取り専用です。|1.2|
|[tableSort](../excel/tablesort.md)|_プロパティ_ > method|テーブルの並べ替えで最後に使用した中国語文字の順序付け方法を表します。読み取り専用です。使用可能な値は次のとおりです。PinYin、StrokeCount。|1.2|
|[tableSort](../excel/tablesort.md)|_リレーションシップ_ > fields|テーブルの最後の並べ替えに使用する現在の条件を表します。読み取り専用です。|1.2|
|[tableSort](../excel/tablesort.md)|_メソッド_ > [apply(fields:SortField[], matchCase: bool, method: string)](../excel/tablesort.md#applyfields-sortfield-matchcase-bool-method-string)|並べ替え操作を実行します。|1.2|
|[tableSort](../excel/tablesort.md)|_メソッド_ > [clear()](../excel/tablesort.md#clear)|テーブルに現在設定されている並べ替えをクリアします。これにより表の順序が変更されることはありませんが、ヘッダーのボタンの状態がクリアされます。|1.2|
|[tableSort](../excel/tablesort.md)|_メソッド_ > [reapply()](../excel/tablesort.md#reapply)|テーブルに、現在の並べ替えパラメーターを再適用します。|1.2|
|[workbook](../excel/workbook.md)|_リレーションシップ_ > functions|このブックを含む Excel アプリケーションのインスタンスを表します。読み取り専用。|1.2|
|[worksheet](../excel/worksheet.md)|_リレーションシップ_ > protection|ワークシートのシート保護オブジェクトを返します。読み取り専用です。|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_プロパティ_ > protected|ワークシートが保護されているかどうかを示します。読み取り専用。読み取り専用。|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_リレーションシップ_ > options|シートの保護のオプション。読み取り専用。|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_メソッド_ > [protect(options:WorksheetProtectionOptions)](../excel/worksheetprotection.md#protectoptions-worksheetprotectionoptions)|ワークシートを保護します。ワークシートが保護されている場合は失敗します。|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_メソッド_ > [unprotect()](../excel/worksheetprotection.md#unprotect)|ワークシートの保護を解除します。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_プロパティ_ > allowAutoFilter|自動フィルター機能の使用を可能にするワークシート保護オプションを表します。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_プロパティ_ > allowDeleteColumns|列の削除を可能にするワークシート保護オプションを表します。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_プロパティ_ > allowDeleteRows|行の削除を可能にするワークシート保護オプションを表します。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_プロパティ_ > allowFormatCells|セルの書式設定を可能にするワークシート保護オプションを表します。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_プロパティ_ > allowFormatColumns|列の書式設定を可能にするワークシート保護オプションを表します。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_プロパティ_ > allowFormatRows|行の書式設定を可能にするワークシート保護オプションを表します。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_プロパティ_ > allowInsertColumns|列の挿入を可能にするワークシート保護オプションを表します。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_プロパティ_ > allowInsertHyperlinks|ハイパーリンクの挿入を可能にするワークシート保護オプションを表します。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_プロパティ_ > allowInsertRows|行の挿入を可能にするワークシート保護オプションを表します。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_プロパティ_ > allowPivotTables|ピボットテーブル機能の使用を可能にするワークシート保護オプションを表します。|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_プロパティ_ > allowSort|並ベ替え機能の使用を可能にするワークシート保護オプションを表します。|1.2|

## <a name="excel-javascript-api-11"></a>Excel JavaScript API 1.1
Excel JavaScript API 1.1 は、API の最初のバージョンです。API について詳しくは、Excel JavaScript API リファレンスのトピックをご覧ください。  

## <a name="additional-resources"></a>追加リソース

- [Office のホストと API の要件を指定する](../../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../docs/overview/add-in-manifests.md)
