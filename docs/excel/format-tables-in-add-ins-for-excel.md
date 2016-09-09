
# Excel 用アドインで表の書式設定を行う


この記事では、書式設定 API のさまざまな機能を説明し、それらの使用方法について概説します。 このリリースでは、(**Office.CoercionType.Text** または **Office.CoercionType.Matrix** データ構造ではなく) テーブルのみにセルの書式設定とその他のオプションをプログラムで指定でき、また Excel アドインでのみこれを行うことができます。 次のようにして、アドインで書式を設定します。

- ユーザーがテーブル (またはプログラムでテーブルを挿入する場所) を選択します。その後、書式を設定するテーブルに対してアドインで **Document.setSelectedDataAsync** メソッドを呼び出すことができます。

- または、バインドされたテーブルがブックに既に含まれている場合 (またはアドインの初期化時にアドインが [Bindings](../../reference/shared/bindings.bindings.md) オブジェクトの "addFrom" メソッドのいずれかを使用して、バインドされたテーブルを作成する場合)、アドインは、バインドされたテーブルで書式設定を行うために **Binding.setDataAsync** メソッドを呼び出すことができます。
    
>**重要:** Excel アドインで表の書式設定をするための新しいメソッドおよび更新されたメソッドを使用するには、アドインのプロジェクトで [Office.js v1.1 以上を使用するか、それを使用するようにプロジェクトを更新する](../../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md)必要があります。

## 書式を指定する

設定する書式を指定するには、キーと値のペアを 1 つ以上含む JavaScript オブジェクト リテラルを作成します。一連の書式設定を JavaScript オブジェクト内のリストにまとめることができます。以下に例を示します。 


```js
var myFormat = {fontStyle:"bold", width:"autoFit", borderColor:"purple"};
```

書式設定を適用するために、JavaScript オブジェクトをデータの書式設定やテーブルのその他の機能をサポートするメソッドのいずれかに渡します。

書式設定は次の 2 つの方法で操作できます。


- アドインは、初めてデータを選択範囲またはバインドに書き込む際に、[Document.setSelectedDataAysnc](../../reference/shared/document.setselecteddataasync.md) メソッドまたは [Binding.setDataAsync](../../reference/shared/binding.setdataasync.md) メソッドに渡される _options_ オブジェクトにあるオプションの _cellFormat_ パラメーターまたは _tableOptions_ パラメーターを指定します。
    
- 書式設定した後は、[書式設定をクリアまたは更新する](#書式設定をクリアまたは更新する)ための専用の新しいメソッドのいずれかを使用して、書式設定をクリアまたは更新できます。
    

## データ設定メソッドでオプションのパラメーターを使用する

テーブル バインディングでは、_tableOptions_ と _cellFormat_ というオプション パラメーターを使用して **Document.setSelectedData** メソッドと **Binding.setDataAsync** メソッドのどちらかを使用してデータを設定するときに書式設定を指定できます。


### tableOptions オプション パラメーター

_tableOptions_ オプション パラメーターは、既定のテーブル スタイルを指定し、**見出し行**、**集計行**、**縞模様行**など、特定のテーブル機能をオンまたはオフにするために使用します。 _tableOptions_ パラメーターとして渡す値はキーと値のペアのリストを含む JavaScript オブジェクトです。 次に例を示します。


```js
tableOptions: {bandedRows: true, filterButton: false, style:"TableStyleMedium3"};
```


### cellFormat オプション パラメーター

_cellFormat_ オプション パラメーターは、幅、高さ、フォント、背景、配置など、セルの書式設定値を変更するために使用します。 _cellFormat_ パラメーターとして渡す値は、対象とするセルと、それらのセルに適用する書式設定を指定する JavaScript オブジェクトのリストを含む配列です。 次に例を示します。


```js
cellFormat: 
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: Office.Table.Headers, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}]
```

複数の `cells:` と `format:` のペアを _cellFormat_ 配列にまとめることで、書式設定の適用に必要な関数呼び出しの回数を最低限に抑えることができます。


#### cells

Use  `cells:` to specify the range of columns, rows, and cells you want to apply formatting to.


**サポートされている cells 値の範囲**


|**cells の範囲の設定**|**説明**|
|:-----|:-----|
| `{row: i}`|テーブル内の i データ行までの範囲を指定します。|
| `{column: i}`|テーブル内の i データ列までのセルの範囲を指定します。|
| `{row: i, column: j}`|テーブル内の i データ行から j データ列までのセルの範囲を指定します。|
| `Office.Table.All`|列見出し、データ、集計 (もしあれば) を含むテーブル全体を指定します。|
| `Office.Table.Data`|テーブル内のデータのみ (見出しと集計を含まない) を指定します。|
| `Office.Table.Headers`|見出し行のみを指定します。|

#### format

`format:` は、JavaScript のキー/値ペアのリストとして `cells:` で定義した範囲に適用する書式設定を指定するために使用します。 サポートされる値のリストについては、「[サポートされる書式設定のキーと値](#サポートされる書式設定のキーと値)」を参照してください。

 **Excel Online での書式設定の指定の制限**

Excel Online で書式を設定する場合、_cellFormat_ パラメーターに渡される_書式設定グループ_の数が 100 を超えることはできません。 1 つの書式設定グループは、指定のセル範囲に適用される書式設定のセットから構成されます  (つまり、配列内の `cells:`オブジェクトのリテラルの 1 つで指定したものすべてが _cellFormat_ に渡されるということです)。たとえば、次の呼び出しは、2 つの書式設定グループを _cellFormat_ に渡します。




```js
Office.context.document.setSelectedDataAsync(
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});
```


#### オプション パラメーターの適用

このリリースでは、同一呼び出し内で _tableOptions_ および _cellFormat_ オプション パラメーターを使用してテーブルに対するデータの書き込みと書式の設定をサポートするメソッドは、**Document.setSelectedDataAsync** と **TableBinding.setDataAsync** のみです。 次に示す例では、各メソッドの最初のパラメーター (_data_ パラメーター) に渡される `tableData` 値は、書き込むテーブルとデータの定義が格納されている [TableData](../../reference/shared/tabledata.md) オブジェクトであることが必要です。

 **Document.setSelectedDataAsync の例**




```js
Office.context.document.setSelectedDataAsync(tableData, 
    {tableOptions: {headerRow:false}, 
        cellFormat: [{cells: Office.Table.Headers, format: {fontColor: "yellow"}}]}, 
    function (asyncResult) {});
```

 **TableBinding.setDataAsync example**




```js
Office.select("bindings#myBinding").setDataAsync(tableData, 
    {tableOptions: {headerRow:false}, 
        cellFormat: [{cells: Office.Table.Headers, format: {fontColor: "yellow"}}]}, 
    function (asyncResult) {});
```

 >**注:**:`Office.select("bindings#myBinding")` の呼び出しは、既に `myBinding` という名前のバインドがワークシートに存在していると仮定しています。


## 書式設定の更新と解除


**Document.setSelectedDataAsync** または **TableBinding.setDataAsync** メソッドの _cellFormat_ および _tableOptions_ オプション パラメーターで書式を設定する場合は、メソッドの最初の呼び出し時にのみ書式が設定されます。 書式を更新または解除するには、**TableBinding** オブジェクトの新しい 3 つのメソッド (**setFormatsAsync**、**setTableOptionsAsync**、および **clearFormatsAsync**) を使用する必要があります。


### 書式設定の更新

[TableBinding.setFormatsAsync](../../reference/shared/binding.tablebinding.setformatsasync.md) メソッドは、セルの書式設定 (幅、高さ、フォント、背景、配置など) の更新専用です。 このメソッドは、必須パラメーターとして _cellFormat_ を受け取ります。


```js
Office.select("bindings#myBinding").setFormatsAsync(
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}], 
    function (asyncResult){});
```

[TableBinding.setTableOptionsAsync](../../reference/shared/binding.tablebinding.settableoptionsasync.md) メソッドは、テーブル オプション (縞模様の行やフィルター ボタンなど) の更新専用です。 このメソッドは、必須パラメーターとして _tableOptions_ を受け取ります。




```js
var tableOptions = {bandedRows: true, filterButton: false, style: "TableStyleMedium3"}; 

Office.select("bindings#myBinding").setTableOptionsAsync(tableOptions, function(asyncResult){});
```


### 書式設定の解除

[TableBinding.clearFormatsAsync](../../reference/shared/binding.tablebinding.clearformatsasync.md) メソッドは、テーブル内のすべての書式設定を解除します。 このメソッドは、_asyncContext_ オプション パラメーターとオプションのコールバック関数を受け取ります。


```js
Office.select("bindings#myBinding").clearFormatsAsync();
```


## サポートされる書式設定のキーと値


次の表は、サポートされるキー/値ペアの一覧です。これらのペアは、_cellFormat_ または _tableOptions_ パラメーターに渡すことができます。

`format:` の値については、[**セルの書式設定**] ダイアログ ボックス (右クリック > [**セルの書式設定**] またはリボンの [**ホーム**] タブ > [**書式設定**] > [**セルの書式設定**]) の設定のサブセットに対応する設定値が使用できます。 `tableOptions:` の値については、リボンの [**表ツール**] |[**デザイン**] タブにある [**テーブル スタイルのオプション**] と [**テーブル スタイル**] グループの設定値に対応します。


 >**重要**: 書式設定 API のメソッドは、以下にまとめたオプションと値のみをサポートします。 その他の書式設定オプションまたは値を指定する場合、処理動作は定義されていません。 このような未定義の処理の動作は、サポートされているプラットフォーム間で整合性が保証されているとは限りません。そのため、特定のプラットフォームで、この未定義動作の副次的な影響に基づくアドインを開発してはいけません。 ただし、未定義処理の動作は、アドインの状態および UI、または操作対象のドキュメントに悪影響を与えることはありません。


**位置揃え**


|**キー**|**値**|**メモ**|
|:-----|:-----|:-----|
| `alignHorizontal:`|"general" \| "left" \| "center" \| "right" \| "fill" \| "justify" \| "center across selection" \| "distributed"|When combined with an indent value, only the following combinations are supported:<br/><br/><ul><li><code>alignHorizontal: "left"</code> and <code>indentLeft: \<value\></code></li></ul><ul><li><code>alignHorizontal: "right"</code> and <code>indentRight: \<value\></code></li></ul><ul><li><code>alignHorizontal: "distributed"</code> and <code>indentDistributed: \<value\></code></li></ul>|
| `alignVertical:`|"top" \| "center" \| "bottom" \| "justify" \| "distributed"||



**背景**


|**キー**|**値**|**メモ**|
|:-----|:-----|:-----|
| `backgroundColor:`|"none" \| \<All predefined color names\> \| #RRGGBB|Predefined color names:<br/><br/>備考|



**Border**


|**キー**|**値**|**メモ**|
|:-----|:-----|:-----|
| `borderStyle:`|"none" \| \<All predefined border style names\>|Predefined border style names:<br/><br/>xlonline<br/><br/>指定した範囲のすべての罫線に適用されます  ([**セルの書式設定**] ダイアログ ボックスの [**罫線**] タブで [**外枠**] と [**内側**] プリセットの両方を使用して罫線スタイルを指定することと同じです)。<br/><br/> **注:** Excel 2013 は、13 種すべての定義済み罫線スタイルのレンダリングをサポートしています。 ただし、Excel Online ではサポートされない罫線スタイルがあります。 次の表では、Excel Online でスプレッドシートを開くときに、各枠線スタイルで使用されるレンダリングについて説明します。<br/><br/><table><tr><th>Excel 2013</th><th>Excel Online</th></tr><tr><td>"dash dot"</td><td>dashed (1px)</td></tr><tr><td>"dash dot dot"</td><td>dotted (1px)</td></tr><tr><td>"dashed"</td><td>dotted (1px)</td></tr><tr><td>"dotted"</td><td>dashed (1px)</td></tr><tr><td>"double"</td><td>double (3px)</td></tr><tr><td>"hair"</td><td>solid (1px)</td></tr><tr><td>"medium dash dot"</td><td>dashed (2px)</td></tr><tr><td>"medium dash dot dot"</td><td>dotted (2px)</td></tr><tr><td>"medium dashed"</td><td>dashed (2px)</td></tr><tr><td>"medium"</td><td>solid (2px)</td></tr><tr><td>"slant dash dot"</td><td>dashed (2px)</td></tr><tr><td>"thick"</td><td>solid (3px)</td></tr><tr><td>"thin"</td><td>solid (1px)</td></tr></table>|
| `borderColor:`|"automatic" \| \<All predefined color names\> \| #RRGGBB|指定した範囲のすべての罫線に適用されます。|
| `borderTopStyle:`|"none" \| \<All predefined border style names\>|指定した範囲のすべての罫線に適用されます。|
| `borderTopColor:`|"automatic" \| \<All predefined color names\> \| #RRGGBB|指定した範囲のすべての罫線に適用されます。|
| `borderBottomStyle:`|"none" \| \<All predefined border style names\>|指定した範囲のすべての罫線に適用されます。|
| `borderBottomColor:`|"automatic" \| \<All predefined color names\> \| #RRGGBB|指定した範囲のすべての罫線に適用されます。|
| `borderLeftStyle:`|"none" \| \<All predefined border style names\>|指定した範囲のすべての罫線に適用されます。|
| `borderLeftColor:`|"automatic" \| \<All predefined color names\> \| #RRGGBB|指定した範囲のすべての罫線に適用されます。|
| `borderRightStyle:`|"none" \| \<All predefined border style names\>|指定した範囲のすべての罫線に適用されます。|
| `borderRightColor:`|"automatic" \| \<All predefined color names\> \| #RRGGBB|指定した範囲のすべての罫線に適用されます。|
| `borderOutlineStyle:`|"none" \| \<All predefined border style names\>|指定した範囲のすべての罫線に適用されます。|
| `borderOutlineColor:`|"automatic" \| \<All predefined color names\> \| #RRGGBB|指定した範囲のすべての罫線に適用されます。|
| `borderInlineStyle:`|"none" \| \<すべての定義済みの罫線スタイルの名前\>|指定された範囲内の内側の罫線にのみ適用されます  (**[セルの書式設定]** ダイアログ ボックスの **[罫線]** タブで **[内側]** プリセットのみを使用して線のスタイルを指定した場合と等価です)。|
| `borderInlineColor:`|"automatic" \| \<All predefined color names\> \| #RRGGBB|指定された範囲内の内側の罫線にのみ適用されます。 |



**Cell width, height and wrapping**


|**キー**|**値**|
|:-----|:-----|
| `width:`|"auto fit" \|  **Number**|
| `height:`|"auto fit" \|  **Number**|
| `wrapping:`|**ブール型 (Boolean)**|



**フォント**


|**キー**|**値**|**注**|
|:-----|:-----|:-----|
| `fontFamily:`|\<使用可能なすべてのフォント名\>|Excel Online でフォントを設定したときに、そのフォントがブラウザーで使用できない場合、API は、次の順序でフォントのフォール バックを試行します。Segoe UI、Thonburi、Arial、Verdana、および Microsoft Sans Serif のフォント。 いずれのフォントも使用できない場合は、ブラウザーの既定のフォントが使用されます。|
| `fontStyle:`|"regular" \| "italic" \| "bold" \| "bold italic"|**注**: この記事の発行時点では、`fontStyle:` を "italic" に設定してから "bold" を設定すると (またはその逆順に設定すると)、それら 2 つの設定が結合されたものとして動作します。 つまり、まず "italic" を設定してから "bold" を設定したとすると、その結果は "bold italic" になるということです。 先に bold か italic のどちらかを設定した範囲に、italic または bold _のみ_を設定するには、まず `fontStyle:"regular"` 設定することで前の書式設定をクリアする必要があります。|
| `fontSize:`|**数値**||
| `fontUnderlineStyle:`|"none" \| "single" \| "double" \| "single accounting" \| "double accounting"||
| `fontColor:`|"automatic" \| \<All predefined color names\> \| #RRGGBB||
| `fontDirection:`|"context" \| "left-to-right" \| "right-to-left"|Excel Online は、現時点では右から左方向へのテキストの表示をサポートしていません。 ただし、Excel Online で実行中のアドインが `fontDirection:` を "right-to-left"に設定している場合、その書式設定はブック ファイルに保存され、デスクトップの Excel でブックで開いたときには正しく表示されます。|
| `fontStrikethrough:`|**ブール型 (Boolean)**||
| `fontSuperscript:`|**ブール型 (Boolean)。**||
| `fontSubScript:`|**ブール型 (Boolean)。**||
| `fontNormal:`|**ブール型 (Boolean)**|フォント、フォント スタイル、サイズ、および文字飾りを標準スタイルに設定します。 これにより、セルのフォントの書式設定が既定の値にリセットされます。 **[セルの書式設定]** ダイアログ ボックスの **[フォント]** タブで、**[標準フォント]** チェック ボックスをオンにすることと同じになります。|



**インデント**


|**キー**|**値**|**メモ**|
|:-----|:-----|:-----|
| `indentLeft:`|**Number**|値<br/><br/><ul><li><code>alignHorizontal: "left"</code> and <code>indentLeft: \<value\></code></li></ul>|
| `indentRight:`|**Number**|値<br/><br/><ul><li><code>alignHorizontal: "right"</code> and <code>indentRight: \<value\></code></li></ul>|
| `indentDistributed:`|**Number**|値<br/><br/><ul><li><code>alignHorizontal: "distributed"</code> and <code>indentDistributed: \<value\></code></li></ul>|



**数値の書式**


|**キー**|**値**|**メモ**|
|:-----|:-----|:-----|
| `numberFormat:`|**String**|[分類] リストから [通貨] などの標準表示形式分類を選択します。<br/><br/> `numberFormat:"#,###.00"`<br/><br/>これは、[[セルの書式設定] ダイアログ ボックスの [表示形式] タグにある [ユーザー定義] の分類で作成できる](http://office.microsoft.com/en-us/excel-help/create-or-delete-a-custom-number-format-HA102749035.aspx?CTT=1)ユーザー定義の数値書式設定文字列と同じです。<br/><br/> **ヒント:**次の手順を使用すれば、Excel の [**セルの書式設定**] ダイアログ ボックスの標準カテゴリの書式文字列がどのように表示されるかを確認できます。<br/><br/><ol><li>[<b>カテゴリ</b>] リストから標準的な書式設定カテゴリ (たとえば、[<span class="ui">通貨</span>]) を選びます。</li><li>Set the format's options in the right side of the dialog box.</li><li>[<b>ユーザー定義</b>] カテゴリを選んで、[<b>型</b>] リストの先頭にある書式文字列を表示します。</li></ol>|



**Table options**


|**キー**|**値**|**メモ**|
|:-----|:-----|:-----|
| `style:`|"none" \| \<All predefined table style names\>|Predefined table style names:<br/><br/>"TableStyleLight1", "TableStyleLight2", "TableStyleLight3", "TableStyleLight4", "TableStyleLight5", "TableStyleLight6", "TableStyleLight7", "TableStyleLight8", "TableStyleLight9", "TableStyleLight10", "TableStyleLight11", "TableStyleLight12", "TableStyleLight13", "TableStyleLight14", "TableStyleLight15", "TableStyleLight16", "TableStyleLight17", "TableStyleLight18", "TableStyleLight19", "TableStyleLight20", "TableStyleLight21", "TableStyleMedium1", "TableStyleMedium2", "TableStyleMedium3", "TableStyleMedium4", "TableStyleMedium5", "TableStyleMedium6", "TableStyleMedium7", "TableStyleMedium8", "TableStyleMedium9", "TableStyleMedium10", "TableStyleMedium11", "TableStyleMedium12", "TableStyleMedium13", "TableStyleMedium14", "TableStyleMedium15", "TableStyleMedium16", "TableStyleMedium17", "TableStyleMedium18", "TableStyleMedium19", "TableStyleMedium20", "TableStyleMedium21", "TableStyleMedium22", "TableStyleMedium23", "TableStyleMedium24", "TableStyleMedium25", "TableStyleMedium26", "TableStyleMedium27", "TableStyleMedium28", "TableStyleDark1", "TableStyleDark2", "TableStyleDark3", "TableStyleDark4", "TableStyleDark5", "TableStyleDark6", "TableStyleDark7", "TableStyleDark8", "TableStyleDark9", "TableStyleDark10", "TableStyleDark11"<br/><br/>テーブル スタイルがどのように表示されるかを確認するには、[**テーブル ツール** \] で Excel にテーブルを挿入します。| [**デザイン**] タブで [**クイック スタイル**] ドロップダウンを選び、定義済みのスタイルを選びます。 スタイルのヒントは、上記リスト内の値のいずれかに対応します。|
| `headerRow:`|**ブール型 (Boolean)**||
| `firstColumn:`|**ブール型 (Boolean)。**||
| `filterButton:`|**ブール型 (Boolean)。**||
| `totalRow:`|**ブール型 (Boolean)。**||
| `lastColumn:`|**ブール型 (Boolean)。**||
| `bandedRows:`|**ブール型 (Boolean)。**||
| `bandedColumns:`|**ブール型 (Boolean)**||
