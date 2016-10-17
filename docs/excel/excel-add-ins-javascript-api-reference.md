# <a name="excel-javascript-api-reference"></a>Excel の JavaScript API リファレンス

Excel の JavaScript API を使用して、Excel 2016 用アドインをビルドします。API で使用できる Excel オブジェクトの概要を次に示します。オブジェクトのページの各リンクには、オブジェクトで使用できるプロパティ、リレーションシップ、メソッドの説明が含まれています。リンクを調べて詳細を確認してください。

* [ブック](../../reference/excel/workbook.md): ワークシート、テーブル、範囲などの関連するブック オブジェクトを含む最上位オブジェクトです。関連する参照情報を一覧表示するためにも使用されます。
* [ワークシート](../../reference/excel/worksheet.md): ワークシートのコレクションのメンバーです。ワークシートのコレクションには、ブック内のすべてのワークシート オブジェクトが含まれています。
    * [ワークシートのコレクション](../../reference/excel/worksheetcollection.md): ブックの一部であるすべてのブック オブジ ェクトのコレクション。
* [範囲](../../reference/excel/range.md): 1 つのセル、1 つの行、または 1 つの列を表すか、あるいは、1 つ以上の連続したセル範囲を含むセルの選択範囲を表します。
* [テーブル](../../reference/excel/table.md): データの管理が簡単になるように設計された、体系化されたセルのコレクションを表します。
    * [テーブルのコレクション](../../reference/excel/tablecollection.md): ブックまたはワークシート内のテーブルのコレクションです。
    * [TableColumn コレクション](../../reference/excel/tablecolumncollection.md): テーブル内のすべての列のコレクションです。
    * [TableRow Collection](../../reference/excel/tablerowcollection.md): テーブル内のすべての行のコレクションです。
* [グラフ](../../reference/excel/chart.md): 基になるデータを視覚的に表示する、ワークシート内の Chart オブジェクトを表します。
    * [グラフ コレクション](../../reference/excel/chartcollection.md):ワークシート内のグラフのコレクションです。
* [TableSort](../../reference/excel/tablesort.md):Table オブジェクトの並べ替え操作を行うオブジェクトを表します。
* [RangeSort](../../reference/excel/rangesort.md):Range オブジェクトの並べ替え操作を行うオブジェクトを表します。
* [Filter](../../reference/excel/filter.md):テーブルの列のフィルター処理を管理するフィルター オブジェクトを表します。
* [ワークシート保護](../../reference/excel/worksheetprotection.md)ワークシート オブジェクトの保護を表します。
* [ワークシート関数](../../reference/excel/functions.md)JavaScript から呼び出すことができる Microsoft Excel ワークシート関数のコンテナーを表します。
* [NamedItem](../../reference/excel/nameditem.md):セルまたは値の範囲の定義済みの名前を表します。名前には、プリミティブ型の名前付きオブジェクト、range オブジェクトなどを指定できます。
    * [NamedItem コレクション](../../reference/excel/nameditemcollection.md): ブック内の NamedItem オブジェクトのコレクション。
* [バインド](../../reference/excel/binding.md): ブックのセクションへのバインドを表す抽象クラス。
    * [バインド コレクション](../../reference/excel/bindingcollection.md): ブックの一部であるすべてのバインド オブジェクトのコレクション。
* [TrackedObject コレクション](../../reference/excel/trackedobjectscollection.md): アドインが sync() バッチ間で範囲オブジェクトの参照を管理できるようにします。
* [要求のコンテキスト](../../reference/excel/requestcontext.md): RequestContext オブジェクトは、Excel アプリケーションへの要求を容易にします。


##### <a name="additional-resources"></a>その他のリソース

*  [Excel アドインのプログラミングの概要](excel-add-ins-javascript-programming-overview.md)
*  [最初の Excel アドインをビルドする](build-your-first-excel-add-in.md)
*  [Excel のスニペット エクスプローラー](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)

