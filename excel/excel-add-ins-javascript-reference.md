# Excel アドインの JavaScript API リファレンス

_適用対象: Excel 2016、Office 2016_

以下のリンクは、API で使用できるExcel オブジェクトの概要を示しています。オブジェクトのページの各リンクには、オブジェクトで使用できるプロパティ、リレーションシップ、メソッドの説明が含まれています。リンクを探索して詳細をご確認ください。
	
* [ブック](resources/workbook.md): ワークシート、テーブル、範囲などの関連するブック オブジェクトを含む最上位オブジェクトです。関連する参照情報を一覧表示するためにも使用されます。 
* [ワークシート](resources/worksheet.md): ワークシートのコレクションのメンバーです。ワークシートのコレクションには、ブック内のすべてのワークシート オブジェクトが含まれています。
	* [ワークシートのコレクション](resources/worksheetcollection.md): ブックの一部であるすべてのブック オブジ ェクトのコレクション。 
* [範囲](resources/range.md): 1 つのセル、1 つの行、または 1 つの列を表すか、あるいは、1 つ以上の連続したセル範囲を含むセルの選択範囲を表します。  
* [テーブル](resources/table.md): データの管理が簡単になるように設計された、体系化されたセルのコレクションを表します。 
	* [テーブルのコレクション](resources/tablecollection.md): ブックまたはワークシート内のテーブルのコレクションです。 
	* [TableColumn コレクション](resources/tablecolumncollection.md): テーブル内のすべての列のコレクションです。 
	* [TableRow Collection](resources/tablerowcollection.md): テーブル内のすべての行のコレクションです。 
* [グラフ](resources/chart.md): 基になるデータを視覚的に表示する、ワークシート内の Chart オブジェクトを表します。   
	* [グラフのコレクション](resources/chartcollection.md): ワークシート内のグラフのコレクションです。	
* [NamedItem](resources/nameditem.md): セルまたは値の範囲の定義済みの名前を表します。名前には、プリミティブ型の名前付きオブジェクト、range オブジェクトなどを指定できます。
	* [NamedItem コレクション](resources/nameditemcollection.md): ブック内の NamedItem オブジェクトのコレクション。
* [バインド](resources/binding.md): ブックのセクションへのバインドを表す抽象クラス。
	* [バインド コレクション](resources/bindingcollection.md): ブックの一部であるすべてのバインド オブジェクトのコレクション。 
* [TrackedObject コレクション](resources/trackedobjectscollection.md): アドインが sync() バッチ間で範囲オブジェクトの参照を管理できるようにします。 
* [要求のコンテキスト](resources/requestcontext.md): RequestContext オブジェクトは、Excel アプリケーションへの要求を容易にします。


##### その他の技術情報

*  [Excel アドインのプログラミングの概要](excel-add-ins-programming-overview.md)
*  [最初の Excel アドインをビルドする](build-your-first-excel-add-in.md)
*  [Excel のスニペット エクスプローラー](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
*  [Excel アドインのコード サンプル](excel-add-ins-code-samples.md) 


