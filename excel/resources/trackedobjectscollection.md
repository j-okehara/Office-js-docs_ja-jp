# TrackedObjectsCollection オブジェクト (JavaScript API for Office 2016)

_適用対象: Excel 2016、Excel Online、Office 2016_

アドインが sync() バッチ間で範囲オブジェクトの参照を管理できるようにします。通常、Excel.run() では、明示的に追跡しなくても自動的にバッチ間で参照を維持できます。しかし、アドインのシナリオで範囲オブジェクトを手動で追跡および調整して基になる Excel 範囲の現在の状態を反映する必要がある場合には、このコレクションを使用してそのようなオブジェクトに追跡のマークを付けることができます。範囲オブジェクトに追跡のマークが付けられている場合は、エラーの場合でも、Excel でメモリを解放するために使用時に明示的に削除する必要があります。

## プロパティ
なし

## 関係

なし

## メソッド

trackedObjectsCollection オブジェクトは次の定義されたメソッドを持ちます:

| メソッド     | 戻り値の型    |説明|
|:-----------------|:--------|:----------|
|[add(rangeObject:Range)](#addrangeobject-range)| Null             |範囲の新しい参照を作成します。|
|[remove(rangeObject:Range)](#removerangeobject-range)| Null             |範囲の参照を削除します。  |
|[removeAll()](#removeall)| Null|デバイス上のアドインによって作成されたすべての参照を削除します。|


## API 仕様 

### add(rangeObject: range)
trackedObjectsCollection に range オブジェクトを追加します。バッチ要求間での基になる範囲の変更が追跡され、フォロー アップの更新が範囲オブジェクトの現在の状態に適用されます。 

#### 構文
```js
trackedObjectsCollection.add(rangeObject);
```

#### パラメータ

パラメーター       | 型   | 説明
--------------- | ------ | ------------
`rangeObject`  | [Range](range.md)| trackedObjectCollection に追加する必要のある Range オブジェクト。

#### 戻り値
Null

#### 例

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.trackedObjectsCollection.add(range);
ctx.load(range);

Excel.run(function (ctx) { 
	range.insert("Down");
	Console.log(range.address); // Address should be updated to A3:B4
	return ctx.sync(); 
});
```


### remove(rangeObject: range)

参照オブジェクトをコレクションから削除します。これはメモリおよび追跡対象のオブジェクトの状態を維持するために必要なリソースを開放します。範囲オブジェクトに追跡のマークが付けられている場合は、エラーの場合でも明示的に削除することが必要です。

#### 構文
```js
trackedObjectsCollection.remove(rangeObject);
```

#### パラメータ

パラメーター       | 型   | 説明
--------------- | ------ | ------------
`rangeObject`  | [Range](range.md)| trackedObjectCollection から削除する必要のある Range オブジェクト。

#### 戻り値
Null

#### 例


```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.trackedObjectsCollection.add(range);
ctx.load(range);

Excel.run(function (ctx) { 
	range.insert("Down");
	Console.log(range.address); // Address should be updated to A3:B4
	ctx.trackedObjectsCollection.remove(range); 
	return ctx.sync(); 
});
```

### removeAll(rangeObject: range)

デバイス上のアドインによって作成されたすべての参照を削除します。

#### 構文
```js
trackedObjectsCollection.removeAll();
```

#### パラメーター

なし

#### 戻り値
Null

#### 例

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:B2";
	var ctx = new Excel.RequestContext();
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	ctx.trackedObjectsCollection.add(range);
	ctx.load(range);
	range.insert("Down");
	Console.log(range.address); // Address should be updated to A3:B4
	ctx.trackedObjectsCollection.removeAll(); 
	return ctx.sync(); 
});
```

