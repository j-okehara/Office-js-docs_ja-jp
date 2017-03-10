# <a name="bindingselectionchangedeventargs-object-javascript-api-for-excel"></a>BindingSelectionChangedEventArgs オブジェクト (JavaScript API for Excel))

SelectionChanged イベントが発生したバインドに関する情報を提供します。

## <a name="properties"></a>プロパティ

| プロパティ       | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|columnCount|int|選択されている列の数を取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|rowCount|int|選択されている行の数を取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|startColumn|int|選択範囲の先頭列のインデックス (0 から始まる) を取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|startRow|int|選択範囲の先頭行のインデックス (0 から始まる) を取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

_プロパティのアクセスの[例を参照してください。](#property-access-examples)_

## <a name="relationships"></a>関係
| リレーションシップ | 型    |説明| 要件セット|
|:---------------|:--------|:----------|:----|
|binding|[Binding](binding.md)|SelectionChanged イベントが発生したバインドを表す Binding オブジェクトを取得します。|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## <a name="methods"></a>メソッド
なし

