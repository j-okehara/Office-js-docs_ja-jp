# <a name="breaktype-javascript-api-for-word"></a>BreakType (JavaScript API for Word)

区切りの形式を指定します。

_適用対象:Word 2016、Word for iPad、Word for Mac、Word Online_

API でサポートされている区切りの種類を次に示します。

| **値**         | **型** | **説明**     |
|:-----------------|:--------|:----|
|line| | 改行します。 |
|page| | 挿入位置で改ページします。|
|sectionNext| | 次のページ上にセクション区切りを挿入します。次の種類は使われなくなります。|
|sectionContinuous| | 改ページなしで新しいセクションを開始します。|
|sectionEven| string | セクション区切りを挿入し、次の偶数ページから次のセクションを開始します。セクション区切りを偶数ページに挿入した場合、次の奇数ページは空白になります。|
|sectionOdd| string | セクション区切りを挿入し、次の奇数ページから次のセクションを開始します。セクション区切りを奇数ページに挿入した場合、次の偶数ページは空白になります。|

## <a name="support-details"></a>サポートの詳細
実行時のチェックで[要件セット](../office-add-in-requirement-sets.md)を使用して、アプリケーションが Word のホスト バージョンによってサポートされていることを確かめます。Office ホスト アプリケーションとサーバーの要件の詳細については、「[Office アドインを実行するための要件](../../docs/overview/requirements-for-running-office-add-ins.md)」を参照してください。
