

# diagnostics

## [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics

Outlook アドインに診断情報を提供します。

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|

### メンバー

####  hostName :String

ホスト アプリケーションの名前を表す文字列を取得します。

文字列は、値 `Outlook`、`Mac Outlook`、または `OutlookWebApp` のいずれかになります。

##### 型:

*   String

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|
####  hostVersion :String

ホスト アプリケーションまたは Exchange Server のバージョンを表す文字列を取得します。

メール アドインを Outlook デスクトップ クライアントで実行している場合、`hostVersion` プロパティはホスト アプリケーションである Outlook のバージョンを返します。Outlook Web App では、このプロパティは Exchange Server のバージョンを返します。たとえば、文字列 `15.0.468.0` などです。

##### 型:

*   String

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|
####  OWAView :String

Outlook Web App の現在のビューを表す文字列を取得します。

返される文字列は、値 `OneColumn`、`TwoColumns`、または `ThreeColumns` のいずれかになります。

ホスト アプリケーションが Outlook Web App ではない場合、このプロパティにアクセスすると `undefined` が返されます。

Outlook Web App には、画面とウィンドウの幅、および表示可能な列数に応じて 3 つのビューがあります。

*   画面幅が狭い場合に表示される `OneColumn`。Outlook Web App は、この単一列レイアウトを使用してスマートフォンの画面全体への表示を行います。
*   画面幅がやや広い場合に表示される `TwoColumns`。Outlook Web App は、ほとんどのタブレットでこのビューを使用します。
*   画面幅が広い場合に表示される `ThreeColumns`。Outlook Web App は、デスクトップ コンピューターのフル スクリーン ウィンドウなどでこのビューを使用します。

##### 型:

*   String

##### 要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../tutorial-api-requirement-sets.md)| 1.0|
|[最小のアクセス許可レベル](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|適用可能な Outlook のモード| 作成または読み取り|
