
# <a name="labsjs.labs.core"></a>LabsJS.Labs.Core
LabsJS.Labs.Core JavaScript API リファレンスの概要について説明します。

 _**適用対象:** Office 用アプリ | Office アドイン | Office Mix | PowerPoint_

LabsJS とプレゼンテーション ドライバー (この場合は、Office Mix) の間で共有され、この 2 者間の橋渡しをする、中心的なインターフェイス、データ構造、クラス。

## <a name="labsjs.labs.core-api-module"></a>LabsJS.Labs.Core API module

Labs.Core モジュールには次の種類が含まれます。


### <a name="classes"></a>クラス


|||
|:-----|:-----|
|[Labs.Core.Permissions](../../reference/office-mix/labs.core.permissions.md)|特定のラボのユーザーに対して有効な権限を表す静的クラス。|

### <a name="interfaces"></a>インターフェイス


|||
|:-----|:-----|
|[Labs.Core.IAction](../../reference/office-mix/labs.core.iaction.md)|指定されたラボでのユーザー相互作用であるラボのアクションを表します。|
|[Labs.Core.IActionResult](../../reference/office-mix/labs.core.iactionresult.md)|アクションの実行結果。アクションが実行されたとき、結果はサーバーによって設定されるか、またはクライアントによって提供されます。いずれになるかはアクションの種類に応じて決まります。|
|[Labs.Core.IComponentInstance](../../reference/office-mix/labs.core.icomponentinstance.md)|ラボ コンポーネントのインスタンスの基底クラス。|
|[Labs.Core.IConfigurationInfo](../../reference/office-mix/labs.core.iconfigurationinfo.md)|ラボ構成に関する情報。|
|[Labs.Core.IConnectionResponse](../../reference/office-mix/labs.core.iconnectionresponse.md)|接続の呼び出しから返される応答情報。|
|[Labs.Core.IGetActionOptions](../../reference/office-mix/labs.core.igetactionoptions.md)|**get** アクションの一貫として渡されるオプション。|
|[Labs.Core.ILabCreationOptions](../../reference/office-mix/labs.core.ilabcreationoptions.md)|ラボ作成の操作の一環として渡されるオプション。|
|[Labs.Core.ILabHostVersionInfo](../../reference/office-mix/labs.core.ilabhostversioninfo.md)|ラボのホストに関するバージョン情報。|
|[Labs.Core.IActionOptions](../../reference/office-mix/labs.core.iactionoptions.md)|ラボ アクションのオプションの定義。指定のアクションを実行するときに渡されるオプション。|
|[Labs.Core.IUserInfo](../../reference/office-mix/labs.core.iuserinfo.md)|ラボに関するユーザー情報を提供します。|
|[Labs.Core.IValueInstance](../../reference/office-mix/labs.core.ivalueinstance.md)|値データを含む場合の [Labs.Core.IValue](../../reference/office-mix/labs.core.ivalue.md) オブジェクト インスタンス。|
|[Labs.Core.IVersion](../../reference/office-mix/labs.core.iversion.md)|ラボのバージョン情報を提供します。|
|[Labs.Core.IAnalyticsConfiguration](../../reference/office-mix/labs.core.ianalyticsconfiguration.md)|カスタム分析の構成情報。ユーザーによるラボの実行でカスタム分析を表示するために、読み込む IFrame を指定することを許可します。|
|[Labs.Core.ICompletionStatus](../../reference/office-mix/labs.core.icompletionstatus.md)|ラボの完了状態。ラボを終了するとき、対話の結果を示すために、状態が渡されます。|
|[Labs.Core.ILabCallback](../../reference/office-mix/labs.core.ilabcallback.md)|Labs.js コールバック メソッドを処理するインターフェイス。|
|[Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md)|ラボに関連付けられたオブジェクト。オブジェクトには、オブジェクトの種類を示す種類フィールドが含まれています。|
|[Labs.Core.ITimelineConfiguration](../../reference/office-mix/labs.core.itimelineconfiguration.md)|[Labs.Timeline](../../reference/office-mix/labs.timeline.md) の構成オプション。一連のタイムライン構成オプションを指定するのを許可します。|
|[Labs.Core.IUserData](../../reference/office-mix/labs.core.iuserdata.md)|オブジェクトに格納されているカスタム ユーザー データを表す基本インターフェイス。|
|[Labs.Core.IValue](../../reference/office-mix/labs.core.ivalue.md)|ラボに格納されている値の基底クラス。|
|[Labs.Core.IConfiguration](../../reference/office-mix/labs.core.iconfiguration.md)|ラボ構成のデータ構造。|
|[Labs.Core.IConfigurationInstance](../../reference/office-mix/labs.core.iconfigurationinstance.md)|ラボ構成のインスタンスの基底クラス。|
|[Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md)|ラボのコンポーネントを表す基底クラス。|
|[Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md)|Labs.js をホストに接続するための抽象レイヤーを提供します。|
|[Labs.Core.ModeChangedEventData](../../reference/office-mix/labs.core.modechangedeventdata.md)|モードの変更イベントに関連付けられたデータ。|
|[Labs.Core.IEventCallback](../../reference/office-mix/labs.core.ieventcallback.md)|EventManager コールバックを処理するためのインターフェイス。|

### <a name="enumerations"></a>列挙体


|||
|:-----|:-----|
|[Labs.Core.LabMode](../../reference/office-mix/labs.core.labmode.md)|ラボの現在の状態を示す値。|
