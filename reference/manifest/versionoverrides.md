# VersionOverrides 要素

アドインによって実装されたアドイン コマンドに関する情報を格納するルート要素です。 **VersionOverrides** は、マニフェスト内の [OfficeApp](./officeapp.md) 要素の子要素です。 この要素は、マニフェスト スキーマ v1.1 以降でサポートされていますが、VersionOverrides v1.0 スキーマで定義されています。 

## 属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **xmlns**       |  はい  |  スキーマの場所。`http://schemas.microsoft.com/office/mailappversionoverrides` にする必要があります。|
|  **xsi:type**  |  はい  | スキーマのバージョン。 この時点で有効な値は `VersionOverridesV1_0` のみです。 |


## 子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  **説明**    |  いいえ   |  アドインについての説明。 これは、マニフェスト内の任意の親部分の `Description` 要素を上書きします。 説明のテキストは、[Resources](./resources.md) 要素の **LongString** 要素の子要素に含まれています。 **Description** 要素の `resid` の属性は、テキストを含む `String` 要素の `id` 属性の値に設定されています。|
|  **Requirements**  |  いいえ   |  アドインに必要な最小の Office.js のセットおよびバージョンを指定します。 これは、マニフェストの親部分の `Requirements` 要素を上書きします。| 
|  [Hosts](./hosts.md)                |  はい  |  Office ホストのコレクションを指定します。 子の Host 要素は、マニフェストの親部分の Host 要素を上書きします。  |
|  [Resources](./resources.md)    |  はい  | マニフェストの他の要素によって参照されるリソースのコレクション (文字列、URL、画像) を定義します。|



### VersionOverrides の例
```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources> 
      <!-- add information on resources -->
   </Resources>
</VersionOverrides>
...
</OfficeApp>
```
