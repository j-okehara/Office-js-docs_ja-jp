# FormFactor 要素

特定のフォーム ファクターのアドインの設定を指定します。 たとえば、型 `MailHost` と `DesktopFormFactor` の `Host` を定義すると、デスクトップ用 Outlook には適用されますが、Web App または Outlook.com には適用されません。__ **Resources** ノードを除くそのフォーム ファクターのアドイン情報をすべて含みます。

各 FormFactor の定義には、**FunctionFile** 要素と、1 つ以上の **ExtensionPoint** 要素が含まれます。 詳細については、「[FunctionFile 要素](./functionfile.md)」と「[ExtensionPoint 要素](./extensionpoint.md)」を参照してください。 

次の FormFactors がサポートされます。

- `DesktopFormFactor` (Windows または Mac クライアント用 Office)

## 子要素

| 要素                               | 必須 | 説明  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](./extensionpoint.md) | はい      | アドインが機能を公開する場所を定義します。 |
| [FunctionFile](./functionfile.md)     | はい      | JavaScript 関数を含むファイルの URL。|
| [GetStarted](./getstarted.md)         | いいえ       | Word、Excel、または PowerPoint のホストにアドインをインストールするときに表示される吹き出しを定義します。 |

## FormFactor の例

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
