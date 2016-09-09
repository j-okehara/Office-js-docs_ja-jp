# Hosts 要素

Office アドインをアクティブにする Office クライアント アプリケーションを指定します。 **Host** 要素のコレクションとその設定が含まれます。 

[VersionOverrides](./versionoverrides.md) ノードに含まれる場合、この要素は、マニフェストの親部分の **Hosts** 要素よりも優先されます。 

## 子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [Host](#host)    |  はい   |  ホストとその設定について説明します。 |

> ** 注:** Outlook では、`Hosts` に `MailHost` の `Host` 定義を含める必要があります。

---- 

## Host 要素
“ドキュメント“、“ブック“、“プレゼンテーション“、“プロジェクト“、“メールボックス“、“ノートブック“ などの、アドインでアクティブ化すべき Office アプリケーションの種類を指定します。

### 属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  はい  | これらの設定を適用する Office ホストについて説明します。|

### 子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [FormFactor](./formfactor.md)    |  はい   |  影響を受けるフォーム ファクターを定義します。 |


### xsi:type
含まれている設定を適用する Office ホスト (Word、Excel、PowerPoint、Outlook、OneNote) を制御します。 この値は、次のいずれかである必要があります。

- `MailHost` (Outlook)    


## ホストの例 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
