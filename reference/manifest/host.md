
# Host 要素
Office アドインがサポートしている Office のホスト アプリケーションの種類を指定します。

 **アドインの種類:**コンテンツ、作業ウィンドウ、メール


## 構文:


```XML
<Host Name= ["Document" | "Database" | "Mailbox" | "Presentation" | "Project" | "Workbook"] />
```


## 属性



|**属性**|**種類**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|名前|string|必須|Office ホスト アプリケーションの種類の名前。|

## 注釈

以下の値を **Host** 要素の **Name** 属性に指定できます。それぞれの値はアドインがサポートする 1 つ以上の Office ホスト アプリケーションのセットに対応します。



|**名前**|**Office ホスト アプリケーション**|
|:-----|:-----|
| `"Document"`|Word、Word Online、Word (iPad)。|
| `"Database"`|Access Web アプリ|
| `"Mailbox"`|Outlook、Outlook Web App、デバイス用 OWA|
| `"Notebook"`|OneNote Online|
| `"Presentation"`|PowerPoint、PowerPoint Online、PowerPoint (iPad)|
| `"Project"`|Project|
| `"Workbook"`|Excel、Excel Online、Excel (iPad)|

## 注釈

ホストのサポートを指定する方法の詳細については、「[Office ホストと API 要件を指定する](../../docs/overview/specify-office-hosts-and-api-requirements.md)」をご覧ください。

