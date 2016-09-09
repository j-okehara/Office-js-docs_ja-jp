
# Method 要素
Office アドインをアクティブにするために必要な JavaScript API for Office の個別のメソッドを指定します。

 **アドインの種類:**コンテンツ、作業ウィンドウ


## 構文:


```XML
<Method Name="string "/>
```


## 次に含まれる:

 _ [メソッド](../../reference/manifest/methods.md)_


## 属性



|**属性**|**種類**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|名前|string|必須|必要なメソッドの名前をその親オブジェクトで修飾して指定します。たとえば、**getSelectedDataAsync** メソッドを指定するには、`"Document.getSelectedDataAsync"` と指定する必要があります。|

## 注釈

**Methods** 要素と **Method** 要素はメール アドインではサポートされていません。要件セットの詳細については、「[Office ホストと API 要件を指定する](../../docs/overview/specify-office-hosts-and-api-requirements.md#SpecifyRequirementSets_intro)」をご覧ください。


 >**重要** 個々のメソッドの最小バージョン要件を指定する方法がないため、メソッドが実行時に使用可能であることを確認するには、そのメソッドをアドインのスクリプトで呼び出す際に、**if** ステートメントも使用する必要があります。これを行う方法の詳細については、「[JavaScript API for Office について](../../docs/develop/understanding-the-javascript-api-for-office.md#HostAPISupport_UsingIfStatements)」をご覧ください。

