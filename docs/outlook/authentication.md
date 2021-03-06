
# <a name="authenticate-an-outlook-add-in-by-using-exchange-identity-tokens"></a>Exchange の ID トークンを使用して Outlook アドインを認証する

Outlook アドインは、アドインをホストするサーバーであれ、内部ネットワークであれ、クラウド内のその他どこかの場所であれ、インターネット上の任意の場所にある情報をユーザーに提供できます。ただし、その情報が保護されている場合は、Exchange の電子メール アカウントと情報サービスを関連付ける手段がアドインに必要です。Exchange 2013 では、要求を行っている電子メール アカウントを識別するトークンを提供することで、アドインでシングル サインオン (SSO) を有効にできます。そのトークンをアプリケーションの登録ユーザーに関連付けることにより、アドインがいつサービスに接続してもユーザーが認識されるようにできます。

## <a name="identity-tokens"></a>ID トークン


2 つのサンプル アドインがあります。どちらもパブリックに利用できる情報を使用します。1 つはメッセージ内のアドレスの Bing マップを表示するアドイン、もう 1 つはメッセージ内の YouTube ビデオ リンクのプレビューを表示するアドインです。もちろん、アドインはパブリックではない情報にアクセスすることもできます。アドインをホストするサーバーを使用して、内部ネットワークやクラウド内の情報にアドインをリンクできます。

アドイン ユーザーの識別と認証にはさまざまな方法を使用できます。Exchange 2013では、特定の Exchange 電子メール アカウントを識別する ID トークンをアドインに提供することで、ユーザー認証を簡素化します。サービスのこのトークンを登録ユーザーに関連付けると、Outlook アドインを使用するユーザーに対してシングル サインオン (SSO) を有効にできます。 

アドインで SSO を使用するには、コードで次の処理を行います。


* ID トークンを返す Outlook アドイン API の関数を呼び出します。
* トークンと要求をサーバーに送信します。
* サーバーからの応答を展開して、サービスからの情報を表示します。
    
サーバー側の処理はもう少し複雑です。サーバーは Outlook アドインから要求を受信すると、次のような処理を行います。

* トークンを検証します。[管理トークン検証ライブラリ](../../docs/outlook/use-the-token-validation-library.md)を使用できます。また、サービス用の[独自のライブラリを作成](../../docs/outlook/validate-an-identity-token.md)することもできます。
* トークンから一意の識別子を調べて、既知の ID と関連付けられているかどうかを確認します。サービスでは、そのサービスの既知のユーザーとその[識別子を照合するメソッドを実装する](../../docs/outlook/authenticate-a-user-with-an-identity-token.md)必要があります。
* 一意の識別子が、一連の資格情報と共にサーバーに既に保存されている識別子と一致する場合、サーバーは、ユーザーがサービスにログオンしなくても、要求された情報に応答できます。
* 一意の識別子が不明の場合は、ユーザーに対してログインにサーバーの資格情報を使用するように求める応答を送信します。
* 資格情報がサーバー上の既知の ID と一致する場合は、その ID をトークン内の一意の識別子にマップし、次に要求を受け取ったときに、追加のログオン手順を必要とすることなく、サーバーが応答できるようにします。

 >**メモ**  これは ID トークンを使用する方法に関する 1 つの提案に過ぎません。ID と認証を処理するときは、コードが組織のセキュリティ要件を満たしていることを確認する必要があります。

詳細な説明は、以下の記事を参照してください。これらの記事では、ID トークンとメッセージ内で見つかった電話番号のリストを Web サービスに送信する簡単な Outlook アドインを使用します。 

- [Exchange の ID トークンの内部](../outlook/inside-the-identity-token.md)
- [Exchange で ID トークンを使用して Outlook アドインからサービスを呼び出す](../outlook/call-a-service-by-using-an-identity-token.md)
- [Exchange のトークン検証ライブラリを使用する](../outlvalidate-an-identity-token.md ook/use-the-token-validation-library.md)
- [Exchange の ID トークンを検証する](../outlook/validate-an-identity-token.md )
- [Exchange の ID トークンを使用してユーザーを認証する](../outlook/validate-an-identity-token.md)


## <a name="additional-resources"></a>その他のリソース



- [Outlook アドイン](../outlook/outlook-add-ins.md)
    
- [Outlook アドインから Web サービスを呼び出す](../outlook/web-services.md)
    


