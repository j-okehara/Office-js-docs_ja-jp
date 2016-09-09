
# Exchange Web サービス マネージ API トークン検証ライブラリを使用する

Exchange Server 2013 または Exchange Online を実行するサーバーにアドインが要求する ID トークンを使用して、Outlook アドインのクライアントを識別できます。JSON Web トークンとしてフォーマットされたトークンは、Exchange サーバーの電子メール アカウントに一意の識別子を提供します。Exchange Web Services (EWS) マネージ API には ID トークンの使用を簡素化するヘルパー クラスがあります。

## 検証ライブラリを使用する前提条件

Exchange の ID トークンを検証するには、[EWS マネージ API ライブラリ](https://www.nuget.org/packages/Microsoft.Exchange.WebServices)をインストールする必要があります。

## Exchange の ID トークンを検証する

EWS マネージ API 検証ライブラリには、Exchange の ID トークンを管理する **AppIdentityToken** クラスがあります。次のメソッドは **AppIdentityToken** インスタンスを作成し、**Validate** メソッドを呼び出して、トークンが有効であることを検証する方法を示します。メソッドは、以下のパラメーターをとります。

- *rawToken*:[**Office.context.mailbox.getUserIdentityTokenAsync**](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox) メソッドから、トークンの文字列表現が Outlook アドインに返されます。
- *hostUri*:**getUserIdentityTokenAsync** と呼ばれる、Outlook アドインのページへの完全修飾 URI。

```C#
// Required to use the validation library.
using Microsoft.Exchange.WebServices.Auth.Validate;

private AppIdentityToken CreateAndValidateIdentityToken(string rawToken, string hostUri)
{
    try
    {
        AppIdentityToken token = (AppIdentityToken)AuthToken.Parse(rawToken);
        token.Validate(new Uri(hostUri));

        return token;
    }
    catch (TokenValidationException ex)
    {
        throw new ApplicationException("A client identity token validation error occurred.", ex);
    }
}
```

## その他のリソース

- [Exchange の ID トークンを使用して Outlook アドインを認証する](../outlook/authentication.md)  
- [Exchange の ID トークンの内部](../outlook/inside-the-identity-token.md)
- [Exchange の ID トークンを検証する](../outlook/validate-an-identity-token.md)
    
