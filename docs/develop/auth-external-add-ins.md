# <a name="authorize-external-services-in-your-office-add-in"></a>Office アドインで外部サービスを承認する

大手のオンライン サービス (Office 365、Google、Facebook、LinkedIn、SalesForce、GitHub など) では、開発者は、ユーザーが自分のアカウントに別のアプリケーションからアクセスできるようにすることが可能です。これにより、開発者は、これらのサービスを Office アドインに含めることができるようになります。 

Web アプリケーションからオンライン サービスへのアクセスを可能にするための業界標準のフレームワークを OAuth 2.0 と呼びます。ほとんどの場合、このフレームワークをアドインで使用するために、その動作のしくみを詳しく知る必要はありません。開発者は、この詳細を抽象化している多数のライブラリを使用できます。

OAuth の基本的な考え方は、ユーザーやグループと同様に、アプリケーションは専用の ID とアクセス許可のセットによって、それ自体がセキュリティ プリンシパルになり得るということです。通常のシナリオでは、ユーザーがオンライン サービスを必要とする Office アドインのアクションを実行すると、アドインは、ユーザーのアカウントへ特定のセットのアクセス許可を付与するよう求める要求をサービスに送信します。サービスは、該当するアクセス許可をアドインに付与するように求めるプロンプトをユーザーに表示します。アクセス許可が付与されると、サービスは小さなエンコードされた*アクセス トークン*をアドインに送信します。アドインは、サービスの API へのすべての要求にトークンを含めることで、サービスを使用できるようになります。ただし、そのアドインが実行できるアクションは、ユーザーが付与したアクセス許可の範囲内に限定されます。また、トークンは特定の時間が経過すると期限切れになります。

さまざまなシナリオに向けて、*フロー*または*許可の種類*と呼ばれる、いくつかの OAuth パターンが設計されています。次に、最も重要な 2 つのものを示します。

- **暗黙的フロー**:アドインとオンライン サービスとの通信は、クライアント側の JavaScript で実装されます。
- **認証コード フロー**:アドインの Web アプリケーションとオンライン サービスとの通信は、*サーバー間*で行われます。そのため、これはサーバー側のコードで実装されます。

フローの目的は、アプリケーションの ID と承認の安全を確保することです。認証コード フローでは、*クライアント シークレット*が提供されます。これは、秘密にしておく必要があります。単一ページ アプリケーション (SPA) には、シークレットを保護する方法がありません。そのため、SPA には暗黙的フローを使用するようにお勧めします。 

この 2 つのフローの長所と短所については、詳しく理解しておく必要があります。[認証コード](https://tools.ietf.org/html/rfc6749#section-1.3.1)と[暗黙的](https://tools.ietf.org/html/rfc6749#section-1.3.2)の公式な定義を出発点にすると良いでしょう。 

>**注:**仲介者サービスを用意するというオプションもあります。このサービスは、すべての承認を開発者のかわりに実行し、アドインにアクセス トークンを渡します。詳細については、後述の「*仲介者サービス*」を参照してください。

## <a name="using-the-implicit-flow-in-office-add-ins"></a>Office アドインに暗黙的フローを使用する
オンライン サービスが暗黙的フローをサポートしているかどうかを判断する最良の方法は、ドキュメントを調べることです。

これをサポートしているサービス用に、次に示す JavaScript ライブラリが提供されています。このライブラリは詳細部分の操作をすべて実行します。

[Office-js-helpers](https://github.com/OfficeDev/office-js-helpers)

後述の「**ライブラリ**」セクションも参照してください。

## <a name="using-the-authorization-code-flow-in-office-add-ins"></a>Office アドインに認証コード フローを使用する

次に示す、認証コード フローを使用するサンプル アドインが用意されています。

- [Office-Add-in-Nodejs-ServerAuth](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth) (NodeJS)
- [PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) (ASP.NET MVC)

各種の言語とフレームワークで認証コード フローを実装するために、利用できるライブラリが多数あります。詳細については、後述の「**ライブラリ**」セクションを参照してください。

### <a name="relayproxy-functions"></a>Relay 関数と Proxy 関数

サーバーなしの Web アプリケーションでも、サービスでホストされる簡単な関数 ([Azure Functions](https://azure.microsoft.com/en-us/services/functions) や [Amazon Lambda](https://aws.amazon.com/lambda) など) で*クライアント ID* と*クライアント シークレット*の値を使用すると、認証コード フローを使用できます。この関数は、特定のコードを適切な*アクセス トークン*に交換して、それを中継してクライアントに戻します。このアプローチのセキュリティは、関数へのアクセスが、どの程度適切に保護されているかによって異なります。

この技法を使用する場合は、アドインでオンライン サービス (Google や Facebook など) のログイン画面を示す UI やポップアップを表示します。ユーザーがオンライン サービスにログオンして、自分のリソースへのアクセス権をアドインに付与すると、開発者はオンライン関数に送信できるコードを受信します。後述の**仲介者サービス**は、これと同様のフローを使用します。 

## <a name="libraries"></a>ライブラリ

多数の言語とプラットフォーム、および両方のフローに対応したライブラリが利用できます。汎用のものや、特定のオンライン サービス向けのものがあります。 

**認証プロバイダーとして Azure Active Directory を使用する Office 365 などのサービス**:[Azure Active Directory 認証ライブラリ](https://azure.microsoft.com/en-us/documentation/articles/active-directory-authentication-libraries/)。[Microsoft 認証ライブラリ](https://www.nuget.org/packages/Microsoft.Identity.Client)のプレビュー版も利用できます。

**Google**:[GitHub.com/Google](https://github.com/google) で "auth" または目的の言語の名前を検索します。最も関連のあるリポジトリには、`google-auth-library-[name of language]` という名前が付いています。

**Facebook**:[Facebook for Developers](https://developers.facebook.com) で "library" または "sdk" を検索します。 

**汎用の OAuth 2.0**:数十の言語に対応したライブラリへのリンクが、「[OAuth Code](http://oauth.net/code/)」のページに掲載されています。このページは、IETF OAuth 作業部会によって維持されています。これらのライブラリの一部は、OAuth 準拠のサービスを実装するためのものです。アドイン開発者にとって重要なライブラリは、このページに記載された*クライアント*と呼ばれるライブラリです。これは、目的の Web サーバーが OAuth 準拠のサービスのクライアントになるためです。

## <a name="middleman-services"></a>仲介者サービス

アドインでは OAuth.io や Auth0 などの仲介者サービスを使用できます。このサービスは、多数の大手オンライン サービスに対応したアクセス トークンを提供するか、アドインのソーシャル ログインのプロセスを簡略化します (またはその両方です)。わずかなコードによって、仲介者に接続するクライアント側スクリプトまたはサーバー側コードをアドインで使用できるようになり、その後オンライン サービスに必要なトークンが仲介者から送り返されます。すべての承認の実装コードは、仲介者サービスに含まれています。 

次のサンプルでは、Auth0 を使用して Facebook、Google、および Microsoft のアカウントへのソーシャル ログインを可能にしています。

[Office-Add-in-Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)

次のサンプルでは、OAuth.io を使用して、Facebook と Google からのアクセス トークンを取得します。

[Office-Add-in-OAuth.io](https://github.com/OfficeDev/Office-Add-in-OAuth.io)

## <a name="what-is-cors"></a>CORS とは

CORS とは、[Cross Origin Resource Sharing](https://developer.mozilla.org/en-US/docs/Web/HTTP/Access_control_CORS) (クロス オリジン リソース共有) の略語です。アドイン内で CORS を使用する方法の詳細については、「[Office アドインにおける同一生成元ポリシーの制限への対処](http://dev.office.com/docs/add-ins/develop/addressing-same-origin-policy-limitations)」を参照してください。
