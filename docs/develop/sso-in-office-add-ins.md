---
title: Office アドインのシングル サインオンを有効化する
description: ''
ms.date: 12/04/2017
---

# <a name="enable-single-sign-on-for-office-add-ins-preview"></a>Office アドインのシングル サインオンを有効化する (プレビュー)

ユーザーは個人用の Microsoft アカウントまたは職場や学校の (Office 365) アカウントのいずれかを使用して、Office (オンライン、モバイル、およびデスクトップ プラットフォーム) にサインインします。 これを利用し、SSO を使用すれば、ユーザーに 2 度目のサインインを求める必要なく、次に示す操作を実行できます。

* ユーザーにアドインへのサインインを承認する。
* アドインに [Microsoft Graph](https://developer.microsoft.com/graph/docs) へのアクセスを承認する。

![アドインのサインイン プロセスを示す画像](../images/office-host-title-bar-sign-in.png)

> [!NOTE]
> 現在、シングル サインオン API は Word、Excel、Outlook、PowerPoint のプレビューでサポートされています。 シングル サインオン API の現在のサポート状態に関する詳細は、「[IdentityAPI の要件セット](https://dev.office.com/reference/add-ins/requirement-sets/identity-api-requirement-sets)」を参照してください。
> Outlook アドインで作業している場合は、Office 365 テナントの先進認証が有効になっていることを確認してください。 この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。

ユーザーにとっては、サインインが 1 回だけになり、アドインの実行エクスペリエンスがスムーズになります。開発者にとっては、アドインでユーザーを認証して、Microsoft Graph を経由したユーザーのデータへの承認済みアクセスを得るために、ユーザーが既に Office アプリケーションに提示した資格情報を使用できるということを意味します。

## <a name="sso-add-in-architecture"></a>SSO アドインのアーキテクチャ

Web アプリケーションのページと JavaScript をホスティングすることに加え、アドインでは、同一の[完全修飾ドメイン名](https://msdn.microsoft.com/ja-jp/library/windows/desktop/ms682135.aspx#_dns_fully_qualified_domain_name_fqdn__gly)で、Microsoft Graph へのアクセス トークンを取得して、要求を送信するための 1 つ以上の Web API をホストする必要もあります。

アドイン マニフェストには、アドインを Azure Active Directory (Azure AD) v2.0 エンドポイントに登録する方法と、アドインが必要とする Microsoft Graph へのアクセス許可を指定するマークアップが含まれています。

### <a name="how-it-works-at-runtime"></a>実行時の動作のしくみ

次の図は、SSO の動作のしくみを示しています。
<!-- Minor fixes to the text in the diagram - change V2 to v2.0, and change "(e.g. Word, Excel, etc.)" to "(for example, Word, Excel)". -->

![SSO プロセスを示す図](../images/sso-overview-diagram.png)

1. アドインでは、JavaScript は新しい Office.js API `getAccessTokenAsync` を呼び出します。これにより、Office ホスト アプリケーションにアドインへのアクセス トークンを取得するように指示します (今後は、これを**アドイン トークン**と呼びます)。
1. ユーザーがサインインしていない場合、Office ホスト アプリケーションはユーザーにサインインを求めるポップアップ ウィンドウを開きます。
1.  現在のユーザーが初めてアドインを使用する場合は、そのユーザーに同意を求めるダイアログを表示します。
1. Office ホスト アプリケーションは、Azure AD v2.0 エンドポイントから現在のユーザーの**アドイン トークン**を要求します。
1. Azure AD は、Office ホスト アプリケーションにアドイン トークンを送信します。
1. Office ホスト アプリケーションは、`getAccessTokenAsync` 呼び出しによって返される結果オブジェクトの一部として、アドインに**アドイン トークン**を送信します。
1. アドインの JavaScript は、アドインと同じ完全修飾ドメイン名でホストされている Web API に HTTP 要求を送信して、その要求に承認の証拠として**アドイン トークン**を含めます。  
1. サーバー側のコードでは、受け取った**アドイン トークン**を検証します。
1. サーバー側のコードは、「代理」フロー (「[OAuth2 Token Exchange](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02)」と「[デーモンまたはサーバー アプリケーション対 Web API Azure シナリオ](https://docs.microsoft.com/ja-jp/azure/active-directory/develop/active-directory-authentication-scenarios#daemon-or-server-application-to-web-api)」で定義されているフロー) を使用して、アドイン トークンの交換で Microsoft Graph のアクセス トークン (今後は、「**MSG トークン**」と呼ぶ) を取得します。
1. Azure AD は、**MSG トークン**をアドインに返します (アドインが *offline_access* アクセス許可を要求した場合は、更新トークンも返します)。
1. サーバー側のコードで、**MSG トークン**をキャッシュします。
1. サーバー側のコードで、Microsoft Graph に要求を送信して、その要求に **MSG トークン**を含めます。
1. Microsoft Graph は、アドインの UI に渡すことができるデータをアドインに返します。
1. MSG トークンの期限が切れると、サーバー側のコードは、更新トークンを使用して新しい **MSG トークン**を取得します。

## <a name="develop-an-sso-add-in"></a>SSO アドインの開発

このセクションでは、SSO を使用する Office アドインの作成に関連するタスクについて説明します。 ここでは、これらのタスクについて、言語とフレームワークに依存しない方法で説明しています。 詳細なチュートリアルの例については、次を参照してください。

* [シングル サインオンを使用する Node.js Office アドインを作成する](create-sso-office-add-ins-nodejs.md)
* [シングル サインオンを使用する ASP.NET Office アドインを作成する](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a>サービス アプリケーションを作成する

Azure v2.0 エンドポイントの登録ポータル (https://apps.dev.microsoft.com) でアドインを登録します。このプロセスには、次に示すタスクを含めて 5 分から 10 分の時間がかかります。

* アドインのクライアント ID とシークレットを取得します。
* アドインが必要とする Microsoft Graph へのアクセス許可を指定します。
* Office ホスト アプリケーション信頼をアドインに付与します。
* 既定のアクセス許可 *access_as_user* を使用して、Office ホスト アプリケーションのアドインへのアクセスを事前認証します。

### <a name="configure-the-add-in"></a>アドインを構成する

新しいマークアップをアドイン マニフェストに追加します。

* **WebApplicationInfo** - 次の要素の親。
* **Id** - アドインのクライアント ID。
* **Resource** - アドインの URL。
* **Scopes** - 1 つ以上の **Scope** 要素の親。
* **Scope** - アドインが Microsoft Graph に必要なアクセス許可を指定します。 たとえば、`User.Read`、`Mail.Read` または `offline_access` などです。 詳細については、「[Microsoft Graph のアクセス許可](https://developer.microsoft.com/en-us/graph/docs/concepts/permissions_reference)」を参照してください。

Outlook 以外の Office ホストでは、`<VersionOverrides ... xsi:type="VersionOverridesV1_0">` セクションの末尾にマークアップを追加します。Outlook では、`<VersionOverrides ... xsi:type="VersionOverridesV1_1">` セクションの末尾にマークアップを追加します。

### <a name="add-client-side-code"></a>クライアント側のコードを追加する

アドインに JavaScript を追加します。

* `Office.context.auth.getAccessTokenAsync(myTokenHandler)` を呼び出します。
* アドインのサーバー側コードにアドイン トークンを渡すためのハンドラーを作成します。たとえば、次のように入力します。

```js
function mytokenHandler(asyncResult) {
    // Passes asyncResult.value (which has the add-in access token)
    // to the add-in’s web API as an Authorization header.
}
```

### <a name="when-to-call-the-method"></a>メソッドを呼び出すタイミング

Office にログインしているユーザーがなく、Office にアドインへのアクセス トークンがないときに、アドインを使用できない場合には、*アドインを起動するときに* `getAccessTokenAsync` を呼び出す必要があります。

アドインに、Microsoft Graph またはログイン ユーザーへのアクセスを必要としない機能がある場合、*ユーザーが Microsoft Graph へのアクセス、または少なくともログイン ユーザーへのアクセスを必要とするアクションを実行するときに* `getAccessTokenAsync` を呼び出します。 `getAccessTokenAsync` の重複呼び出しによってパフォーマンスが大幅に低下することはありません。これは、Office ではアクセス トークンがキャッシュされ、それが期限切れになるまで、`getAccessTokenAsync` が呼び出されても AAD V.2.0 エンドポイントが再度呼び出されずに再利用されるためです。 このため、`getAccessTokenAsync` の呼び出しを、このトークンが必要とされる場所でアクションを開始するすべての関数とハンドラーに追加します。

### <a name="add-server-side-code"></a>サーバー側のコードを追加する

Microsoft Graph データを取得する 1 つ以上の Web API メソッドを作成します。使用する言語とフレームワークによっては、記述する必要のあるコードが簡単になるライブラリを使用できることがあります。サーバー側のコードでは、次に示す操作を実行する必要があります。

* 前の手順で作成したトークン ハンドラーから受け取ったアドイン トークンを検証します。
* Azure AD v2.0 エンドポイントの呼び出しによって、「代理」フローを開始します。これには、アドイン アクセス トークン、ユーザーに関するメタデータ、およびアドインの資格情報 (ID とシークレット) を含めます。
* 返された MSG トークンをキャッシュします。
* MSG トークンを使用して Microsoft Graph からデータを取得します。
