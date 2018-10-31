---
title: Office アドインのシングル サインオンを有効にする
description: ''
ms.date: 09/26/2018
ms.openlocfilehash: fb4eacee9419339116e15ef3fccc03b291faf3ec
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506029"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a>Office アドインのシングル サインオンを有効にする (プレビュー)

ユーザーは、個人用の Microsoft アカウントまたは職場や学校の (Office 365) アカウントのいずれかを用いて Office (オンライン、モバイル、デスクトップ プラットフォーム) にサインインします。この特性を生かして、ユーザーに2 回目のサインインを要求しないよう、アドインへユーザーを承認するためのシングル サインオン (SSO) を使用できます。

![アドインのサインイン プロセスを示す画像](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a>プレビュー ステータス

シングル サインオンの API は、現在プレビュー版のみがサポートされています。開発者は検証用に使用できますが、製品版のアドインで使用できません。また SSO を使用するアドインは、 [AppSource](https://appsource.microsoft.com) には対応していません。

SSO のプレビュー版は、すべての Office アプリケーションでサポートしているわけではありません。使用できるのは、Word、Excel、Outlook、PowerPoint です。シングル サインオンの API のサポート状況の詳細については、「 [IdentityAPI 要件の設定](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js) 」を参照してください。

### <a name="requirements-and-best-practices"></a>要件とベスト プラクティス

SSO を使用するには、アドインの HTML 起動ページの `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` から Office JavaScript ライブラリのベータ版を読み込む必要があります。

 **Outlook** アドインで SSO を使用するには、Office 365 テナントの先進認証を有効にする必要があります。詳細な方法については、「 [Exchange Online: テナントで先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx) 」を参照してください。

アドインの唯一の認証方法として SSOを使用することは *おやめください* 。特定のエラー時に使用できる代替の認証システムを実装する必要があります。ユーザー テーブルおよび認証システム、あるいはソーシャル ログイン プロバイダーのいずれかを使用できます。Office アドインでこれを行う方法は、「[Office アドインで外部サービスを承認する](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins)」を参照してください。*Outlook*では、推奨されるフォールバック システムがあります。詳細は「[シナリオ: Outlook のアドインで、サービスへのシングル サインオンを実装する](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)」をご覧ください。

### <a name="how-sso-works-at-runtime"></a>実行時に SSO が動作する仕組み

次の図は、SSO の動作の仕組みを示しています。

![SSO プロセスを示す図](../images/sso-overview-diagram.png)

1. アドインでは、JavaScript は新しい Office.js API [getAccessTokenAsync](#sso-api-reference) を呼び出します。これにより、Office ホスト アプリケーションにアドインへのアクセス トークンを取得するように指示します。「[アクセス トークンの使用例](#example-access-token)」 をご覧ください。
2. ユーザーがサインインしていない場合、Office ホスト アプリケーションはユーザーにサインインを求めるポップアップ ウィンドウを開きます。
3. 現在のユーザーが初めてアドインを使用する場合は、そのユーザーの同意を求めるダイアログを表示します。
4. Office ホスト アプリケーションは、Azure AD v2.0 エンドポイントに現在のユーザーの**アドイン トークン**を要求します。
5. Azure AD は、Office ホスト アプリケーションにアドイン トークンを送信します。
6. Office ホスト アプリケーションは、`getAccessTokenAsync` 呼び出しによって返される結果オブジェクトの一部として、アドインに**アドイン トークン**を送信します。
7. アドインの JavaScript は、トークンを解析し、必要な情報 (ユーザーの電子メールアドレスなど) を抽出できます。 
8. オプションで、アドインはサーバー側に HTTP 要求を送信して、ユーザーに関する詳細なデータ (ユーザーの嗜好など) を得ることができます。また、自身のアクセス トークンを解析および検証用にサーバー側に送信することもできます。 

## <a name="develop-an-sso-add-in"></a>SSO アドインの開発

このセクションでは、SSO を使用する Office アドインを作成するための作業について説明します。これらのタスクについては、言語とフレームワークにとらわれない方法でここで説明します。詳細なチュートリアルの例については、次のように表示されます。

* [シングル サインオンを使用する Node.js Office アドインを作成する](create-sso-office-add-ins-nodejs.md)
* [シングル サインオンを使用する ASP.NET Office アドインを作成する](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a>サービス アプリケーションを作成する

Azure v2.0 エンドポイントの登録ポータルでアドインを登録します。 https://apps.dev.microsoft.comこのプロセスには、次に示すタスクを含めて 5 分から 10 分の時間がかかります。

* アドインのクライアント ID とシークレットを取得します。
* AAD v. 2.0 エンドポイント  (および必要に応じて Microsoft Graph) で必要とするアクセス許可を指定します。「プロファイル」のアクセス許可は、常に必要です。
* Office ホスト アプリケーション信頼をアドインに付与します。
* 既定のアクセス許可 *access_as_user* を使用して、Office ホスト アプリケーションのアドインへのアクセスを事前認証します。

このプロセスの詳細については、「 [Azure AD v2.0 のエンドポイントで SSO を使用する Office アドインを登録する](register-sso-add-in-aad-v2.md)」をご覧ください。

### <a name="configure-the-add-in"></a>アドインを構成する

アドイン マニフェストに新しいマークアップを追加します。

* **WebApplicationInfo** - 次の要素の親。
* **ID** -アドインのクライアント ID。これは、アドイン登録の一環として取得するアプリケーション IDです。「 [Azure AD v2.0 のエンドポイントで SSO を使用している Office アドインの登録](register-sso-add-in-aad-v2.md) 」を参照してください。
* **Resource** - アドインの URL。
* **スコープ** - 1 つ以上の **スコープ** 要素の親。
*  **スコープ**  - AAD に必要な追加のアクセス許可を指定します。 `profile` 許可 は、常に必要で、アドインがアクセスしていない Microsoft Graph に対して、必要な唯一のアクセス許可がある可能性があります。その場合は、また**  スコープ** 要素が Microsoft Graph に必要なアクセス許可です。たとえば、 `User.Read`、 `Mail.Read`です。グラフにアクセスするためにコード内で使用するライブラリには、追加のアクセス許可が必要です。たとえば、、Microsoft 認証ライブラリ (MSAL) .NET は `offline_access` のアクセス許可を必要とします。詳細については、 [Office アドインから Graph に承認](authorize-to-microsoft-graph.md)を参照してください。

Outlook 以外の Office ホストでは、 `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` セクションの末尾にマークアップを追加します。Outlook では、 `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` セクションの末尾にマークアップを追加します。

マークアップの例を次に示します。

```xml
<WebApplicationInfo>
    <Id>5661fed9-f33d-4e95-b6cf-624a34a2f51d</Id>
    <Resource>api://addin.contoso.com/5661fed9-f33d-4e95-b6cf-624a34a2f51d</Resource>
    <Scopes>
        <Scope>user.read</Scope>
        <Scope>files.read</Scope>
        <Scope>profile</Scope>
    </Scopes>
</WebApplicationInfo>
```

### <a name="add-client-side-code"></a>クライアント側のコードを追加する

アドインに JavaScript を追加します。

* [GetAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)を呼び出します。

* アクセストークンを解析するか、アドインのサーバー側コードに渡します。 

 `getAccessTokenAsync` の呼び出しの簡単な例を示します。 

> [!NOTE]
> この例は、一種類のみのエラーを処理しているものです。より複雑なエラー処理の例については、「 [Office-Add-in-ASPNET-SSO の Home.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) 」および「 [Office-Add-in-NodeJS-SSO の program.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js) 」を参照してください。また、「 [シングル サインオン (SSO) のエラー メッセージのトラブルシューティング](troubleshoot-sso-in-office-add-ins.md) 」もご覧ください。
 

```js
Office.context.auth.getAccessTokenAsync(function (result) {
    if (result.status === "succeeded") {
        // Use this token to call Web API
        var ssoToken = result.value;
        ...
    } else {
        if (result.error.code === 13003) {
            // SSO is not supported for domain user accounts, only
            // work or school (Office 365) or Microsoft Account IDs.
        } else {
            // Handle error
        }
    }
});
```

以下は、サーバー側に追加のトークンを渡すことの簡単な例です。トークンは、サーバー側に要求を送信するときに `Authorization` ヘッダーの一つとして含まれます。この例では JSON データを送信しているため `POST` メソッドを使用していますが、サーバーへの書き込みを行わない場合は、アクセス トークンの送信に `GET` だけで十分です。

```js
$.ajax({
    type: "POST",
    url: "/api/DoSomething",
    headers: {
        "Authorization": "Bearer " + ssoToken
    },
    data: { /* some JSON payload */ },
    contentType: "application/json; charset=utf-8"
}).done(function (data) {
    // Handle success
}).fail(function (error) {
    // Handle error
}).always(function () {
    // Cleanup
});
```

#### <a name="when-to-call-the-method"></a>メソッドを呼び出すタイミング

Office にログインしているユーザーがなく、アドインを使用できない場合には、*アドインを起動するときに* `getAccessTokenAsync` を呼び出す必要があります。

アドインがログインを必要としないいくつかの機能を持っている場合、* ユーザーがログイン中のユーザーを必要とするアクションを実行する際*  、 `getAccessTokenAsync` を呼び出します 。 `getAccessTokenAsync` の重複呼び出しによってパフォーマンスgが大幅に低下することはありません。これは、Office ではアクセス トークンがキャッシュされ、それが期限切れになるまで、 `getAccessTokenAsync` が呼び出されても AAD v. 2.0 エンドポイントを再度呼び出すことなく再利用するためです。そのため `getAccessTokenAsync` の呼び出しを、すべての関数と、トークンを必要とするアクションを開始するハンドラへ追加することができます。

### <a name="add-server-side-code"></a>サーバー側のコードを追加する

ほとんどのシナリオでは、アドインが取得しサーバーに渡したアクセス トークンをサーバー側で使用しない場合は、アクセス トークンを取得することに意味はありません。いくつかのサーバ側のタスクでは、アドインは以下のことが可能です。

* トークンから抽出されたユーザーに関する情報を使用する 1 つ以上の Web API のメソッドを作成します。たとえば、ホスト型データベース内のユーザーの嗜好を参照するメソッドです。(「 **SSO トークンを ID としてを使用する** 」を参照してください。) お使いの言語とフレームワークによっては、記述するコードを簡略化することのできるライブラリがある場合があります。
* Microsoft Graph のデータを取得します。サーバー側コードは、次の操作を行います。

    * アクセストークンを検証する (以下の「 **アクセストークンを検証する** 」を参照してください)。
    * アクセス トークン、ユーザー、およびアドイン (ID とシークレット) の資格情報に関するいくつかのメタデータが含まれている Azure AD v2.0  エンドポイントへの呼び出しに ”on behalf of” フローを開始します。ここでは、アクセス トークンは、ブートス トラップのトークンと呼ばれます。
    * "On-Behalf-Of" フローから返された新しいアクセス トークンをキャッシュします。
    * 新しいトークンを使用して Microsoft Graph からデータを取得します。

 ユーザーの Microsoft Graph データへの承認済みアクセスを取得する方法の詳細については、「[Office アドインで Microsoft Graph を承認する](authorize-to-microsoft-graph.md)」を参照してください。

#### <a name="validate-the-access-token"></a>アクセス トークンを検証する

Web API がアクセス トークンを受信した後は、使用前に検証を行う必要があります。このトークンは JSON Web トークン (JWT) であるため、その検証はほとんどの標準 OAuth のトークン検証と同様に動作します。JWT の検証を処理するライブラリはいくつかありますが、基本は以下のものです。

- トークンが整形式であることを確認する
- トークンが意図した証明機関から発行されたことを確認する
- トークンが Web API を対象にしていることを確認する

トークンを検証するときは、次のガイドラインに留意してください。

- 有効な SSO トークンは Azure 証明機関 `https://login.microsoftonline.com` から発行されます。トークン内の `iss` 要求は、この値で始まっている必要があります。
- トークンの `aud` パラメータは、アドインの登録のアプリケーション ID に設定します。
- トークンの `scp` パラメータは `access_as_user` に設定します。

#### <a name="using-the-sso-token-as-an-identity"></a>ID として SSO トークンを使用する

アドインでユーザーの ID を検証する必要がある場合、SSO トークンには ID を確定するために使用できる情報が含まれています。以下は、ID に関連するトークンの要求です。

- `name` - ユーザーの表示名。
- `preferred_username` - ユーザーの電子メール アドレス。
- `oid` - Azure Active Directory でユーザーの ID を表す GUID。
- `tid` - Azure Active Directory でユーザーの組織の ID を表す GUID。

 `name` と `preferred_username` の値は変化することがあるため、ID とバックエンドの承認サービスを関連付けるために `oid` と `tid` の値を使用するようにしてください。

たとえば、サービスでは、これらの値を `{oid-value}@{tid-value}` のような形式にまとめて、内部ユーザー データベースのユーザーのレコードに値として保存できます。その後の要求では、同じ値を使用してユーザーを取得できるようになり、特定のリソースへのアクセスについては既存のアクセス制御メカニズムに基づいて決定できます。

### <a name="example-access-token"></a>アクセス トークンの例

以下は、アクセス トークンの標準的なデコードされたペイロードです。プロパティの詳細については、「 [Azure Active Directory バージョン 2.0 のトークンの参照](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens) 」をご覧ください。


```js
{
    aud: "2c3caa80-93f9-425e-8b85-0745f50c0d24",         
    iss: "https://login.microsoftonline.com/fec4f964-8bc9-4fac-b972-1c1da35adbcd/v2.0",         
    iat: 1521143967,         
    nbf: 1521143967,         
    exp: 1521147867,         
    aio: "ATQAy/8GAAAA0agfnU4DTJUlEqGLisMtBk5q6z+6DB+sgiRjB/Ni73q83y0B86yBHU/WFJnlMQJ8",         
    azp: "e4590ed6-62b3-5102-beff-bad2292ab01c",         
    azpacr: "0",         
    e_exp: 262800,         
    name: "Mila Nikolova",         
    oid: "6467882c-fdfd-4354-a1ed-4e13f064be25",         
    preferred_username: "milan@contoso.com",         
    scp: "access_as_user",         
    sub: "XkjgWjdmaZ-_xDmhgN1BMP2vL2YOfeVxfPT_o8GRWaw",         
    tid: "fec4f964-8bc9-4fac-b972-1c1da35adbcd",         
    uti: "MICAQyhrH02ov54bCtIDAA",         
    ver: "2.0"
}
```

## <a name="using-sso-with-an-outlook-add-in"></a>Outlook アドインでの SSO の使用

SSO を Outlook のアドインで使用する場合と、Excel、PowerPoint、または  Word のアドインで使用する場合とでは、小さくても重要な違いがあります。「[Outlook アドインでシングル サインオン トークンを使用してユーザーを認証する](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)」と「[シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)」を必ずお読みください。

## <a name="sso-api-reference"></a>SSO API 参照

### <a name="getaccesstokenasync"></a>getAccessTokenAsync

Office Auth の名前空間 `Office.context.auth` は、Office ホストがアドインの Web アプリケーションへのアクセス トークンを取得できるようにするメソッド `getAccessTokenAsync` を提供します。これにより間接的に、ユーザーが2回目のサインインを行うことなく、アドインがサインイン中のユーザーの Microsoft Graph のデータにアクセスすることも可能になります。

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

このメソッドは、アドインの Web アプリケーションにアクセス トークンを取得するため、 Azure Active Directory V 2.0 エンドポイントを呼び出します。これにより、アドインを使用してユーザーを識別できます。 ["on behalf of"  OAuth フロー](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of) を使用することにより、サーバー側のコードはこのトークンを使用して、アドインの Web アプリケーションの Microsoft Graph にアクセスできます。

> [!NOTE]
> Outlook では、アドインが Outlook.com または Gmail のメールボックスに読み込まれている場合は、この API はサポートされていません。

<table><tr><td>Hosts</td><td>Excel、OneNote、Outlook、PowerPoint、Word</td></tr>

 <tr><td>要件セット</td><td>[IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td></tr></table>

#### <a name="parameters"></a>パラメータ

`options` - 任意。サインオン時の動作を定義するために、 `AuthOptions` オブジェクト (下記参照) を受け入れます。

`callback` - 任意。ユーザーの ID のトークンを解析したり、Microsoft Graph へのアクセスを取得する "on behalf of" フローでトークンを使用したりするコールバック メソッドを受け取ります。 [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` が「成功」した場合、 `AsyncResult.value` が AAD v. 2.0 形式の生アクセス トークンとなります。

 `AuthOptions` インターフェイスは、Office が `getAccessTokenAsync` メソッドで AAD v. 2.0 からアドインへのアクセス トークンを取得する場合に、ユーザー エクスペリエンスのオプションを提供します。

```typescript
interface AuthOptions {
    /**
        * Causes Office to display the add-in consent experience. Useful if the add-in's Azure permissions have changed or if the user's consent has 
        * been revoked.
        */
    forceConsent?: boolean,
    /**
        * Prompts the user to add their Office account (or to switch to it, if it is already added).
        */
    forceAddAccount?: boolean,
    /**
        * Causes Office to prompt the user to provide the additional factor when the tenancy being targeted by Microsoft Graph requires multifactor 
        * authentication. The string value identifies the type of additional factor that is required. In most cases, you won't know at development 
        * time whether the user's tenant requires an additional factor or what the string should be. So this option would be used in a "second try" 
        * call of getAccessTokenAsync after Microsoft Graph has sent an error requesting the additional factor and containing the string that should 
        * be used with the authChallenge option.
        */
    authChallenge?: string
    /**
        * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
        */
    asyncContext?: any
}
```



