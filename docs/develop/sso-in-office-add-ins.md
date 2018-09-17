---
title: Office アドインのシングル サインオンを有効化する
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 534ac41e7518756a2aa5b4408ce7adb0f434e27d
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945342"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a>Office アドインのシングル サインオンを有効化する (プレビュー)

ユーザーは個人用の Microsoft アカウントまたは職場や学校の (Office 365) アカウントのいずれかを使用して、Office (オンライン、モバイル、およびデスクトップ プラットフォーム) にサインインします。 これを利用して、シングル サインオン (SSO) を使用すれば、ユーザーに 2 度目のサインインを求めなくても、ユーザーにアドインの使用を承認できます。


![アドインのサインインプロセスを示す画像](../images/office-host-title-bar-sign-in.png)

> [!NOTE]
> 現在、シングル サインオン API は Word、Excel、Outlook、PowerPoint のプレビューでサポートされています。 シングル サインオン API の現在のサポート状態に関する詳細は、「[IdentityAPI の要件セット](https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets?view=office-js)」を参照してください。
> SSO を使用するには、アドインの HTML 起動ページの https://appsforoffice.microsoft.com/lib/beta/hosted/office.js からベータ版 Office の JavaScript ライブラリを読み込む必要があります。
> Outlook アドインで作業している場合は、Office 365 テナントの先進認証が有効になっていることを確認してください。 この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。

ユーザーにとっては、サインインが一度だけになり、アドインの実行エクスペリエンスがスムーズになります。 つまり開発者にとっては、暗号化されたパスワードを使った独自のユーザー テーブルをアドインに残さなくても良いということです。

### <a name="how-it-works-at-runtime"></a>実行時の動作のしくみ

次の図は、SSO の動作のしくみを示しています。

![SSO プロセスを示す図](../images/sso-overview-diagram.png)

1. アドインでは、JavaScript は新しい Office.js API [getAccessTokenAsync](#sso-api-reference) を呼び出します。 これにより、Office ホスト アプリケーションにアドインへのアクセス トークンを取得するように指示します。 [アクセス トークンの例](#example-access-token) を参照してください。
2. ユーザーがサインインしていない場合、Office ホスト アプリケーションはユーザーにサインインを求めるポップアップ ウィンドウを開きます。
3. 現在のユーザーが初めてアドインを使用する場合は、そのユーザーに同意を求めるダイアログを表示します。
4. Office ホスト アプリケーションは、Azure AD v2.0 エンドポイントから現在のユーザーの**アドイン トークン**を要求します。
5. Azure AD は、Office ホスト アプリケーションにアドイン トークンを送信します。
6. Office ホスト アプリケーションは、`getAccessTokenAsync` 呼び出しによって返される結果オブジェクトの一部として、アドインに**アドイン トークン**を送信します。
7. アドインの JavaScript は、トークンを解析し、必要な情報 (ユーザーの電子メールアドレスなど) を抽出できます。 
8. アプションで、アドインはサーバー側に HTTP 要求を送信して、ユーザーに関する詳細なデータ (ユーザーの嗜好など) を得ることができます。 もしくは、アクセス トークン自体をサーバー側に送信して、解析と検証をすることもできます。 

## <a name="develop-an-sso-add-in"></a>SSO アドインの開発

このセクションでは、SSO を使用する Office アドインの作成に関連するタスクについて説明します。 ここでは、これらのタスクについて、言語とフレームワークに依存しない方法で説明しています。 詳細なチュートリアルの例については、次を参照してください。

* [シングル サインオンを使用する Node.js Office アドインを作成する](create-sso-office-add-ins-nodejs.md)
* [シングル サインオンを使用する ASP.NET Office アドインを作成する](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a>サービス アプリケーションを作成する

Azure v2.0 エンドポイントの登録ポータルでアドインを登録します。https://apps.dev.microsoft.com。 これは、次のタスクを含む 5 〜 10 分程度の時間がかかるプロセスです。

* アドインのクライアント ID とシークレットを取得します。
* アドインが AAD v に必要とするアクセス許可を指定します。 (オプションで Microsoft Graph)。 [プロファイル] のアクセス許可は、常に必要です。
* Office ホスト アプリケーション信頼をアドインに付与します。
* 既定のアクセス許可 *access_as_user* を使用して、Office ホスト アプリケーションのアドインへのアクセスを事前認証します。

このプロセスの詳細については、「 [Azure AD v2.0 のエンドポイントで SSO を使用する Office アドインを登録する](register-sso-add-in-aad-v2.md)」をご覧ください。

### <a name="configure-the-add-in"></a>アドインを構成する

新しいマークアップをアドイン マニフェストに追加します。

* **WebApplicationInfo** - 次の要素の親。
* **Id** - アドインのクライアント ID。これは、アドイン登録の一環として取得するアプリケーション ID です。 「[Azure AD v2.0 のエンドポイントで SSO を使用する Office アドインを登録する](register-sso-add-in-aad-v2.md)」をご覧ください。
* **Resource** - アドインの URL。
* **Scopes** - 1 つ以上の **Scope** 要素の親。
* **Scope** - アドインが AAD に必要とするアクセス許可を指定します。 アクセス許可は常に必要とされます。アドインに Microsoft Graph へのアクセス権がないとき、唯一必要な権限はこの許可である場合があります。`profile` このような場合、必要な Microsoft Graph へのアクセス許可を得るには **Scope** 要素も必要です。たとえば、 `User.Read`、 `Mail.Read` などです。 Microsoft Graph にアクセスするためにコードで使用するライブラリには、アクセス許可がさらに必要な場合があります。 たとえば、Microsoft 認証ライブラリ (MSAL) for .NET は、 `offline_access` アクセス許可が必要です。 詳細については、「 [Office アドインから Microsoft Graph を承認する](authorize-to-microsoft-graph.md)」を参照してください。

Outlook 以外の Office ホストでは、`<VersionOverrides ... xsi:type="VersionOverridesV1_0">` セクションの末尾にマークアップを追加します。Outlook では、`<VersionOverrides ... xsi:type="VersionOverridesV1_1">` セクションの末尾にマークアップを追加します。

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

*  [GetAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)を呼び出します。

* アクセストークンを解析するか、アドインのサーバー側コードに渡すかします。 

ここでは、`getAccessTokenAsync` の呼び出しの簡単な例を示します。 

> [!NOTE]
> この例では、1 種類のエラーのみを明示的に処理しています。 より詳細なエラー処理の例については、「[Office-Add-in-ASPNET-SSO の Home.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js)」および「[Office-Add-in-NodeJS-SSO の program.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)」を参照してください。 また、「[シングル サインオン (SSO) のエラー メッセージのトラブルシューティング](troubleshoot-sso-in-office-add-ins.md)」 もご覧ください。
 

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

ここでは、アドイン トークンをサーバー側に渡す簡単な例を示します。 サーバー側に要求を送り返すときは、トークンが `Authorization` ヘッダーとして含まれます。 この例では、JSON データを送信すると想定しています。そのため、`POST` メソッドを使用しますが、サーバーへの書き込みを行っていない場合、アクセス トークンの送信には `GET` で十分です。

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

アドインにログイン ユーザーを必要としない機能がある場合、*ユーザーがログイン ユーザーを必要とするアクションを実行するときに* `getAccessTokenAsync` を呼び出します。 `getAccessTokenAsync`  の重複呼び出しによってパフォーマンスが大幅に低下することはありません。これは、Office ではアクセス トークンがキャッシュされ、それが期限切れになるまで、AAD V が再度呼び出されることなく、再利用されるためです。 が呼び出されたときに毎回 AAD v. 2.0 エンドポイントを呼び出すことなく再利用するためです。`getAccessTokenAsync` このため、`getAccessTokenAsync` の呼び出しを、このトークンが必要とされる場所でアクションを開始するすべての関数とハンドラーに追加します。

### <a name="add-server-side-code"></a>サーバー側のコードを追加する

ほとんどの場合、アドインがサーバー側に渡してそこで使用しない場合は、アクセス トークンを取得してもあまり意味はありません。 アドインで実行できるサーバー側タスクの一部を次に示します。

* トークンから抽出されたユーザーに関する情報を使用する 1 つ以上の Web API メソッドを作成します。たとえば、ホスト型データベース内のユーザーの嗜好を参照するメソッドです。 (以下の「 **ID として SSO トークンを使用する** 」を参照してください)。使用している言語とフレームワークによっては、作成する必要があるコードを簡略化できるライブラリを利用できる場合があります。
* Microsoft Graph データを取得します。 サーバー側のコードでは、次に示す操作を実行する必要があります。

    * アクセストークンを検証する (以下の「**アクセストークンを検証する**」を参照してください)。
    * アクセス トークン、ユーザーに関するメタデータ、アドインの認証情報 (ID とシークレット) を含む Azure AD v2.0 エンドポイントへの呼び出しを使用して "On-Behalf-Of" フローを開始します。 このコンテキストでは、アクセストークンはブートストラップトークンと呼ばれます。
    * "On-Behalf-Of" フローから返された新しいアクセス トークンをキャッシュします。
    * 新しいトークンを使用して Microsoft Graph からデータを取得します。

 ユーザーの Microsoft Graph データへの承認済みアクセスを取得する方法の詳細については、「[Office アドインで Microsoft Graph を承認する](authorize-to-microsoft-graph.md)」を参照してください。

#### <a name="validate-the-access-token"></a>アクセス トークンを検証する

Web API でアクセス トークンを受信したら、そのアクセス トークンを使用する前に検証する必要があります。 このトークンは、JSON Web トークン (JWT) です。そのため、この検証は最も標準的な OAuth でのトークンの検証とまったく同様に動作します。 JWT の検証を処理できるライブラリが複数入手可能ですが、その基本は次のとおりです。

- トークンが整形式であることを確認する
- トークンが意図した証明機関から発行されたことを確認する
- トークンが Web API を対象にしていることを確認する

トークンの検証時には、次のガイドラインに注意してください。

- 有効な SSO トークンは Azure 証明機関 `https://login.microsoftonline.com` から発行されます。 トークン内の `iss` クレームは、この値で始まっている必要があります。
- トークンの `aud` パラメーターは、アドインの登録のアプリケーション ID に設定します。
- トークンの `scp` パラメーターは `access_as_user` に設定します。

#### <a name="using-the-sso-token-as-an-identity"></a>ID として SSO トークンを使用する

アドインでユーザーの ID を検証する必要がある場合、SSO トークンには ID を確定するために使用できる情報が含まれています。 ID に関連するトークン内のクレームは次のとおりです。

- `name` : ユーザーの表示名。
- `preferred_username` ユーザーの電子メール アドレスです。
- `oid` : Azure Active Directory でユーザーの ID を表す GUID。
- `tid` : Azure Active Directory でユーザーの組織の ID を表す GUID。

`name` と `preferred_username` の値は変化することがあるため、`oid` と `tid` の値を ID とバックエンドの承認サービスを関連付けるために使用するようにしてください。

たとえば、サービスでは、これらの値を `{oid-value}@{tid-value}` のような形式にまとめて、内部ユーザー データベースのユーザーのレコードに値として保存できます。 その後の要求では、同じ値を使用してユーザーを取得できるようになり、特定のリソースへのアクセスについては既存のアクセス制御メカニズムに基づいて決定できます。

### <a name="example-access-token"></a>アクセス トークンの例

アクセス トークンの標準的なデコードされたペイロードを次に示します。 プロパティの詳細については、「[Azure Active Directory v2.0 トークンリファレンス](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens)」を参照してください。


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

## <a name="using-sso-with-an-outlook-add-in"></a>Outlook のアドインでの SSO の使用

SSO を Outlook アドインとして使用する場合と、Excel、PowerPoint、Word アドインとして使用する場合には、ささやかですが重要な違いがあります。 「[Outlook アドインでシングルサインオントークンを使用してユーザーを認証する](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)」と「[シナリオ: Outlook アドインでサービスにシングルサインオンを実装する](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)」を必ず読んでください。

## <a name="sso-api-reference"></a>SSO API リファレンス

### <a name="getaccesstokenasync"></a>getAccessTokenAsync

Office の認証の名前空間 `Office.context.auth` は、Office ホストがアドインの Web アプリケーションへのアクセス トークンを取得できるようにするメソッド `getAccessTokenAsync` を提供します。 これはまた間接的に、ユーザーがもう一度サインインする必要なしに、アドインがサインインしたユーザーの Microsoft Graph データにアクセスできるようにします。

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

このメソッドは、Azure Active Directory V 2.0 のエンドポイントを呼び出して、アドインの Web アプリケーションへのアクセス トークンを取得します。 これにより、アドインを使用してユーザーを識別できます。 ["on behalf of" OAuth フロー](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)を使用することにより、サーバー側のコードはこのトークンを使用して、アドインの Web アプリケーションの Microsoft Graph にアクセスできます。

> [!NOTE]
> [!Outlook メモによれば、アドインが Outlook.com または Gmail のメールボックスに読み込まれている場合は、この API はサポートされていません。]

<table><tr><td>Hosts</td><td>Excel、OneNote、Outlook、PowerPoint、Word</td></tr>

 <tr><td>要件セット</td><td>[IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td></tr></table>

#### <a name="parameters"></a>パラメーター

`options` - 省略可能。 サインオン時の動作を定義するために、`AuthOptions` オブジェクト (下記参照) を受け入れます。

`callback` - 省略可能。 ユーザーの ID のトークンを解析したり、"On-Behalf-Of" フローのトークンを使用して Microsoft Graph へのアクセスを取得することができるコールバック メソッドを受け入れます。 [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` が「成功した」場合は、 `AsyncResult.value` は生の AAD v です。 2.0 形式のアクセス トークンです。

インターフェイスは、Office が AAD v からアドインへのアクセス トークンを取得する場合に、ユーザー エクスペリエンスのオプションを提供します。`AuthOptions` メソッドを持つ 2.0 です。`getAccessTokenAsync`

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



