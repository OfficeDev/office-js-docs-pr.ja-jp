---
title: Office アドインのシングル サインオンを有効にする
description: ''
ms.date: 09/26/2018
ms.openlocfilehash: 1a75f7d619d2375a2f7fcb07f6afb7e0d6261ead
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579906"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a>Office アドインのシングル サインオンを有効にする (プレビュー)

ユーザーは、Office (オンライン、モバイル、デスクトップ プラットフォーム)に個人用の Microsoft アカウントまたは仕事か学校(Office 365)の アカウントのいずれかを用いてサインインします。これの特徴を利用し、ユーザーに2 回目のサインインを要求しないよう、アドインへユーザーを承認するためシングル サインオン (SSO) を使用できます。

![アドインのサインイン プロセスを示す画像](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a>プレビュー ステータス

シングル サインオンの API は、現在プレビューのみです。実験のための開発者に使用可能になります。ですが、本番のアドインで使用できません。さらに、SSO を使用するアドインは、 [AppSource](https://appsource.microsoft.com)では受け入れられません。

すべての Office アプリケーションでは、SSO のプレビューをサポートします。Word、Excel、Outlook、および PowerPoint で使用可能になります。シングル サインオンの API がサポートされている現在の詳細については、 [IdentityAPI 要件の設定](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)を参照してください。

### <a name="requirements-and-best-practices"></a>要件とベスト プラクティス

SSO を使用するには、アドインの HTML 起動ページの `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` からベータ版 Office の JavaScript ライブラリを読み込む必要があります。

 ** Outlook ** アドインで SSO を使用するには、Office 365 テナントの先進認証を有効にする必要があります。詳細な方法については、「[ Exchange Online: テナントで先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。

** SSOを、アドインの唯一の方法の認証として使用しないでください。エラーの特定の状況で、アドインに戻る可能性がある代替の認証システムを実装する必要があります。ユーザー テーブル、および認証のシステムを使用することができます。 またはソーシャル ログイン プロバイダーのいずれかを活用することができます。これを行う Office のアドインの使用方法の詳細については、 [Office アドインで外部サービスの承認](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins)を参照してください。 *Outlook*では、推奨されるフォールバック システムです。詳細についてを参照してください [シナリオ: Outlook のアドインで、サービスへのシングル サインオンを実装](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)。

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
8. オプションで、アドインはHTTP 要求をそのサーバー側にユーザーについて、ユーザーの環境などのより多くのデータを送信できます。また、自身のアクセス トークンを解析および検証するためにサーバー側に送信できます。 

## <a name="develop-an-sso-add-in"></a>SSO アドインの開発

このセクションでは、SSO を使用している Office アドインを作成するための作業について説明します。これらのタスクについては、言語とフレームワークにとらわれない方法でここで説明します。詳細なチュートリアルの例については、次のように表示されます。

* [シングル サインオンを使用する Node.js Office アドインを作成する](create-sso-office-add-ins-nodejs.md)
* [シングル サインオンを使用する ASP.NET Office アドインを作成する](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a>サービス アプリケーションを作成する

Azure v2.0 エンドポイントの登録ポータルでアドインを登録します。 https://apps.dev.microsoft.comこのプロセスには、次に示すタスクを含めて 5 分から 10 分の時間がかかります。

* アドインのクライアント ID とシークレットを取得します。
* アドイン 2.0 エンドポイント v. AAD に (および必要に応じて Microsoft Graph) を必要とするアクセス許可を指定します。”プロファイル” のアクセス許可は、常に必要です。
* Office ホスト アプリケーション信頼をアドインに付与します。
* 既定のアクセス許可 *access_as_user* を使用して、Office ホスト アプリケーションのアドインへのアクセスを事前認証します。

このプロセスの詳細については、「 [Azure AD v2.0 のエンドポイントで SSO を使用する Office アドインを登録する](register-sso-add-in-aad-v2.md)」をご覧ください。

### <a name="configure-the-add-in"></a>アドインを構成する

アドイン マニフェストに新しいマークアップを追加します。

* **WebApplicationInfo** - 次の要素の親。
* **Id** -アドインのアドインの登録の一部として取得したアプリケーション ID は、このクライアント ID です。 [Azure AD v2.0 のエンドポイントで SSO を使用している Office アドインの登録](register-sso-add-in-aad-v2.md)を参照してください。
* **Resource** - アドインの URL。
* **Scopes** - 1 つ以上の **Scope** 要素の親。
* **スコープ**  - AAD に必要な追加のアクセス許可を指定します。 `profile`許可 は、常に必要で、アドインがアクセスしていない Microsoft Graph に対して、必要な唯一のアクセス許可がある可能性があります。その場合は、また **範囲** 要素がMicrosoft Graph に必要なアクセス許可です。たとえば、 `User.Read`、 `Mail.Read`です。グラフにアクセスするためにコード内で使用するライブラリには、追加のアクセス許可が必要です。たとえば、、Microsoft 認証ライブラリ (MSAL) .NET は `offline_access` のアクセス許可を必要とします。詳細については、 [Office アドインから Graph に承認](authorize-to-microsoft-graph.md)を参照してください。

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

* [GetAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)を呼び出します。

* アクセストークンを解析するか、アドインのサーバー側コードに渡します。 

ここでは、`getAccessTokenAsync` の呼び出しの簡単な例を示します。 

> [!NOTE]
> 次の使用例は、エラーの種類を明示的に処理します。複雑なエラー処理の例については、 [Office の追加-で-ASPNET の SSO で Home.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) と [program.js では、Office の追加-で-NodeJS に SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)を参照してください。 [シングル サインオン (SSO) のエラー メッセージのトラブルシューティング](troubleshoot-sso-in-office-add-ins.md)を参照してください。
 

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

ここでは、サーバー側に追加のトークンを渡すことの簡単な例です。トークンとしては、 `Authorization` ヘッダーは、要求をサーバー側に送信するときです。次の使用例を使うように、JSON データを送信することを避けるため、 `POST` メソッドが、 `GET` サーバーに書き込むときに、アクセス トークンを送信するだけで十分です。

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

アドインがログインを必要としないいくつかの機能を持っている場合、ユーザーがログインしているユーザーを必要とするとき 、 `getAccessTokenAsync` *を呼び出します* 。 `getAccessTokenAsync` の冗長な呼び出しによるパフォーマンスの大幅な低下はありません。Office は、アクセス トークンをキャッシュし、期限切れになるまで再利用して、 `getAccessTokenAsync` が呼び出される際、AAD v. 2.0 エンドポイントに別の呼び出しを行うことはありません。すべての関数と、トークンが必要な操作を開始するハンドラへの `getAccessTokenAsync` 呼び出しを追加することができます。

### <a name="add-server-side-code"></a>サーバー側のコードを追加する

ほとんどのシナリオでは、アドインがアクセス トークンをサーバー側に渡さず、そこで使わない場合、アクセス トークンの入手はあまり意味がありません。いくつかのサーバ側のタスクでは、アドインは以下のことが可能です。

* トークンから抽出されるユーザーに関する情報を使用する 1 つまたは複数の Web API のメソッドを作成します。たとえば、メソッドを検索、ホストされているデータベース内のユーザーの基本設定をします。( **Id として SSO トークンを使用して** 以下を参照してください)。お使いの言語とフレームワークを使用して、ライブラリがあること、記述するコードが簡略化します。
* グラフのデータを取得します。サーバー側コードは、次の操作を行います。

    * アクセストークンを検証する (以下の「**アクセストークンを検証する**」を参照してください)。
    * アクセス トークン、ユーザー、およびアドイン (ID とシークレット) の資格情報に関するいくつかのメタデータが含まれている Azure AD v2.0  エンドポイントへの呼び出しに ”on behalf of” フローを開始します。ここでは、アクセス トークンは、ブートス トラップのトークンと呼ばれます。
    * "On-Behalf-Of" フローから返された新しいアクセス トークンをキャッシュします。
    * 新しいトークンを使用して Microsoft Graph からデータを取得します。

 ユーザーの Microsoft Graph データへの承認済みアクセスを取得する方法の詳細については、「[Office アドインで Microsoft Graph を承認する](authorize-to-microsoft-graph.md)」を参照してください。

#### <a name="validate-the-access-token"></a>アクセス トークンを検証する

Web API では、アクセス トークンを受信した後は、使用する前に検証にする必要があります。トークンでは、JSON Web トークン (JWT) で、検証の動作と同じように OAuth の標準的なフローでトークンの検証を意味します。JWT の検証を処理するライブラリのいくつかが、基本機能が含まれます。

- トークンが整形式であることを確認する
- トークンが意図した証明機関から発行されたことを確認する
- トークンが Web API を対象にしていることを確認する

トークンの検証時には、次のガイドラインに注意してください。

- SSO の有効なトークンは、Azure の機関によって発行される `https://login.microsoftonline.com`です。このトークン内の `iss` 要求は、この値で開始します。
- トークンの `aud` パラメータは、アドインの登録のアプリケーション ID に設定します。
- トークンの `scp` パラメータは `access_as_user` に設定します。

#### <a name="using-the-sso-token-as-an-identity"></a>ID として SSO トークンを使用する

アドインは、ユーザーの身元を確認する必要がある場合、SSO トークンには、id を確立するために使用できる情報が含まれています。トークンの次の要求は、id に関連しています。

- `name` -ユーザーの表示名。
- `preferred_username` - ユーザーの電子メール アドレス。
- `oid` -Azure Active Directory でユーザーの ID を表す GUID。
- `tid` - Azure Active Directory でユーザーの組織の ID を表す GUID。

`name` と `preferred_username` の値は変化することがあるため、`oid` と `tid` の値を ID とバックエンドの承認サービスを関連付けるために使用するようにしてください。

例えば、サービスは、それらの値を `{oid-value}@{tid-value}`と一緒のようにフォーマットでき、内部のユーザー データベース内のユーザーのレコードの値として格納します。以降の要求で、ユーザー同じ値を使用して取得でき、特定のリソースへのアクセスは、既存のアクセス制御機構に基づて判断できます。

### <a name="example-access-token"></a>アクセス トークンの例

以下は、アクセス トークンの標準的なデコードされたペイロードです。プロパティの詳細については、 [Azure Active Directory バージョン 2.0 のトークンの参照](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens)を参照してください。


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

SSO を Outlook のアドインで使用する場合と、Excel、PowerPoint、または  Word のアドインで使用する場合とでは、小さくても重要な違いがあります。「[Outlook アドインでシングル サインオン トークンを使用してユーザーを認証する](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)」と「[シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)」を必ずお読みください。

## <a name="sso-api-reference"></a>SSO API 参照

### <a name="getaccesstokenasync"></a>getAccessTokenAsync

Office Authの名前空間で、 `Office.context.auth`は、メソッドを用意しています。 `getAccessTokenAsync`は Office ホストがアドインの web アプリケーションにアクセス トークンを取得できるようにします。間接的に、これはアドインが、ユーザーに二回目のサインインを必要とさせず、サインイン中のユーザーの Microsoft のグラフのデータにアクセスすることを可能とします。

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

メソッドは、アドインの web アプリケーションにアクセス トークンを取得するため、 Azure Active Directory V 2.0 エンドポイントを呼び出します。これにより、アドインを使用してユーザーを識別できます。サーバー側コードでは、 [ "on behalf of"  OAuth フロー](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)を使用して、アドインの Web アプリケーションのMicrosoft Graphにアクセスするのに、このトークンを使用できます。

> [!NOTE]
> Outlook では、アドインが Outlook.com または Gmail のメールボックスに読み込まれている場合は、この API はサポートされていません。

<table><tr><td>Hosts</td><td>Excel、OneNote、Outlook、PowerPoint、Word</td></tr>

 <tr><td>[要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td><td>[IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)</td></tr></table>

#### <a name="parameters"></a>パラメータ

`options` -省略可能です。サインオン時の動作を定義するために、`AuthOptions` オブジェクト (下記参照) を受け入れます。

`callback` -省略可能です。ユーザーの ID のトークンを解析したり、Microsoft のグラフへのアクセスを取得する "on behalf of"  フローでトークンを使用するコールバック メソッドを受け取ります。[AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` が「成功した」場合、 `AsyncResult.value` は、  生のAAD v. 2.0 形式のアクセス トークンです。

`AuthOptions` インターフェイスは、Office が`getAccessTokenAsync` メソッドでAAD v からアドインへのアクセス トークンを取得する場合に、ユーザー エクスペリエンスのオプションを提供します。

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



