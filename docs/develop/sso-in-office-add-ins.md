---
title: Office アドインのシングル サインオンを有効化する
description: ''
ms.date: 09/26/2018
ms.openlocfilehash: ca8280b72ab863d0e34330585fb307475e3aa9b9
ms.sourcegitcommit: 2ac7d64bb2db75ace516a604866850fce5cb2174
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/14/2018
ms.locfileid: "26298566"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a>Office アドインのシングル サインオンを有効化する (プレビュー)

ユーザーは個人用の Microsoft アカウントまたは職場や学校の (Office 365) アカウントのいずれかを使用して、Office (オンライン、モバイル、およびデスクトップ プラットフォーム) にサインインします。 これとシングル サインオン (SSO) を使用すれば、ユーザーに 2 度目のサインインを求めずに、ご自分のアドインをユーザーに許可できます。

![アドインのサインイン プロセスを示す画像](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a>プレビューの状態

現在、シングル サインオン API はプレビューのみでサポートされています。 これは、試験目的のみで開発者に提供されており、運用環境のアドインには使用してはいけません。 また、SSO を使用するアドインは [AppSource](https://appsource.microsoft.com) では許可されていません。

SSO のプレビューは、すべての Office アプリケーションではサポートされていません。 これは、Word、Excel、Outlook、および PowerPoint で利用できます。 シングル サインオン API の現在のサポート状態に関する詳細は、「[IdentityAPI の要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)」を参照してください。

### <a name="requirements-and-best-practices"></a>要件とベスト プラクティス

SSO を使用するには、アドインのスタートアップ HTML ページの `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` から Office JavaScript ライブラリのベータ版を読み込む必要があります。

**Outlook** アドインで作業している場合は、Office 365 テナントの先進認証が有効になっていることを確認してください。 この方法の詳細については、「[Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」 (Exchange Online: テナントの先進認証を有効にする方法) を参照してください。

SSO をアドインの唯一の認証方法と*しない*ようにする必要があります。 特定のエラー状況でアドインが切り替えることができる、別の認証システムを実装する必要があります。 ユーザー テーブルと認証のシステムを使用するか、ソーシャル ログイン プロバイダーの 1 つを活用できます。 Office アドインでこれを実行する方法の詳細については、「[Office アドインで外部サービスを承認する](https://docs.microsoft.com/ja-JP/office/dev/add-ins/develop/auth-external-add-ins)」を参照してください。 *Outlook* には切り替えることが可能な推奨システムがあります。 詳細については、「[シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)」を参照してください。

### <a name="how-sso-works-at-runtime"></a>実行時の SSO の動作のしくみ

次の図は、SSO の動作のしくみを示しています。

![SSO プロセスを示す図](../images/sso-overview-diagram.png)

1. アドインの JavaScript により、新しい Office.js API [getAccessTokenAsync](#sso-api-reference) が呼び出されます。 これにより、Office ホスト アプリケーションはアドインへのアクセス トークンを取得するよう指示されます。 「[アクセス トークンの例](#example-access-token)」を参照してください。
2. ユーザーがサインインしていない場合、Office ホスト アプリケーションはユーザーにサインインを求めるポップアップ ウィンドウを開きます。
3. 現在のユーザーが初めてアドインを使用する場合は、そのユーザーに同意を求めるダイアログを表示します。
4. Office ホスト アプリケーションは、Azure AD v2.0 エンドポイントから現在のユーザーの**アドイン トークン**を要求します。
5. Azure AD は、Office ホスト アプリケーションにアドイン トークンを送信します。
6. Office ホスト アプリケーションが、`getAccessTokenAsync` 呼び出しによって返される結果オブジェクトの一部として、アドインに**アドイン トークン**を送信します。
7. アドイン内の JavaScript が、トークンを解析し、ユーザーのメール アドレスなど必要な情報を抽出します。 
8. オプションで、アドインで HTTP 要求を送信して、ユーザー設定などユーザーに関する情報をさらにサーバー側から求めることができます。 または、アクセス トークン自体が解析および検証されるようにサーバー側に送信することができます。 

## <a name="develop-an-sso-add-in"></a>SSO アドインの開発

このセクションでは、SSO を使用する Office アドインの作成に関連するタスクについて説明します。 ここでは、これらのタスクについて、言語とフレームワークに依存しない方法で説明しています。 詳細なチュートリアルの例については、次を参照してください。

* [シングル サインオンを使用する Node.js Office アドインを作成する](create-sso-office-add-ins-nodejs.md)
* [シングル サインオンを使用する ASP.NET Office アドインを作成する](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a>サービス アプリケーションを作成する

Azure v2.0 エンドポイントの登録ポータル https://apps.dev.microsoft.com にアドインを登録します。 このプロセスには、次に示すタスクを含めて 5 分から 10 分の時間がかかります。

* アドインのクライアント ID とシークレットを取得します。
* アドインが必要とする AAD v.2.0 エンドポイントへのアクセス許可を指定します  (必要に応じて Microsoft Graph へも指定します)。 "profile" のアクセス許可は常に必要です。
* Office ホスト アプリケーションにアドインへの信頼を付与します。
* 既定のアクセス許可 *access_as_user* を使用して、Office ホスト アプリケーションのアドインへのアクセスを事前認証します。

この手順の詳細については、「[Azure AD v2.0 のエンドポイントに SSO を使用する Office アドインを登録する](register-sso-add-in-aad-v2.md)」をご覧ください。

### <a name="configure-the-add-in"></a>アドインを構成する

新しいマークアップをアドイン マニフェストに追加します。

* **WebApplicationInfo** - 次の要素の親。
* **Id** -このアドインのクライアント ID。これはアドインを登録する一貫として取得するアプリケーション ID です。 詳細については、「[Azure AD v2.0 のエンドポイントに SSO を使用する Office アドインを登録する](register-sso-add-in-aad-v2.md)」をご覧ください。
* **Resource** - アドインの URL。 これは、AAD にアドインを登録したときに使用したのと同じ URI (`api:` プロトコルを含む) です。 この URI のドメイン部分は、アドインのマニフェストの `<Resources>` のセクションの URL で使用されている、すべてのサブドメインを含むドメインと一致している必要があります。
* **Scopes** - 1 つ以上の **Scope** 要素の親。
* **Scope** - アドインが AAD に対して必要なアクセス許可を指定する。 `profile` のアクセス許可は常に必要です。ご使用のアドインが Microsoft Graph にアクセスしない場合、これは必要な唯一のアクセス許可になる場合があります。 アクセスする場合、Microsoft Graph へのアクセスに必要な許可として、`User.Read`、`Mail.Read` など **Scope** 要素も必要になります。 コードで使用している、Microsoft Graph にアクセスするためのライブラリでは、他にもアクセス許可が必要な場合があります。 たとえば、.NET 用の Microsoft 認証ライブラリ (MSAL) では、`offline_access` のアクセス許可が必要です。 詳細については、「[Office アドインで Microsoft Graph へ承認](authorize-to-microsoft-graph.md)」を参照してください。

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

アドインに次のために JavaScript を追加します。

* [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) の呼び出し。

* アクセス トークンを解析するか、それをアドインのサーバー側コードに渡す。 

`getAccessTokenAsync` への呼び出しの単純な例を示します。 

> [!NOTE]
> この例では、1 種類のエラーのみを明示的に処理します。 より複雑なエラー処理の例については、「[Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js)」 (Office-Add-in-ASPNET-SSO の Home.js) と 「[program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)」 (Office-Add-in-NodeJS-SSO の program.js) を参照してください。 また、「[シングル サインオン (SSO) のエラー メッセージのトラブルシューティング](troubleshoot-sso-in-office-add-ins.md)」を参照してください。
 

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

サーバー側にアドイン トークンを渡す単純な例を次に示します。 このトークンは、サーバー側に要求を戻すときの `Authorization` ヘッダーとして含まれています。 この例では JSON データの送信が想定されているので、`POST` メソッドを使用しています。ただし、サーバーに書き込まない場合は、アクセス トークンの送信に `GET` で十分です。

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

Office にログインしているユーザーがいないときにアドインを使用できない場合、*アドインを起動するときに* `getAccessTokenAsync` を呼び出す必要があります。

アドインに、ログイン ユーザーを必要としない機能がある場合、*ユーザーがログイン ユーザーを必要とするアクションを実行するときに* `getAccessTokenAsync` を呼び出します。 `getAccessTokenAsync` の重複呼び出しによってパフォーマンスが大幅に低下することはありません。これは、Office ではアクセス トークンがキャッシュされ、それが期限切れになるまで、 `getAccessTokenAsync` が呼び出されても AAD V.2.0 エンドポイントが再度呼び出されずに再利用されるためです。 このため、`getAccessTokenAsync` の呼び出しを、このトークンが必要とされる場所でアクションを開始するすべての関数とハンドラーに追加できます。

### <a name="add-server-side-code"></a>サーバー側のコードを追加する

ほとんどの場合、アドインがサーバー側に渡してそこで使用しない場合は、アクセス トークンを取得してもあまり意味はありません。 アドインでは、次のいくつかのサーバー側のタスクを実行できます。

* ホストされている使用しているデータベースのユーザー設定を検索するメソッドなど、トークンから抽出されるユーザーに関する情報を使用する 1 つ以上の Web API メソッドを作成します。 (以下の「**ID として SSO トークンを使用する**」を参照) 使用する言語とフレームワークによっては、記述する必要のあるコードを簡単に記述できるライブラリが使用できることがあります。
* Microsoft Graph データを取得します。 サーバー側のコードでは、次に示す操作を実行する必要があります。

    * アクセス トークンを検証します (以下の「**アクセス トークンを検証する**」を参照)。
    * Azure AD v2.0 エンドポイントを呼び出して、「代理」フローを開始します。これには、アクセス トークン、ユーザーに関するメタデータ、およびアドインの資格情報 (ID とシークレット) を含めます。 このコンテキストでは、アクセス トークンはブートストラップ トークンと呼ばれます。
    * 代理フローから戻される新しいアクセス トークンをキャッシュします。
    * 新しいトークンを使用して Microsoft Graph からデータを取得します。

 ユーザーの Microsoft Graph のデータへのアクセス許可を取得するには、「[Microsoft Graph への認証](authorize-to-microsoft-graph.md)」を参照してください。

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

- `name`: ユーザーの表示名。
- `preferred_username`: ユーザーの電子メール アドレス。
- `oid`: Azure Active Directory でユーザーの ID を表す GUID。
- `tid`: Azure Active Directory でユーザーの組織の ID を表す GUID。

`name` と `preferred_username` の値は変化することがあるため、この ID とバックエンドの承認サービスを、`oid` と `tid` の値を使用して関連付けることをお勧めします。

たとえば、サービスでは、これらの値を `{oid-value}@{tid-value}` のような形式にまとめて、内部ユーザー データベースのユーザーのレコードに値として保存できます。 その後の要求では、同じ値を使用してユーザーを取得できるようになり、特定のリソースへのアクセスについては既存のアクセス制御メカニズムに基づいて決定できます。

### <a name="example-access-token"></a>アクセス トークンの例

アクセス トークンの標準的なデコードされたペイロードを次に示します。 プロパティの詳細については、[Azure Active Directory v2.0 トークンのリファレンス](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens)に関するページを参照してください。


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

## <a name="using-sso-with-an-outlook-add-in"></a>Outlook のアドインで SSO を使用する

Excel、PowerPoint、または Word のアドインで SSO を使用する場合と Outlook のアドインでそれを使用する場合とでは、小さいけれど重要な違いがいくつかあります。 「[Authenticate a user with a single sign-on token in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)」 (Outlook アドインでシングル サインオン トークンを使用してユーザーを認証する) と「[シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)」を参照してください。

## <a name="sso-api-reference"></a>SSO API リファレンス

### <a name="getaccesstokenasync"></a>getAccessTokenAsync

Office Auth 名前空間 (`Office.context.auth`) には、Office ホストがアドインの Web アプリケーションへのアクセス トークンを取得することを可能にする `getAccessTokenAsync` というメソッドがあります。 これにより、間接的に、サインインしたユーザーの Microsoft Graph データにアドインがアクセスできるようにもなります。ユーザーがもう一度サインインする必要はありません。

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

このメソッドは、Azure Active Directory V 2.0 のエンドポイントを呼び出して、アドインの Web アプリケーションへのアクセス トークンを取得します。 これにより、アドインがユーザーを識別できるようになります。 ["on behalf of" OAuth フロー](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)を使用することにより、サーバー側のコードはこのトークンを使用して、アドインの Web アプリケーションの Microsoft Graph にアクセスできます。

> [!NOTE]
> Outlook でアドインが Outlook.com または Gmail のメールボックスに読み込まれている場合、この API はサポートされません。

<table><tr><td>ホスト</td><td>Excel、OneNote、Outlook、PowerPoint、Word</td></tr>

 <tr><td>[要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td><td>[IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)</td></tr></table>

#### <a name="parameters"></a>パラメーター

`options` - 省略可能。 サインイン動作を定義する `AuthOptions` オブジェクトが許可されます (下記参照)。

`callback` - 省略可能。 ユーザー ID 用のトークンを解析できるコールバック メソッドが許可されます。または、トークンを Microsoft Graph へのアクセスを取得するために、「代理」フローで使用します。 [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` が "succeeded" である場合、`AsyncResult.value` が生の AAD v. 2.0 形式のアクセス トークンになります。

`AuthOptions` インターフェイスには、Office が `getAccessTokenAsync` メソッドを使用して AAD v. 2.0 からアドインに対するアクセス トークンを取得する場合用のユーザー エクスペリエンス用のオプションがあります。

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



