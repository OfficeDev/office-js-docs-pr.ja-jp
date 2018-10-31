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
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a><span data-ttu-id="6e792-102">Office アドインのシングル サインオンを有効にする (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="6e792-102">Enable single sign-on for Office Add-ins (preview)</span></span>

<span data-ttu-id="6e792-p101">ユーザーは、Office (オンライン、モバイル、デスクトップ プラットフォーム)に個人用の Microsoft アカウントまたは仕事か学校(Office 365)の アカウントのいずれかを用いてサインインします。これの特徴を利用し、ユーザーに2 回目のサインインを要求しないよう、アドインへユーザーを承認するためシングル サインオン (SSO) を使用できます。</span><span class="sxs-lookup"><span data-stu-id="6e792-p101">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account. You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>

![アドインのサインイン プロセスを示す画像](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a><span data-ttu-id="6e792-106">プレビュー ステータス</span><span class="sxs-lookup"><span data-stu-id="6e792-106">Preview Status</span></span>

<span data-ttu-id="6e792-p102">シングル サインオンの API は、現在プレビューのみです。実験のための開発者に使用可能になります。ですが、本番のアドインで使用できません。さらに、SSO を使用するアドインは、 [AppSource](https://appsource.microsoft.com)では受け入れられません。</span><span class="sxs-lookup"><span data-stu-id="6e792-p102">The Single Sign-on API is currently supported in preview only. It is available to developers for experimentation; but it should not be used in a production add-in. In addition, add-ins that use SSO are not accepted in [AppSource](https://appsource.microsoft.com).</span></span>

<span data-ttu-id="6e792-p103">すべての Office アプリケーションでは、SSO のプレビューをサポートします。Word、Excel、Outlook、および PowerPoint で使用可能になります。シングル サインオンの API がサポートされている現在の詳細については、 [IdentityAPI 要件の設定](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6e792-p103">Not all Office applications support the SSO preview. It is available in Word, Excel, Outlook, and PowerPoint. For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js).</span></span>

### <a name="requirements-and-best-practices"></a><span data-ttu-id="6e792-113">要件とベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="6e792-113">Requirements and Best Practices</span></span>

<span data-ttu-id="6e792-114">SSO を使用するには、アドインの HTML 起動ページの `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` からベータ版 Office の JavaScript ライブラリを読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="6e792-114">To use SSO, you must load the beta version of the Office JavaScript Library from `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` in the startup HTML page of the add-in.</span></span>

<span data-ttu-id="6e792-p104"> \*\* Outlook \*\* アドインで SSO を使用するには、Office 365 テナントの先進認証を有効にする必要があります。詳細な方法については、「[ Exchange Online: テナントで先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6e792-p104">If you are working with an **Outlook** add-in, be sure to enable Modern Authentication for the Office 365 tenancy. For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="6e792-p105">\*\* SSOを、アドインの唯一の方法の認証として使用しないでください。エラーの特定の状況で、アドインに戻る可能性がある代替の認証システムを実装する必要があります。ユーザー テーブル、および認証のシステムを使用することができます。 またはソーシャル ログイン プロバイダーのいずれかを活用することができます。これを行う Office のアドインの使用方法の詳細については、 [Office アドインで外部サービスの承認](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins)を参照してください。 *Outlook*では、推奨されるフォールバック システムです。詳細についてを参照してください [シナリオ: Outlook のアドインで、サービスへのシングル サインオンを実装](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)。</span><span class="sxs-lookup"><span data-stu-id="6e792-p105">You should *not* rely on SSO as your add-in's only method of authentication. You should implement an alternate authentication system that your add-in can fall back to in certain error situations. You can use a system of user tables and authentication, or you can leverage one of the social login providers. For more information about how to do this with an Office add-in, see [Authorize external services in your Office Add-in](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins). For *Outlook*, there is a recommended fall back system. For more information, see [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

### <a name="how-sso-works-at-runtime"></a><span data-ttu-id="6e792-123">実行時に SSO が動作する仕組み</span><span class="sxs-lookup"><span data-stu-id="6e792-123">How it works at runtime</span></span>

<span data-ttu-id="6e792-124">次の図は、SSO の動作の仕組みを示しています。</span><span class="sxs-lookup"><span data-stu-id="6e792-124">The following diagram shows how the SSO process works.</span></span>

![SSO プロセスを示す図](../images/sso-overview-diagram.png)

1. <span data-ttu-id="6e792-p106">アドインでは、JavaScript は新しい Office.js API [getAccessTokenAsync](#sso-api-reference) を呼び出します。これにより、Office ホスト アプリケーションにアドインへのアクセス トークンを取得するように指示します。「[アクセス トークンの使用例](#example-access-token)」 をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="6e792-p106">In the add-in, JavaScript calls a new Office.js API [getAccessTokenAsync](#sso-api-reference). This tells the Office host application to obtain an access token to the add-in. See [Example access token](#example-access-token).</span></span>
2. <span data-ttu-id="6e792-129">ユーザーがサインインしていない場合、Office ホスト アプリケーションはユーザーにサインインを求めるポップアップ ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="6e792-129">If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="6e792-130">現在のユーザーが初めてアドインを使用する場合は、そのユーザーの同意を求めるダイアログを表示します。</span><span class="sxs-lookup"><span data-stu-id="6e792-130">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="6e792-131">Office ホスト アプリケーションは、Azure AD v2.0 エンドポイントに現在のユーザーの**アドイン トークン**を要求します。</span><span class="sxs-lookup"><span data-stu-id="6e792-131">The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="6e792-132">Azure AD は、Office ホスト アプリケーションにアドイン トークンを送信します。</span><span class="sxs-lookup"><span data-stu-id="6e792-132">Azure AD sends the add-in token to the Office host application.</span></span>
6. <span data-ttu-id="6e792-133">Office ホスト アプリケーションは、`getAccessTokenAsync` 呼び出しによって返される結果オブジェクトの一部として、アドインに**アドイン トークン**を送信します。</span><span class="sxs-lookup"><span data-stu-id="6e792-133">The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessTokenAsync` call.</span></span>
7. <span data-ttu-id="6e792-134">アドインの JavaScript は、トークンを解析し、必要な情報 (ユーザーの電子メールアドレスなど) を抽出できます。</span><span class="sxs-lookup"><span data-stu-id="6e792-134">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span> 
8. <span data-ttu-id="6e792-p107">オプションで、アドインはHTTP 要求をそのサーバー側にユーザーについて、ユーザーの環境などのより多くのデータを送信できます。また、自身のアクセス トークンを解析および検証するためにサーバー側に送信できます。</span><span class="sxs-lookup"><span data-stu-id="6e792-p107">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences. Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span> 

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="6e792-137">SSO アドインの開発</span><span class="sxs-lookup"><span data-stu-id="6e792-137">Develop an SSO add-in</span></span>

<span data-ttu-id="6e792-p108">このセクションでは、SSO を使用している Office アドインを作成するための作業について説明します。これらのタスクについては、言語とフレームワークにとらわれない方法でここで説明します。詳細なチュートリアルの例については、次のように表示されます。</span><span class="sxs-lookup"><span data-stu-id="6e792-p108">This section describes the tasks involved in creating an Office Add-in that uses SSO. These tasks are described here in a language- and framework-agnostic way. For examples of detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="6e792-141">シングル サインオンを使用する Node.js Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="6e792-141">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="6e792-142">シングル サインオンを使用する ASP.NET Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="6e792-142">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a><span data-ttu-id="6e792-143">サービス アプリケーションを作成する</span><span class="sxs-lookup"><span data-stu-id="6e792-143">Create the service application</span></span>

<span data-ttu-id="6e792-p109">Azure v2.0 エンドポイントの登録ポータルでアドインを登録します。 https://apps.dev.microsoft.comこのプロセスには、次に示すタスクを含めて 5 分から 10 分の時間がかかります。</span><span class="sxs-lookup"><span data-stu-id="6e792-p109">Register the add-in at the registration portal for the Azure v2.0 endpoint: https://apps.dev.microsoft.com. This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="6e792-146">アドインのクライアント ID とシークレットを取得します。</span><span class="sxs-lookup"><span data-stu-id="6e792-146">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="6e792-p110">アドイン 2.0 エンドポイント v. AAD に (および必要に応じて Microsoft Graph) を必要とするアクセス許可を指定します。”プロファイル” のアクセス許可は、常に必要です。</span><span class="sxs-lookup"><span data-stu-id="6e792-p110">Specify the permissions that your add-in needs to AAD v. 2.0 endpoint (and optionally to Microsoft Graph). The "profile" permission is always needed.</span></span>
* <span data-ttu-id="6e792-150">Office ホスト アプリケーション信頼をアドインに付与します。</span><span class="sxs-lookup"><span data-stu-id="6e792-150">Grant the Office host application trust to the add-in.</span></span>
* <span data-ttu-id="6e792-151">既定のアクセス許可 *access_as_user* を使用して、Office ホスト アプリケーションのアドインへのアクセスを事前認証します。</span><span class="sxs-lookup"><span data-stu-id="6e792-151">Preauthorize the Office host application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="6e792-152">このプロセスの詳細については、「 [Azure AD v2.0 のエンドポイントで SSO を使用する Office アドインを登録する](register-sso-add-in-aad-v2.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="6e792-152">For more information, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="6e792-153">アドインを構成する</span><span class="sxs-lookup"><span data-stu-id="6e792-153">Configure the add-in</span></span>

<span data-ttu-id="6e792-154">アドイン マニフェストに新しいマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="6e792-154">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="6e792-155">**WebApplicationInfo** - 次の要素の親。</span><span class="sxs-lookup"><span data-stu-id="6e792-155">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="6e792-p111">**Id** -アドインのアドインの登録の一部として取得したアプリケーション ID は、このクライアント ID です。 [Azure AD v2.0 のエンドポイントで SSO を使用している Office アドインの登録](register-sso-add-in-aad-v2.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6e792-p111">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in. See [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="6e792-158">**Resource** - アドインの URL。</span><span class="sxs-lookup"><span data-stu-id="6e792-158">**Resource** - The URL of the add-in.</span></span>
* <span data-ttu-id="6e792-159">**Scopes** - 1 つ以上の **Scope** 要素の親。</span><span class="sxs-lookup"><span data-stu-id="6e792-159">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="6e792-p112">**スコープ**  - AAD に必要な追加のアクセス許可を指定します。 `profile`許可 は、常に必要で、アドインがアクセスしていない Microsoft Graph に対して、必要な唯一のアクセス許可がある可能性があります。その場合は、また **範囲** 要素がMicrosoft Graph に必要なアクセス許可です。たとえば、 `User.Read`、 `Mail.Read`です。グラフにアクセスするためにコード内で使用するライブラリには、追加のアクセス許可が必要です。たとえば、、Microsoft 認証ライブラリ (MSAL) .NET は `offline_access` のアクセス許可を必要とします。詳細については、 [Office アドインから Graph に承認](authorize-to-microsoft-graph.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6e792-p112">**Scope** - Specifies a permission that the add-in needs to AAD. The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph. If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`. Libraries that you use in your code to access Microsoft Graph may need additional permissions. For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission. For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="6e792-p113">Outlook 以外の Office ホストでは、`<VersionOverrides ... xsi:type="VersionOverridesV1_0">` セクションの末尾にマークアップを追加します。Outlook では、`<VersionOverrides ... xsi:type="VersionOverridesV1_1">` セクションの末尾にマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="6e792-p113">For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="6e792-168">マークアップの例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="6e792-168">The following is an example of the markup:</span></span>

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

### <a name="add-client-side-code"></a><span data-ttu-id="6e792-169">クライアント側のコードを追加する</span><span class="sxs-lookup"><span data-stu-id="6e792-169">Add client-side code</span></span>

<span data-ttu-id="6e792-170">アドインに JavaScript を追加します。</span><span class="sxs-lookup"><span data-stu-id="6e792-170">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="6e792-171">[GetAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="6e792-171">Call [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span></span>

* <span data-ttu-id="6e792-172">アクセストークンを解析するか、アドインのサーバー側コードに渡します。</span><span class="sxs-lookup"><span data-stu-id="6e792-172">Parse the access token or pass it to the add-in’s server-side code.</span></span> 

<span data-ttu-id="6e792-173">ここでは、`getAccessTokenAsync` の呼び出しの簡単な例を示します。</span><span class="sxs-lookup"><span data-stu-id="6e792-173">Here's a simple example of a call to `getAccessTokenAsync`.</span></span> 

> [!NOTE]
> <span data-ttu-id="6e792-p114">次の使用例は、エラーの種類を明示的に処理します。複雑なエラー処理の例については、 [Office の追加-で-ASPNET の SSO で Home.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) と [program.js では、Office の追加-で-NodeJS に SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)を参照してください。 [シングル サインオン (SSO) のエラー メッセージのトラブルシューティング](troubleshoot-sso-in-office-add-ins.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6e792-p114">This example handles only one kind of error explicitly. For examples of more elaborate error handling, see [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) and [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js). And see [Troubleshoot error messages for single sign-on (SSO)](troubleshoot-sso-in-office-add-ins.md).</span></span>
 

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

<span data-ttu-id="6e792-p115">ここでは、サーバー側に追加のトークンを渡すことの簡単な例です。トークンとしては、 `Authorization` ヘッダーは、要求をサーバー側に送信するときです。次の使用例を使うように、JSON データを送信することを避けるため、 `POST` メソッドが、 `GET` サーバーに書き込むときに、アクセス トークンを送信するだけで十分です。</span><span class="sxs-lookup"><span data-stu-id="6e792-p115">Here's a simple example of passing the add-in token to the server-side. The token is included as an `Authorization` header when sending a request back to the server-side. This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

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

#### <a name="when-to-call-the-method"></a><span data-ttu-id="6e792-180">メソッドを呼び出すタイミング</span><span class="sxs-lookup"><span data-stu-id="6e792-180">When to call the method</span></span>

<span data-ttu-id="6e792-181">Office にログインしているユーザーがなく、アドインを使用できない場合には、*アドインを起動するときに* `getAccessTokenAsync` を呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="6e792-181">If your add-in cannot be used when a no user is logged into Office and Office does not have an access token to your add-in, then you should call `getAccessTokenAsync` *when the add-in launches*.</span></span>

<span data-ttu-id="6e792-p116">アドインがログインを必要としないいくつかの機能を持っている場合、ユーザーがログインしているユーザーを必要とするとき 、 `getAccessTokenAsync` *を呼び出します* 。 `getAccessTokenAsync` の冗長な呼び出しによるパフォーマンスの大幅な低下はありません。Office は、アクセス トークンをキャッシュし、期限切れになるまで再利用して、 `getAccessTokenAsync` が呼び出される際、AAD v. 2.0 エンドポイントに別の呼び出しを行うことはありません。すべての関数と、トークンが必要な操作を開始するハンドラへの `getAccessTokenAsync` 呼び出しを追加することができます。</span><span class="sxs-lookup"><span data-stu-id="6e792-p116">If the add-in has some functionality that doesn't require a logged in user, then you call `getAccessTokenAsync` *when the user takes an action that requires a logged in user*. There is no significant performance degradation with redundant calls of `getAccessTokenAsync` because Office caches the access token and will reuse it, until it expires, without making another call to the AAD v. 2.0 endpoint whenever `getAccessTokenAsync` is called. So you can add calls of `getAccessTokenAsync` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="6e792-186">サーバー側のコードを追加する</span><span class="sxs-lookup"><span data-stu-id="6e792-186">Add server-side code</span></span>

<span data-ttu-id="6e792-p117">ほとんどのシナリオでは、アドインがアクセス トークンをサーバー側に渡さず、そこで使わない場合、アクセス トークンの入手はあまり意味がありません。いくつかのサーバ側のタスクでは、アドインは以下のことが可能です。</span><span class="sxs-lookup"><span data-stu-id="6e792-p117">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there. Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="6e792-p118">トークンから抽出されるユーザーに関する情報を使用する 1 つまたは複数の Web API のメソッドを作成します。たとえば、メソッドを検索、ホストされているデータベース内のユーザーの基本設定をします。( **Id として SSO トークンを使用して** 以下を参照してください)。お使いの言語とフレームワークを使用して、ライブラリがあること、記述するコードが簡略化します。</span><span class="sxs-lookup"><span data-stu-id="6e792-p118">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base. (See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="6e792-p119">グラフのデータを取得します。サーバー側コードは、次の操作を行います。</span><span class="sxs-lookup"><span data-stu-id="6e792-p119">Get Microsoft Graph data. Your server-side code should do the following:</span></span>

    * <span data-ttu-id="6e792-193">アクセストークンを検証する (以下の「**アクセストークンを検証する**」を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="6e792-193">Validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="6e792-p120">アクセス トークン、ユーザー、およびアドイン (ID とシークレット) の資格情報に関するいくつかのメタデータが含まれている Azure AD v2.0  エンドポイントへの呼び出しに ”on behalf of” フローを開始します。ここでは、アクセス トークンは、ブートス トラップのトークンと呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="6e792-p120">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the access token, some metadata about the user, and the credentials of the add-in (its ID and secret). In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="6e792-196">"On-Behalf-Of" フローから返された新しいアクセス トークンをキャッシュします。</span><span class="sxs-lookup"><span data-stu-id="6e792-196">Cache the new access token that is returned from the on-behalf-of flow.</span></span>
    * <span data-ttu-id="6e792-197">新しいトークンを使用して Microsoft Graph からデータを取得します。</span><span class="sxs-lookup"><span data-stu-id="6e792-197">Get data from Microsoft Graph by using the MSG token.</span></span>

 <span data-ttu-id="6e792-198">ユーザーの Microsoft Graph データへの承認済みアクセスを取得する方法の詳細については、「[Office アドインで Microsoft Graph を承認する](authorize-to-microsoft-graph.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6e792-198">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="6e792-199">アクセス トークンを検証する</span><span class="sxs-lookup"><span data-stu-id="6e792-199">For more information, see Validate the access token.</span></span>

<span data-ttu-id="6e792-p121">Web API では、アクセス トークンを受信した後は、使用する前に検証にする必要があります。トークンでは、JSON Web トークン (JWT) で、検証の動作と同じように OAuth の標準的なフローでトークンの検証を意味します。JWT の検証を処理するライブラリのいくつかが、基本機能が含まれます。</span><span class="sxs-lookup"><span data-stu-id="6e792-p121">Once the Web API receives the access token, it must validate it before using it. The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows. There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="6e792-203">トークンが整形式であることを確認する</span><span class="sxs-lookup"><span data-stu-id="6e792-203">Checking that the token is well-formed</span></span>
- <span data-ttu-id="6e792-204">トークンが意図した証明機関から発行されたことを確認する</span><span class="sxs-lookup"><span data-stu-id="6e792-204">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="6e792-205">トークンが Web API を対象にしていることを確認する</span><span class="sxs-lookup"><span data-stu-id="6e792-205">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="6e792-206">トークンの検証時には、次のガイドラインに注意してください。</span><span class="sxs-lookup"><span data-stu-id="6e792-206">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="6e792-p122">SSO の有効なトークンは、Azure の機関によって発行される `https://login.microsoftonline.com`です。このトークン内の `iss` 要求は、この値で開始します。</span><span class="sxs-lookup"><span data-stu-id="6e792-p122">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`. The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="6e792-209">トークンの `aud` パラメータは、アドインの登録のアプリケーション ID に設定します。</span><span class="sxs-lookup"><span data-stu-id="6e792-209">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="6e792-210">トークンの `scp` パラメータは `access_as_user` に設定します。</span><span class="sxs-lookup"><span data-stu-id="6e792-210">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="6e792-211">ID として SSO トークンを使用する</span><span class="sxs-lookup"><span data-stu-id="6e792-211">Using the SSO token as an identity</span></span>

<span data-ttu-id="6e792-p123">アドインは、ユーザーの身元を確認する必要がある場合、SSO トークンには、id を確立するために使用できる情報が含まれています。トークンの次の要求は、id に関連しています。</span><span class="sxs-lookup"><span data-stu-id="6e792-p123">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity. The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="6e792-214">`name` -ユーザーの表示名。</span><span class="sxs-lookup"><span data-stu-id="6e792-214">`name` - The user's display name.</span></span>
- <span data-ttu-id="6e792-215">`preferred_username` - ユーザーの電子メール アドレス。</span><span class="sxs-lookup"><span data-stu-id="6e792-215">`preferred_username`The user's email address.</span></span>
- <span data-ttu-id="6e792-216">`oid` -Azure Active Directory でユーザーの ID を表す GUID。</span><span class="sxs-lookup"><span data-stu-id="6e792-216">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="6e792-217">`tid` - Azure Active Directory でユーザーの組織の ID を表す GUID。</span><span class="sxs-lookup"><span data-stu-id="6e792-217">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="6e792-218">`name` と `preferred_username` の値は変化することがあるため、`oid` と `tid` の値を ID とバックエンドの承認サービスを関連付けるために使用するようにしてください。</span><span class="sxs-lookup"><span data-stu-id="6e792-218">Since the `name` and `preferred_username` values could change, it's recommended that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="6e792-p124">例えば、サービスは、それらの値を `{oid-value}@{tid-value}`と一緒のようにフォーマットでき、内部のユーザー データベース内のユーザーのレコードの値として格納します。以降の要求で、ユーザー同じ値を使用して取得でき、特定のリソースへのアクセスは、既存のアクセス制御機構に基づて判断できます。</span><span class="sxs-lookup"><span data-stu-id="6e792-p124">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database. Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="6e792-221">アクセス トークンの例</span><span class="sxs-lookup"><span data-stu-id="6e792-221">Example access token</span></span>

<span data-ttu-id="6e792-p125">以下は、アクセス トークンの標準的なデコードされたペイロードです。プロパティの詳細については、 [Azure Active Directory バージョン 2.0 のトークンの参照](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6e792-p125">The following is a typical decoded payload of an access token. For information about the properties, see [Azure Active Directory v2.0 tokens reference](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).</span></span>


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

## <a name="using-sso-with-an-outlook-add-in"></a><span data-ttu-id="6e792-224">Outlook のアドインでの SSO の使用</span><span class="sxs-lookup"><span data-stu-id="6e792-224">Using SSO with and Outlook add-in</span></span>

<span data-ttu-id="6e792-p126">SSO を Outlook のアドインで使用する場合と、Excel、PowerPoint、または  Word のアドインで使用する場合とでは、小さくても重要な違いがあります。「[Outlook アドインでシングル サインオン トークンを使用してユーザーを認証する](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)」と「[シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)」を必ずお読みください。</span><span class="sxs-lookup"><span data-stu-id="6e792-p126">There are some small, but important differences in using SSO in an Outlook add-in from using it in an Excel, PowerPoint, or Word add-in. Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) and [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

## <a name="sso-api-reference"></a><span data-ttu-id="6e792-227">SSO API 参照</span><span class="sxs-lookup"><span data-stu-id="6e792-227">SSO API reference</span></span>

### <a name="getaccesstokenasync"></a><span data-ttu-id="6e792-228">getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="6e792-228">getAccessTokenAsync</span></span>

<span data-ttu-id="6e792-p127">Office Authの名前空間で、 `Office.context.auth`は、メソッドを用意しています。 `getAccessTokenAsync`は Office ホストがアドインの web アプリケーションにアクセス トークンを取得できるようにします。間接的に、これはアドインが、ユーザーに二回目のサインインを必要とさせず、サインイン中のユーザーの Microsoft のグラフのデータにアクセスすることを可能とします。</span><span class="sxs-lookup"><span data-stu-id="6e792-p127">The Office Auth namespace, `Office.context.auth`, provides a method, `getAccessTokenAsync` that enables the Office host to obtain an access token to the add-in's web application. Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

<span data-ttu-id="6e792-p128">メソッドは、アドインの web アプリケーションにアクセス トークンを取得するため、 Azure Active Directory V 2.0 エンドポイントを呼び出します。これにより、アドインを使用してユーザーを識別できます。サーバー側コードでは、 [ "on behalf of"  OAuth フロー](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)を使用して、アドインの Web アプリケーションのMicrosoft Graphにアクセスするのに、このトークンを使用できます。</span><span class="sxs-lookup"><span data-stu-id="6e792-p128">The method calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application. This enables add-ins to identify users. Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span>

> [!NOTE]
> <span data-ttu-id="6e792-234">Outlook では、アドインが Outlook.com または Gmail のメールボックスに読み込まれている場合は、この API はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="6e792-234">[!Note In Outlook, this API is not supported if the add-in is loaded in an Outlook.com or Gmail mailbox.]</span></span>

<table><tr><td><span data-ttu-id="6e792-235">Hosts</span><span class="sxs-lookup"><span data-stu-id="6e792-235">Hosts</span></span></td><td><span data-ttu-id="6e792-236">Excel、OneNote、Outlook、PowerPoint、Word</span><span class="sxs-lookup"><span data-stu-id="6e792-236">Excel, Outlook, PowerPoint, Word</span></span></td></tr>

 <tr><td>[<span data-ttu-id="6e792-237">要件セット</span><span class="sxs-lookup"><span data-stu-id="6e792-237">Requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td><td>[<span data-ttu-id="6e792-238">IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="6e792-238">IdentityAPI</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)</td></tr></table>

#### <a name="parameters"></a><span data-ttu-id="6e792-239">パラメータ</span><span class="sxs-lookup"><span data-stu-id="6e792-239">Parameters</span></span>

<span data-ttu-id="6e792-p129">`options` -省略可能です。サインオン時の動作を定義するために、`AuthOptions` オブジェクト (下記参照) を受け入れます。</span><span class="sxs-lookup"><span data-stu-id="6e792-p129">`options` - Optional. Accepts an `AuthOptions` object (see below) to define sign-on behaviors.</span></span>

<span data-ttu-id="6e792-p130">`callback` -省略可能です。ユーザーの ID のトークンを解析したり、Microsoft のグラフへのアクセスを取得する "on behalf of"  フローでトークンを使用するコールバック メソッドを受け取ります。[AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` が「成功した」場合、 `AsyncResult.value` は、  生のAAD v. 2.0 形式のアクセス トークンです。</span><span class="sxs-lookup"><span data-stu-id="6e792-p130">`callback` - Optional. Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph. If [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v. 2.0-formatted access token.</span></span>

<span data-ttu-id="6e792-p131">`AuthOptions` インターフェイスは、Office が`getAccessTokenAsync` メソッドでAAD v からアドインへのアクセス トークンを取得する場合に、ユーザー エクスペリエンスのオプションを提供します。</span><span class="sxs-lookup"><span data-stu-id="6e792-p131">The `AuthOptions` interface provides options for the user experience when Office obtains an access token to the add-in from AAD v. 2.0 with the `getAccessTokenAsync` method.</span></span>

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



