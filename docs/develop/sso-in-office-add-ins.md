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
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a><span data-ttu-id="0f56f-102">Office アドインのシングル サインオンを有効にする (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="0f56f-102">Enable single sign-on for Office Add-ins (preview)</span></span>

<span data-ttu-id="0f56f-p101">ユーザーは、個人用の Microsoft アカウントまたは職場や学校の (Office 365) アカウントのいずれかを用いて Office (オンライン、モバイル、デスクトップ プラットフォーム) にサインインします。この特性を生かして、ユーザーに2 回目のサインインを要求しないよう、アドインへユーザーを承認するためのシングル サインオン (SSO) を使用できます。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p101">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account. You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>

![アドインのサインイン プロセスを示す画像](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a><span data-ttu-id="0f56f-106">プレビュー ステータス</span><span class="sxs-lookup"><span data-stu-id="0f56f-106">Preview Status</span></span>

<span data-ttu-id="0f56f-p102">シングル サインオンの API は、現在プレビュー版のみがサポートされています。開発者は検証用に使用できますが、製品版のアドインで使用できません。また SSO を使用するアドインは、 [AppSource](https://appsource.microsoft.com) には対応していません。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p102">The Single Sign-on API is currently supported in preview only. It is available to developers for experimentation; but it should not be used in a production add-in. In addition, add-ins that use SSO are not accepted in [AppSource](https://appsource.microsoft.com).</span></span>

<span data-ttu-id="0f56f-p103">SSO のプレビュー版は、すべての Office アプリケーションでサポートしているわけではありません。使用できるのは、Word、Excel、Outlook、PowerPoint です。シングル サインオンの API のサポート状況の詳細については、「 [IdentityAPI 要件の設定](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p103">Not all Office applications support the SSO preview. It is available in Word, Excel, Outlook, and PowerPoint. For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js).</span></span>

### <a name="requirements-and-best-practices"></a><span data-ttu-id="0f56f-113">要件とベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="0f56f-113">Requirements and Best Practices</span></span>

<span data-ttu-id="0f56f-114">SSO を使用するには、アドインの HTML 起動ページの `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` から Office JavaScript ライブラリのベータ版を読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="0f56f-114">To use SSO, you must load the beta version of the Office JavaScript Library from `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` in the startup HTML page of the add-in.</span></span>

<span data-ttu-id="0f56f-p104"> *\*Outlook** アドインで SSO を使用するには、Office 365 テナントの先進認証を有効にする必要があります。詳細な方法については、「 [Exchange Online: テナントで先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p104">If you are working with an **Outlook** add-in, be sure to enable Modern Authentication for the Office 365 tenancy. For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="0f56f-p105">アドインの唯一の認証方法として SSOを使用することは *おやめください* 。特定のエラー時に使用できる代替の認証システムを実装する必要があります。ユーザー テーブルおよび認証システム、あるいはソーシャル ログイン プロバイダーのいずれかを使用できます。Office アドインでこれを行う方法は、「[Office アドインで外部サービスを承認する](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins)」を参照してください。*Outlook*では、推奨されるフォールバック システムがあります。詳細は「[シナリオ: Outlook のアドインで、サービスへのシングル サインオンを実装する](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p105">You should *not* rely on SSO as your add-in's only method of authentication. You should implement an alternate authentication system that your add-in can fall back to in certain error situations. You can use a system of user tables and authentication, or you can leverage one of the social login providers. For more information about how to do this with an Office add-in, see [Authorize external services in your Office Add-in](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins). For *Outlook*, there is a recommended fall back system. For more information, see [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

### <a name="how-sso-works-at-runtime"></a><span data-ttu-id="0f56f-123">実行時に SSO が動作する仕組み</span><span class="sxs-lookup"><span data-stu-id="0f56f-123">How it works at runtime</span></span>

<span data-ttu-id="0f56f-124">次の図は、SSO の動作の仕組みを示しています。</span><span class="sxs-lookup"><span data-stu-id="0f56f-124">The following diagram shows how the SSO process works.</span></span>

![SSO プロセスを示す図](../images/sso-overview-diagram.png)

1. <span data-ttu-id="0f56f-p106">アドインでは、JavaScript は新しい Office.js API [getAccessTokenAsync](#sso-api-reference) を呼び出します。これにより、Office ホスト アプリケーションにアドインへのアクセス トークンを取得するように指示します。「[アクセス トークンの使用例](#example-access-token)」 をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p106">In the add-in, JavaScript calls a new Office.js API [getAccessTokenAsync](#sso-api-reference). This tells the Office host application to obtain an access token to the add-in. See [Example access token](#example-access-token).</span></span>
2. <span data-ttu-id="0f56f-129">ユーザーがサインインしていない場合、Office ホスト アプリケーションはユーザーにサインインを求めるポップアップ ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="0f56f-129">If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="0f56f-130">現在のユーザーが初めてアドインを使用する場合は、そのユーザーの同意を求めるダイアログを表示します。</span><span class="sxs-lookup"><span data-stu-id="0f56f-130">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="0f56f-131">Office ホスト アプリケーションは、Azure AD v2.0 エンドポイントに現在のユーザーの**アドイン トークン**を要求します。</span><span class="sxs-lookup"><span data-stu-id="0f56f-131">The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="0f56f-132">Azure AD は、Office ホスト アプリケーションにアドイン トークンを送信します。</span><span class="sxs-lookup"><span data-stu-id="0f56f-132">Azure AD sends the add-in token to the Office host application.</span></span>
6. <span data-ttu-id="0f56f-133">Office ホスト アプリケーションは、`getAccessTokenAsync` 呼び出しによって返される結果オブジェクトの一部として、アドインに**アドイン トークン**を送信します。</span><span class="sxs-lookup"><span data-stu-id="0f56f-133">The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessTokenAsync` call.</span></span>
7. <span data-ttu-id="0f56f-134">アドインの JavaScript は、トークンを解析し、必要な情報 (ユーザーの電子メールアドレスなど) を抽出できます。</span><span class="sxs-lookup"><span data-stu-id="0f56f-134">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span> 
8. <span data-ttu-id="0f56f-p107">オプションで、アドインはサーバー側に HTTP 要求を送信して、ユーザーに関する詳細なデータ (ユーザーの嗜好など) を得ることができます。また、自身のアクセス トークンを解析および検証用にサーバー側に送信することもできます。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p107">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences. Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span> 

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="0f56f-137">SSO アドインの開発</span><span class="sxs-lookup"><span data-stu-id="0f56f-137">Develop an SSO add-in</span></span>

<span data-ttu-id="0f56f-p108">このセクションでは、SSO を使用する Office アドインを作成するための作業について説明します。これらのタスクについては、言語とフレームワークにとらわれない方法でここで説明します。詳細なチュートリアルの例については、次のように表示されます。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p108">This section describes the tasks involved in creating an Office Add-in that uses SSO. These tasks are described here in a language- and framework-agnostic way. For examples of detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="0f56f-141">シングル サインオンを使用する Node.js Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="0f56f-141">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="0f56f-142">シングル サインオンを使用する ASP.NET Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="0f56f-142">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a><span data-ttu-id="0f56f-143">サービス アプリケーションを作成する</span><span class="sxs-lookup"><span data-stu-id="0f56f-143">Create the service application</span></span>

<span data-ttu-id="0f56f-p109">Azure v2.0 エンドポイントの登録ポータルでアドインを登録します。 https://apps.dev.microsoft.comこのプロセスには、次に示すタスクを含めて 5 分から 10 分の時間がかかります。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p109">Register the add-in at the registration portal for the Azure v2.0 endpoint: https://apps.dev.microsoft.com. This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="0f56f-146">アドインのクライアント ID とシークレットを取得します。</span><span class="sxs-lookup"><span data-stu-id="0f56f-146">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="0f56f-p110">AAD v. 2.0 エンドポイント  (および必要に応じて Microsoft Graph) で必要とするアクセス許可を指定します。「プロファイル」のアクセス許可は、常に必要です。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p110">Specify the permissions that your add-in needs to AAD v. 2.0 endpoint (and optionally to Microsoft Graph). The "profile" permission is always needed.</span></span>
* <span data-ttu-id="0f56f-150">Office ホスト アプリケーション信頼をアドインに付与します。</span><span class="sxs-lookup"><span data-stu-id="0f56f-150">Grant the Office host application trust to the add-in.</span></span>
* <span data-ttu-id="0f56f-151">既定のアクセス許可 *access_as_user* を使用して、Office ホスト アプリケーションのアドインへのアクセスを事前認証します。</span><span class="sxs-lookup"><span data-stu-id="0f56f-151">Preauthorize the Office host application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="0f56f-152">このプロセスの詳細については、「 [Azure AD v2.0 のエンドポイントで SSO を使用する Office アドインを登録する](register-sso-add-in-aad-v2.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="0f56f-152">For more information, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="0f56f-153">アドインを構成する</span><span class="sxs-lookup"><span data-stu-id="0f56f-153">Configure the add-in</span></span>

<span data-ttu-id="0f56f-154">アドイン マニフェストに新しいマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="0f56f-154">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="0f56f-155">**WebApplicationInfo** - 次の要素の親。</span><span class="sxs-lookup"><span data-stu-id="0f56f-155">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="0f56f-p111">**ID** -アドインのクライアント ID。これは、アドイン登録の一環として取得するアプリケーション IDです。「 [Azure AD v2.0 のエンドポイントで SSO を使用している Office アドインの登録](register-sso-add-in-aad-v2.md) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p111">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in. See [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="0f56f-158">**Resource** - アドインの URL。</span><span class="sxs-lookup"><span data-stu-id="0f56f-158">**Resource** - The URL of the add-in.</span></span>
* <span data-ttu-id="0f56f-159">**スコープ** - 1 つ以上の **スコープ** 要素の親。</span><span class="sxs-lookup"><span data-stu-id="0f56f-159">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="0f56f-p112"> *\*スコープ**  - AAD に必要な追加のアクセス許可を指定します。 `profile` 許可 は、常に必要で、アドインがアクセスしていない Microsoft Graph に対して、必要な唯一のアクセス許可がある可能性があります。その場合は、また\*\*  スコープ\*\* 要素が Microsoft Graph に必要なアクセス許可です。たとえば、 `User.Read\`、 `Mail.Read\`です。グラフにアクセスするためにコード内で使用するライブラリには、追加のアクセス許可が必要です。たとえば、、Microsoft 認証ライブラリ (MSAL) .NET は `offline_access` のアクセス許可を必要とします。詳細については、 [Office アドインから Graph に承認](authorize-to-microsoft-graph.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p112">**Scope** - Specifies a permission that the add-in needs to AAD. The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph. If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`. Libraries that you use in your code to access Microsoft Graph may need additional permissions. For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission. For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="0f56f-p113">Outlook 以外の Office ホストでは、 `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` セクションの末尾にマークアップを追加します。Outlook では、 `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` セクションの末尾にマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p113">For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="0f56f-168">マークアップの例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="0f56f-168">The following is an example of the markup:</span></span>

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

### <a name="add-client-side-code"></a><span data-ttu-id="0f56f-169">クライアント側のコードを追加する</span><span class="sxs-lookup"><span data-stu-id="0f56f-169">Add client-side code</span></span>

<span data-ttu-id="0f56f-170">アドインに JavaScript を追加します。</span><span class="sxs-lookup"><span data-stu-id="0f56f-170">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="0f56f-171">[GetAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="0f56f-171">Call [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span></span>

* <span data-ttu-id="0f56f-172">アクセストークンを解析するか、アドインのサーバー側コードに渡します。</span><span class="sxs-lookup"><span data-stu-id="0f56f-172">Parse the access token or pass it to the add-in’s server-side code.</span></span> 

<span data-ttu-id="0f56f-173"> `getAccessTokenAsync` の呼び出しの簡単な例を示します。</span><span class="sxs-lookup"><span data-stu-id="0f56f-173">Here's a simple example of a call to `getAccessTokenAsync`.</span></span> 

> [!NOTE]
> <span data-ttu-id="0f56f-p114">この例は、一種類のみのエラーを処理しているものです。より複雑なエラー処理の例については、「 [Office-Add-in-ASPNET-SSO の Home.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) 」および「 [Office-Add-in-NodeJS-SSO の program.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js) 」を参照してください。また、「 [シングル サインオン (SSO) のエラー メッセージのトラブルシューティング](troubleshoot-sso-in-office-add-ins.md) 」もご覧ください。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p114">This example handles only one kind of error explicitly. For examples of more elaborate error handling, see [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) and [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js). And see [Troubleshoot error messages for single sign-on (SSO)](troubleshoot-sso-in-office-add-ins.md).</span></span>
 

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

<span data-ttu-id="0f56f-p115">以下は、サーバー側に追加のトークンを渡すことの簡単な例です。トークンは、サーバー側に要求を送信するときに `Authorization` ヘッダーの一つとして含まれます。この例では JSON データを送信しているため `POST` メソッドを使用していますが、サーバーへの書き込みを行わない場合は、アクセス トークンの送信に `GET` だけで十分です。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p115">Here's a simple example of passing the add-in token to the server-side. The token is included as an `Authorization` header when sending a request back to the server-side. This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

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

#### <a name="when-to-call-the-method"></a><span data-ttu-id="0f56f-180">メソッドを呼び出すタイミング</span><span class="sxs-lookup"><span data-stu-id="0f56f-180">When to call the method</span></span>

<span data-ttu-id="0f56f-181">Office にログインしているユーザーがなく、アドインを使用できない場合には、*アドインを起動するときに* `getAccessTokenAsync` を呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="0f56f-181">If your add-in cannot be used when a no user is logged into Office and Office does not have an access token to your add-in, then you should call `getAccessTokenAsync` *when the add-in launches*.</span></span>

<span data-ttu-id="0f56f-p116">アドインがログインを必要としないいくつかの機能を持っている場合、\* ユーザーがログイン中のユーザーを必要とするアクションを実行する際\*  、 `getAccessTokenAsync` を呼び出します 。 `getAccessTokenAsync` の重複呼び出しによってパフォーマンスgが大幅に低下することはありません。これは、Office ではアクセス トークンがキャッシュされ、それが期限切れになるまで、 `getAccessTokenAsync` が呼び出されても AAD v. 2.0 エンドポイントを再度呼び出すことなく再利用するためです。そのため `getAccessTokenAsync` の呼び出しを、すべての関数と、トークンを必要とするアクションを開始するハンドラへ追加することができます。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p116">If the add-in has some functionality that doesn't require a logged in user, then you call `getAccessTokenAsync` *when the user takes an action that requires a logged in user*. There is no significant performance degradation with redundant calls of `getAccessTokenAsync` because Office caches the access token and will reuse it, until it expires, without making another call to the AAD v. 2.0 endpoint whenever `getAccessTokenAsync` is called. So you can add calls of `getAccessTokenAsync` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="0f56f-186">サーバー側のコードを追加する</span><span class="sxs-lookup"><span data-stu-id="0f56f-186">Add server-side code</span></span>

<span data-ttu-id="0f56f-p117">ほとんどのシナリオでは、アドインが取得しサーバーに渡したアクセス トークンをサーバー側で使用しない場合は、アクセス トークンを取得することに意味はありません。いくつかのサーバ側のタスクでは、アドインは以下のことが可能です。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p117">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there. Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="0f56f-p118">トークンから抽出されたユーザーに関する情報を使用する 1 つ以上の Web API のメソッドを作成します。たとえば、ホスト型データベース内のユーザーの嗜好を参照するメソッドです。(「 **SSO トークンを ID としてを使用する** 」を参照してください。) お使いの言語とフレームワークによっては、記述するコードを簡略化することのできるライブラリがある場合があります。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p118">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base. (See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="0f56f-p119">Microsoft Graph のデータを取得します。サーバー側コードは、次の操作を行います。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p119">Get Microsoft Graph data. Your server-side code should do the following:</span></span>

    * <span data-ttu-id="0f56f-193">アクセストークンを検証する (以下の「 **アクセストークンを検証する** 」を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="0f56f-193">Validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="0f56f-p120">アクセス トークン、ユーザー、およびアドイン (ID とシークレット) の資格情報に関するいくつかのメタデータが含まれている Azure AD v2.0  エンドポイントへの呼び出しに ”on behalf of” フローを開始します。ここでは、アクセス トークンは、ブートス トラップのトークンと呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p120">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the access token, some metadata about the user, and the credentials of the add-in (its ID and secret). In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="0f56f-196">"On-Behalf-Of" フローから返された新しいアクセス トークンをキャッシュします。</span><span class="sxs-lookup"><span data-stu-id="0f56f-196">Cache the new access token that is returned from the on-behalf-of flow.</span></span>
    * <span data-ttu-id="0f56f-197">新しいトークンを使用して Microsoft Graph からデータを取得します。</span><span class="sxs-lookup"><span data-stu-id="0f56f-197">Get data from Microsoft Graph by using the MSG token.</span></span>

 <span data-ttu-id="0f56f-198">ユーザーの Microsoft Graph データへの承認済みアクセスを取得する方法の詳細については、「[Office アドインで Microsoft Graph を承認する](authorize-to-microsoft-graph.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0f56f-198">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="0f56f-199">アクセス トークンを検証する</span><span class="sxs-lookup"><span data-stu-id="0f56f-199">For more information, see Validate the access token.</span></span>

<span data-ttu-id="0f56f-p121">Web API がアクセス トークンを受信した後は、使用前に検証を行う必要があります。このトークンは JSON Web トークン (JWT) であるため、その検証はほとんどの標準 OAuth のトークン検証と同様に動作します。JWT の検証を処理するライブラリはいくつかありますが、基本は以下のものです。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p121">Once the Web API receives the access token, it must validate it before using it. The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows. There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="0f56f-203">トークンが整形式であることを確認する</span><span class="sxs-lookup"><span data-stu-id="0f56f-203">Checking that the token is well-formed</span></span>
- <span data-ttu-id="0f56f-204">トークンが意図した証明機関から発行されたことを確認する</span><span class="sxs-lookup"><span data-stu-id="0f56f-204">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="0f56f-205">トークンが Web API を対象にしていることを確認する</span><span class="sxs-lookup"><span data-stu-id="0f56f-205">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="0f56f-206">トークンを検証するときは、次のガイドラインに留意してください。</span><span class="sxs-lookup"><span data-stu-id="0f56f-206">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="0f56f-p122">有効な SSO トークンは Azure 証明機関 `https://login.microsoftonline.com` から発行されます。トークン内の `iss` 要求は、この値で始まっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p122">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`. The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="0f56f-209">トークンの `aud` パラメータは、アドインの登録のアプリケーション ID に設定します。</span><span class="sxs-lookup"><span data-stu-id="0f56f-209">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="0f56f-210">トークンの `scp` パラメータは `access_as_user` に設定します。</span><span class="sxs-lookup"><span data-stu-id="0f56f-210">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="0f56f-211">ID として SSO トークンを使用する</span><span class="sxs-lookup"><span data-stu-id="0f56f-211">Using the SSO token as an identity</span></span>

<span data-ttu-id="0f56f-p123">アドインでユーザーの ID を検証する必要がある場合、SSO トークンには ID を確定するために使用できる情報が含まれています。以下は、ID に関連するトークンの要求です。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p123">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity. The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="0f56f-214">`name` - ユーザーの表示名。</span><span class="sxs-lookup"><span data-stu-id="0f56f-214">`name` - The user's display name.</span></span>
- <span data-ttu-id="0f56f-215">`preferred_username` - ユーザーの電子メール アドレス。</span><span class="sxs-lookup"><span data-stu-id="0f56f-215">`preferred_username`The user's email address.</span></span>
- <span data-ttu-id="0f56f-216">`oid` - Azure Active Directory でユーザーの ID を表す GUID。</span><span class="sxs-lookup"><span data-stu-id="0f56f-216">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="0f56f-217">`tid` - Azure Active Directory でユーザーの組織の ID を表す GUID。</span><span class="sxs-lookup"><span data-stu-id="0f56f-217">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="0f56f-218"> `name` と `preferred_username` の値は変化することがあるため、ID とバックエンドの承認サービスを関連付けるために `oid` と `tid` の値を使用するようにしてください。</span><span class="sxs-lookup"><span data-stu-id="0f56f-218">Since the `name` and `preferred_username` values could change, it's recommended that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="0f56f-p124">たとえば、サービスでは、これらの値を `{oid-value}@{tid-value}` のような形式にまとめて、内部ユーザー データベースのユーザーのレコードに値として保存できます。その後の要求では、同じ値を使用してユーザーを取得できるようになり、特定のリソースへのアクセスについては既存のアクセス制御メカニズムに基づいて決定できます。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p124">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database. Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="0f56f-221">アクセス トークンの例</span><span class="sxs-lookup"><span data-stu-id="0f56f-221">Example access token</span></span>

<span data-ttu-id="0f56f-p125">以下は、アクセス トークンの標準的なデコードされたペイロードです。プロパティの詳細については、「 [Azure Active Directory バージョン 2.0 のトークンの参照](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens) 」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p125">The following is a typical decoded payload of an access token. For information about the properties, see [Azure Active Directory v2.0 tokens reference](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).</span></span>


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

## <a name="using-sso-with-an-outlook-add-in"></a><span data-ttu-id="0f56f-224">Outlook アドインでの SSO の使用</span><span class="sxs-lookup"><span data-stu-id="0f56f-224">Using SSO with and Outlook add-in</span></span>

<span data-ttu-id="0f56f-p126">SSO を Outlook のアドインで使用する場合と、Excel、PowerPoint、または  Word のアドインで使用する場合とでは、小さくても重要な違いがあります。「[Outlook アドインでシングル サインオン トークンを使用してユーザーを認証する](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)」と「[シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)」を必ずお読みください。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p126">There are some small, but important differences in using SSO in an Outlook add-in from using it in an Excel, PowerPoint, or Word add-in. Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) and [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

## <a name="sso-api-reference"></a><span data-ttu-id="0f56f-227">SSO API 参照</span><span class="sxs-lookup"><span data-stu-id="0f56f-227">SSO API reference</span></span>

### <a name="getaccesstokenasync"></a><span data-ttu-id="0f56f-228">getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="0f56f-228">getAccessTokenAsync</span></span>

<span data-ttu-id="0f56f-p127">Office Auth の名前空間 `Office.context.auth` は、Office ホストがアドインの Web アプリケーションへのアクセス トークンを取得できるようにするメソッド `getAccessTokenAsync` を提供します。これにより間接的に、ユーザーが2回目のサインインを行うことなく、アドインがサインイン中のユーザーの Microsoft Graph のデータにアクセスすることも可能になります。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p127">The Office Auth namespace, `Office.context.auth`, provides a method, `getAccessTokenAsync` that enables the Office host to obtain an access token to the add-in's web application. Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

<span data-ttu-id="0f56f-p128">このメソッドは、アドインの Web アプリケーションにアクセス トークンを取得するため、 Azure Active Directory V 2.0 エンドポイントを呼び出します。これにより、アドインを使用してユーザーを識別できます。 ["on behalf of"  OAuth フロー](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of) を使用することにより、サーバー側のコードはこのトークンを使用して、アドインの Web アプリケーションの Microsoft Graph にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p128">The method calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application. This enables add-ins to identify users. Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span>

> [!NOTE]
> <span data-ttu-id="0f56f-234">Outlook では、アドインが Outlook.com または Gmail のメールボックスに読み込まれている場合は、この API はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0f56f-234">[!Note In Outlook, this API is not supported if the add-in is loaded in an Outlook.com or Gmail mailbox.]</span></span>

<table><tr><td><span data-ttu-id="0f56f-235">Hosts</span><span class="sxs-lookup"><span data-stu-id="0f56f-235">Hosts</span></span></td><td><span data-ttu-id="0f56f-236">Excel、OneNote、Outlook、PowerPoint、Word</span><span class="sxs-lookup"><span data-stu-id="0f56f-236">Excel, Outlook, PowerPoint, Word</span></span></td></tr>

 <tr><td><span data-ttu-id="0f56f-237">要件セット</span><span class="sxs-lookup"><span data-stu-id="0f56f-237">Requirement sets</span></span></td><td>[<span data-ttu-id="0f56f-238">IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="0f56f-238">IdentityAPI</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td></tr></table>

#### <a name="parameters"></a><span data-ttu-id="0f56f-239">パラメータ</span><span class="sxs-lookup"><span data-stu-id="0f56f-239">Parameters</span></span>

<span data-ttu-id="0f56f-p129">`options` - 任意。サインオン時の動作を定義するために、 `AuthOptions` オブジェクト (下記参照) を受け入れます。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p129">`options` - Optional. Accepts an `AuthOptions` object (see below) to define sign-on behaviors.</span></span>

<span data-ttu-id="0f56f-p130">`callback` - 任意。ユーザーの ID のトークンを解析したり、Microsoft Graph へのアクセスを取得する "on behalf of" フローでトークンを使用したりするコールバック メソッドを受け取ります。 [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` が「成功」した場合、 `AsyncResult.value` が AAD v. 2.0 形式の生アクセス トークンとなります。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p130">`callback` - Optional. Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph. If [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v. 2.0-formatted access token.</span></span>

<span data-ttu-id="0f56f-p131"> `AuthOptions` インターフェイスは、Office が `getAccessTokenAsync` メソッドで AAD v. 2.0 からアドインへのアクセス トークンを取得する場合に、ユーザー エクスペリエンスのオプションを提供します。</span><span class="sxs-lookup"><span data-stu-id="0f56f-p131">The `AuthOptions` interface provides options for the user experience when Office obtains an access token to the add-in from AAD v. 2.0 with the `getAccessTokenAsync` method.</span></span>

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



