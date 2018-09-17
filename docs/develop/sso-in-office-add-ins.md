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
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a><span data-ttu-id="b70a2-102">Office アドインのシングル サインオンを有効化する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="b70a2-102">Enable single sign-on for Office Add-ins (preview)</span></span>

<span data-ttu-id="b70a2-103">ユーザーは個人用の Microsoft アカウントまたは職場や学校の (Office 365) アカウントのいずれかを使用して、Office (オンライン、モバイル、およびデスクトップ プラットフォーム) にサインインします。</span><span class="sxs-lookup"><span data-stu-id="b70a2-103">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account.</span></span> <span data-ttu-id="b70a2-104">これを利用して、シングル サインオン (SSO) を使用すれば、ユーザーに 2 度目のサインインを求めなくても、ユーザーにアドインの使用を承認できます。</span><span class="sxs-lookup"><span data-stu-id="b70a2-104">You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>


![アドインのサインインプロセスを示す画像](../images/office-host-title-bar-sign-in.png)

> [!NOTE]
> <span data-ttu-id="b70a2-106">現在、シングル サインオン API は Word、Excel、Outlook、PowerPoint のプレビューでサポートされています。</span><span class="sxs-lookup"><span data-stu-id="b70a2-106">The Single Sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="b70a2-107">シングル サインオン API の現在のサポート状態に関する詳細は、「[IdentityAPI の要件セット](https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets?view=office-js)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b70a2-107">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets?view=office-js).</span></span>
> <span data-ttu-id="b70a2-108">SSO を使用するには、アドインの HTML 起動ページの https://appsforoffice.microsoft.com/lib/beta/hosted/office.js からベータ版 Office の JavaScript ライブラリを読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="b70a2-108">To use SSO, you must load the beta version of the Office JavaScript Library from https://appsforoffice.microsoft.com/lib/beta/hosted/office.js in the startup HTML page of the add-in.</span></span>
> <span data-ttu-id="b70a2-109">Outlook アドインで作業している場合は、Office 365 テナントの先進認証が有効になっていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="b70a2-109">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="b70a2-110">この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b70a2-110">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="b70a2-111">ユーザーにとっては、サインインが一度だけになり、アドインの実行エクスペリエンスがスムーズになります。</span><span class="sxs-lookup"><span data-stu-id="b70a2-111">For users, this makes running your add-in a smooth experience that involves at signing in only once.</span></span> <span data-ttu-id="b70a2-112">つまり開発者にとっては、暗号化されたパスワードを使った独自のユーザー テーブルをアドインに残さなくても良いということです。</span><span class="sxs-lookup"><span data-stu-id="b70a2-112">For developers, this means that your add-in does not have to maintain it's own user tables with encrypted passwords.</span></span>

### <a name="how-it-works-at-runtime"></a><span data-ttu-id="b70a2-113">実行時の動作のしくみ</span><span class="sxs-lookup"><span data-stu-id="b70a2-113">How it works at runtime</span></span>

<span data-ttu-id="b70a2-114">次の図は、SSO の動作のしくみを示しています。</span><span class="sxs-lookup"><span data-stu-id="b70a2-114">The following diagram shows how the SSO process works.</span></span>

![SSO プロセスを示す図](../images/sso-overview-diagram.png)

1. <span data-ttu-id="b70a2-116">アドインでは、JavaScript は新しい Office.js API [getAccessTokenAsync](#sso-api-reference) を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-116">In the add-in, JavaScript calls a new Office.js API [](#sso-api-reference).</span></span> <span data-ttu-id="b70a2-117">これにより、Office ホスト アプリケーションにアドインへのアクセス トークンを取得するように指示します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-117">This tells the Office host application to obtain an access token to the add-in.</span></span> <span data-ttu-id="b70a2-118">[アクセス トークンの例](#example-access-token) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b70a2-118">See [Example access token](#example-access-token).</span></span>
2. <span data-ttu-id="b70a2-119">ユーザーがサインインしていない場合、Office ホスト アプリケーションはユーザーにサインインを求めるポップアップ ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="b70a2-119">If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="b70a2-120">現在のユーザーが初めてアドインを使用する場合は、そのユーザーに同意を求めるダイアログを表示します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-120">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="b70a2-121">Office ホスト アプリケーションは、Azure AD v2.0 エンドポイントから現在のユーザーの**アドイン トークン**を要求します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-121">The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="b70a2-122">Azure AD は、Office ホスト アプリケーションにアドイン トークンを送信します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-122">Azure AD sends the add-in token to the Office host application.</span></span>
6. <span data-ttu-id="b70a2-123">Office ホスト アプリケーションは、`getAccessTokenAsync` 呼び出しによって返される結果オブジェクトの一部として、アドインに**アドイン トークン**を送信します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-123">The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessTokenAsync` call.</span></span>
7. <span data-ttu-id="b70a2-124">アドインの JavaScript は、トークンを解析し、必要な情報 (ユーザーの電子メールアドレスなど) を抽出できます。</span><span class="sxs-lookup"><span data-stu-id="b70a2-124">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span> 
8. <span data-ttu-id="b70a2-125">アプションで、アドインはサーバー側に HTTP 要求を送信して、ユーザーに関する詳細なデータ (ユーザーの嗜好など) を得ることができます。</span><span class="sxs-lookup"><span data-stu-id="b70a2-125">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences.</span></span> <span data-ttu-id="b70a2-126">もしくは、アクセス トークン自体をサーバー側に送信して、解析と検証をすることもできます。</span><span class="sxs-lookup"><span data-stu-id="b70a2-126">Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span> 

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="b70a2-127">SSO アドインの開発</span><span class="sxs-lookup"><span data-stu-id="b70a2-127">Develop an SSO add-in</span></span>

<span data-ttu-id="b70a2-128">このセクションでは、SSO を使用する Office アドインの作成に関連するタスクについて説明します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-128">This section describes the tasks involved in creating an Office Add-in that uses SSO.</span></span> <span data-ttu-id="b70a2-129">ここでは、これらのタスクについて、言語とフレームワークに依存しない方法で説明しています。</span><span class="sxs-lookup"><span data-stu-id="b70a2-129">These tasks are described here in a language- and framework-agnostic way.</span></span> <span data-ttu-id="b70a2-130">詳細なチュートリアルの例については、次を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b70a2-130">For examples of detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="b70a2-131">シングル サインオンを使用する Node.js Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="b70a2-131">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="b70a2-132">シングル サインオンを使用する ASP.NET Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="b70a2-132">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a><span data-ttu-id="b70a2-133">サービス アプリケーションを作成する</span><span class="sxs-lookup"><span data-stu-id="b70a2-133">Create the service application</span></span>

<span data-ttu-id="b70a2-134">Azure v2.0 エンドポイントの登録ポータルでアドインを登録します。https://apps.dev.microsoft.com。</span><span class="sxs-lookup"><span data-stu-id="b70a2-134">Register the add-in at the registration portal for the Azure v2.0 endpoint: https://apps.dev.microsoft.com. This is a 5–10 minute process that includes the following tasks:</span></span> <span data-ttu-id="b70a2-135">これは、次のタスクを含む 5 〜 10 分程度の時間がかかるプロセスです。</span><span class="sxs-lookup"><span data-stu-id="b70a2-135">This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="b70a2-136">アドインのクライアント ID とシークレットを取得します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-136">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="b70a2-137">アドインが AAD v に必要とするアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-137">Specify the permissions that your add-in needs to Microsoft Graph.</span></span> <span data-ttu-id="b70a2-138">(オプションで Microsoft Graph)。</span><span class="sxs-lookup"><span data-stu-id="b70a2-138">2.0 endpoint (and optionally to Microsoft Graph).</span></span> <span data-ttu-id="b70a2-139">[プロファイル] のアクセス許可は、常に必要です。</span><span class="sxs-lookup"><span data-stu-id="b70a2-139">The "profile" permission is always needed.</span></span>
* <span data-ttu-id="b70a2-140">Office ホスト アプリケーション信頼をアドインに付与します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-140">Grant the Office host application trust to the add-in.</span></span>
* <span data-ttu-id="b70a2-141">既定のアクセス許可 *access_as_user* を使用して、Office ホスト アプリケーションのアドインへのアクセスを事前認証します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-141">Preauthorize the Office host application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="b70a2-142">このプロセスの詳細については、「 [Azure AD v2.0 のエンドポイントで SSO を使用する Office アドインを登録する](register-sso-add-in-aad-v2.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="b70a2-142">For more information, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="b70a2-143">アドインを構成する</span><span class="sxs-lookup"><span data-stu-id="b70a2-143">Configure the add-in</span></span>

<span data-ttu-id="b70a2-144">新しいマークアップをアドイン マニフェストに追加します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-144">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="b70a2-145">**WebApplicationInfo** - 次の要素の親。</span><span class="sxs-lookup"><span data-stu-id="b70a2-145">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="b70a2-146">**Id** - アドインのクライアント ID。これは、アドイン登録の一環として取得するアプリケーション ID です。</span><span class="sxs-lookup"><span data-stu-id="b70a2-146">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in.</span></span> <span data-ttu-id="b70a2-147">「[Azure AD v2.0 のエンドポイントで SSO を使用する Office アドインを登録する](register-sso-add-in-aad-v2.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="b70a2-147">Details are at: [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="b70a2-148">**Resource** - アドインの URL。</span><span class="sxs-lookup"><span data-stu-id="b70a2-148">**Resource** - The URL of the add-in.</span></span>
* <span data-ttu-id="b70a2-149">**Scopes** - 1 つ以上の **Scope** 要素の親。</span><span class="sxs-lookup"><span data-stu-id="b70a2-149">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="b70a2-150">**Scope** - アドインが AAD に必要とするアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-150">**Scope** - Specifies a permission that the add-in needs to Microsoft Graph.</span></span> <span data-ttu-id="b70a2-151">アクセス許可は常に必要とされます。アドインに Microsoft Graph へのアクセス権がないとき、唯一必要な権限はこの許可である場合があります。`profile`</span><span class="sxs-lookup"><span data-stu-id="b70a2-151">The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph.</span></span> <span data-ttu-id="b70a2-152">このような場合、必要な Microsoft Graph へのアクセス許可を得るには **Scope** 要素も必要です。たとえば、 `User.Read`、 `Mail.Read` などです。</span><span class="sxs-lookup"><span data-stu-id="b70a2-152">If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`.</span></span> <span data-ttu-id="b70a2-153">Microsoft Graph にアクセスするためにコードで使用するライブラリには、アクセス許可がさらに必要な場合があります。</span><span class="sxs-lookup"><span data-stu-id="b70a2-153">Libraries that you use in your code to access Microsoft Graph may need additional permissions.</span></span> <span data-ttu-id="b70a2-154">たとえば、Microsoft 認証ライブラリ (MSAL) for .NET は、 `offline_access` アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="b70a2-154">For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission.</span></span> <span data-ttu-id="b70a2-155">詳細については、「 [Office アドインから Microsoft Graph を承認する](authorize-to-microsoft-graph.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b70a2-155">For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="b70a2-p111">Outlook 以外の Office ホストでは、`<VersionOverrides ... xsi:type="VersionOverridesV1_0">` セクションの末尾にマークアップを追加します。Outlook では、`<VersionOverrides ... xsi:type="VersionOverridesV1_1">` セクションの末尾にマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-p111">For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="b70a2-158">マークアップの例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-158">The following is an example of the markup:</span></span>

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

### <a name="add-client-side-code"></a><span data-ttu-id="b70a2-159">クライアント側のコードを追加する</span><span class="sxs-lookup"><span data-stu-id="b70a2-159">Add client-side code</span></span>

<span data-ttu-id="b70a2-160">アドインに JavaScript を追加します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-160">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="b70a2-161"> [GetAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-161">Call [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span></span>

* <span data-ttu-id="b70a2-162">アクセストークンを解析するか、アドインのサーバー側コードに渡すかします。</span><span class="sxs-lookup"><span data-stu-id="b70a2-162">Parse the access token or pass it to the add-in’s server-side code.</span></span> 

<span data-ttu-id="b70a2-163">ここでは、`getAccessTokenAsync` の呼び出しの簡単な例を示します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-163">Here's a simple example of a call to `getAccessTokenAsync`.</span></span> 

> [!NOTE]
> <span data-ttu-id="b70a2-164">この例では、1 種類のエラーのみを明示的に処理しています。</span><span class="sxs-lookup"><span data-stu-id="b70a2-164">This example handles only one kind of error explicitly.</span></span> <span data-ttu-id="b70a2-165">より詳細なエラー処理の例については、「[Office-Add-in-ASPNET-SSO の Home.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js)」および「[Office-Add-in-NodeJS-SSO の program.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b70a2-165">For examples of more elaborate error handling, see [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) and [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span></span> <span data-ttu-id="b70a2-166">また、「[シングル サインオン (SSO) のエラー メッセージのトラブルシューティング](troubleshoot-sso-in-office-add-ins.md)」 もご覧ください。</span><span class="sxs-lookup"><span data-stu-id="b70a2-166">Troubleshoot error messages for single sign-on (SSO)</span></span>
 

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

<span data-ttu-id="b70a2-167">ここでは、アドイン トークンをサーバー側に渡す簡単な例を示します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-167">Here's a simple example of passing the add-in token to the server-side.</span></span> <span data-ttu-id="b70a2-168">サーバー側に要求を送り返すときは、トークンが `Authorization` ヘッダーとして含まれます。</span><span class="sxs-lookup"><span data-stu-id="b70a2-168">The token is included as an `Authorization` header when sending a request back to the server-side.</span></span> <span data-ttu-id="b70a2-169">この例では、JSON データを送信すると想定しています。そのため、`POST` メソッドを使用しますが、サーバーへの書き込みを行っていない場合、アクセス トークンの送信には `GET` で十分です。</span><span class="sxs-lookup"><span data-stu-id="b70a2-169">This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

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

#### <a name="when-to-call-the-method"></a><span data-ttu-id="b70a2-170">メソッドを呼び出すタイミング</span><span class="sxs-lookup"><span data-stu-id="b70a2-170">When to call the method</span></span>

<span data-ttu-id="b70a2-171">Office にログインしているユーザーがなく、アドインを使用できない場合には、*アドインを起動するときに* `getAccessTokenAsync` を呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="b70a2-171">If your add-in cannot be used when a no user is logged into Office and Office does not have an access token to your add-in, then you should call `getAccessTokenAsync` *when the add-in launches*.</span></span>

<span data-ttu-id="b70a2-172">アドインにログイン ユーザーを必要としない機能がある場合、*ユーザーがログイン ユーザーを必要とするアクションを実行するときに* `getAccessTokenAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-172">If the add-in has some functionality that doesn't require access to Microsoft Graph or even a logged in user, then you call `getAccessTokenAsync` *when the user takes an action that requires access to Microsoft Graph or, at least, a logged in user*.</span></span> <span data-ttu-id="b70a2-173">`getAccessTokenAsync`  の重複呼び出しによってパフォーマンスが大幅に低下することはありません。これは、Office ではアクセス トークンがキャッシュされ、それが期限切れになるまで、AAD V が再度呼び出されることなく、再利用されるためです。</span><span class="sxs-lookup"><span data-stu-id="b70a2-173">There is no significant performance degradation with redundant calls of `getAccessTokenAsync` because Office caches the access token and will reuse it, until it expires, without making another call to the AAD V. 2.0 endpoint whenever  is called.</span></span> <span data-ttu-id="b70a2-174">が呼び出されたときに毎回 AAD v. 2.0 エンドポイントを呼び出すことなく再利用するためです。`getAccessTokenAsync`</span><span class="sxs-lookup"><span data-stu-id="b70a2-174">2.0 endpoint whenever `getAccessTokenAsync` is called.</span></span> <span data-ttu-id="b70a2-175">このため、`getAccessTokenAsync` の呼び出しを、このトークンが必要とされる場所でアクションを開始するすべての関数とハンドラーに追加します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-175">So you can add calls of `getAccessTokenAsync` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="b70a2-176">サーバー側のコードを追加する</span><span class="sxs-lookup"><span data-stu-id="b70a2-176">Add server-side code</span></span>

<span data-ttu-id="b70a2-177">ほとんどの場合、アドインがサーバー側に渡してそこで使用しない場合は、アクセス トークンを取得してもあまり意味はありません。</span><span class="sxs-lookup"><span data-stu-id="b70a2-177">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="b70a2-178">アドインで実行できるサーバー側タスクの一部を次に示します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-178">Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="b70a2-179">トークンから抽出されたユーザーに関する情報を使用する 1 つ以上の Web API メソッドを作成します。たとえば、ホスト型データベース内のユーザーの嗜好を参照するメソッドです。</span><span class="sxs-lookup"><span data-stu-id="b70a2-179">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base.</span></span> <span data-ttu-id="b70a2-180">(以下の「 **ID として SSO トークンを使用する** 」を参照してください)。使用している言語とフレームワークによっては、作成する必要があるコードを簡略化できるライブラリを利用できる場合があります。</span><span class="sxs-lookup"><span data-stu-id="b70a2-180">(See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="b70a2-181">Microsoft Graph データを取得します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-181">Get Microsoft Graph data.</span></span> <span data-ttu-id="b70a2-182">サーバー側のコードでは、次に示す操作を実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b70a2-182">Your server-side code should do the following:</span></span>

    * <span data-ttu-id="b70a2-183">アクセストークンを検証する (以下の「**アクセストークンを検証する**」を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="b70a2-183">Validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="b70a2-184">アクセス トークン、ユーザーに関するメタデータ、アドインの認証情報 (ID とシークレット) を含む Azure AD v2.0 エンドポイントへの呼び出しを使用して "On-Behalf-Of" フローを開始します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-184">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the add-in access token, some metadata about the user, and the credentials of the add-in (its ID and secret).</span></span> <span data-ttu-id="b70a2-185">このコンテキストでは、アクセストークンはブートストラップトークンと呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="b70a2-185">In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="b70a2-186">"On-Behalf-Of" フローから返された新しいアクセス トークンをキャッシュします。</span><span class="sxs-lookup"><span data-stu-id="b70a2-186">Cache the new access token that is returned from the on-behalf-of flow.</span></span>
    * <span data-ttu-id="b70a2-187">新しいトークンを使用して Microsoft Graph からデータを取得します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-187">Get data from Microsoft Graph by using the MSG token.</span></span>

 <span data-ttu-id="b70a2-188">ユーザーの Microsoft Graph データへの承認済みアクセスを取得する方法の詳細については、「[Office アドインで Microsoft Graph を承認する](authorize-to-microsoft-graph.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b70a2-188">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="b70a2-189">アクセス トークンを検証する</span><span class="sxs-lookup"><span data-stu-id="b70a2-189">For more information, see Validate the access token.</span></span>

<span data-ttu-id="b70a2-190">Web API でアクセス トークンを受信したら、そのアクセス トークンを使用する前に検証する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b70a2-190">Once the Web API receives the access token, it must validate it before using it.</span></span> <span data-ttu-id="b70a2-191">このトークンは、JSON Web トークン (JWT) です。そのため、この検証は最も標準的な OAuth でのトークンの検証とまったく同様に動作します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-191">The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows.</span></span> <span data-ttu-id="b70a2-192">JWT の検証を処理できるライブラリが複数入手可能ですが、その基本は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="b70a2-192">There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="b70a2-193">トークンが整形式であることを確認する</span><span class="sxs-lookup"><span data-stu-id="b70a2-193">Checking that the token is well-formed</span></span>
- <span data-ttu-id="b70a2-194">トークンが意図した証明機関から発行されたことを確認する</span><span class="sxs-lookup"><span data-stu-id="b70a2-194">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="b70a2-195">トークンが Web API を対象にしていることを確認する</span><span class="sxs-lookup"><span data-stu-id="b70a2-195">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="b70a2-196">トークンの検証時には、次のガイドラインに注意してください。</span><span class="sxs-lookup"><span data-stu-id="b70a2-196">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="b70a2-197">有効な SSO トークンは Azure 証明機関 `https://login.microsoftonline.com` から発行されます。</span><span class="sxs-lookup"><span data-stu-id="b70a2-197">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`.</span></span> <span data-ttu-id="b70a2-198">トークン内の `iss` クレームは、この値で始まっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="b70a2-198">The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="b70a2-199">トークンの `aud` パラメーターは、アドインの登録のアプリケーション ID に設定します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-199">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="b70a2-200">トークンの `scp` パラメーターは `access_as_user` に設定します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-200">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="b70a2-201">ID として SSO トークンを使用する</span><span class="sxs-lookup"><span data-stu-id="b70a2-201">Using the SSO token as an identity</span></span>

<span data-ttu-id="b70a2-202">アドインでユーザーの ID を検証する必要がある場合、SSO トークンには ID を確定するために使用できる情報が含まれています。</span><span class="sxs-lookup"><span data-stu-id="b70a2-202">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity.</span></span> <span data-ttu-id="b70a2-203">ID に関連するトークン内のクレームは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="b70a2-203">The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="b70a2-204">`name` : ユーザーの表示名。</span><span class="sxs-lookup"><span data-stu-id="b70a2-204">`name` - The user's display name.</span></span>
- <span data-ttu-id="b70a2-205">`preferred_username` ユーザーの電子メール アドレスです。</span><span class="sxs-lookup"><span data-stu-id="b70a2-205">`preferred_username`The user's email address.</span></span>
- <span data-ttu-id="b70a2-206">`oid` : Azure Active Directory でユーザーの ID を表す GUID。</span><span class="sxs-lookup"><span data-stu-id="b70a2-206">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="b70a2-207">`tid` : Azure Active Directory でユーザーの組織の ID を表す GUID。</span><span class="sxs-lookup"><span data-stu-id="b70a2-207">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="b70a2-208">`name` と `preferred_username` の値は変化することがあるため、`oid` と `tid` の値を ID とバックエンドの承認サービスを関連付けるために使用するようにしてください。</span><span class="sxs-lookup"><span data-stu-id="b70a2-208">Since the `name` and `preferred_username` values could change, it's recommended that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="b70a2-209">たとえば、サービスでは、これらの値を `{oid-value}@{tid-value}` のような形式にまとめて、内部ユーザー データベースのユーザーのレコードに値として保存できます。</span><span class="sxs-lookup"><span data-stu-id="b70a2-209">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database.</span></span> <span data-ttu-id="b70a2-210">その後の要求では、同じ値を使用してユーザーを取得できるようになり、特定のリソースへのアクセスについては既存のアクセス制御メカニズムに基づいて決定できます。</span><span class="sxs-lookup"><span data-stu-id="b70a2-210">Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="b70a2-211">アクセス トークンの例</span><span class="sxs-lookup"><span data-stu-id="b70a2-211">Example access token</span></span>

<span data-ttu-id="b70a2-212">アクセス トークンの標準的なデコードされたペイロードを次に示します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-212">The following is a typical decoded payload of an access token.</span></span> <span data-ttu-id="b70a2-213">プロパティの詳細については、「[Azure Active Directory v2.0 トークンリファレンス](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b70a2-213">For information about the properties, see [Azure Active Directory v2.0 tokens reference](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).</span></span>


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

## <a name="using-sso-with-an-outlook-add-in"></a><span data-ttu-id="b70a2-214">Outlook のアドインでの SSO の使用</span><span class="sxs-lookup"><span data-stu-id="b70a2-214">Using SSO with and Outlook add-in</span></span>

<span data-ttu-id="b70a2-215">SSO を Outlook アドインとして使用する場合と、Excel、PowerPoint、Word アドインとして使用する場合には、ささやかですが重要な違いがあります。</span><span class="sxs-lookup"><span data-stu-id="b70a2-215">There are some small, but important differences in using SSO in and Outlook add-in from using it in as Excel, PowerPoint, or Word add-in.</span></span> <span data-ttu-id="b70a2-216">「[Outlook アドインでシングルサインオントークンを使用してユーザーを認証する](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)」と「[シナリオ: Outlook アドインでサービスにシングルサインオンを実装する](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)」を必ず読んでください。</span><span class="sxs-lookup"><span data-stu-id="b70a2-216">Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) and [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

## <a name="sso-api-reference"></a><span data-ttu-id="b70a2-217">SSO API リファレンス</span><span class="sxs-lookup"><span data-stu-id="b70a2-217">SSO API reference</span></span>

### <a name="getaccesstokenasync"></a><span data-ttu-id="b70a2-218">getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="b70a2-218">getAccessTokenAsync</span></span>

<span data-ttu-id="b70a2-219">Office の認証の名前空間 `Office.context.auth` は、Office ホストがアドインの Web アプリケーションへのアクセス トークンを取得できるようにするメソッド `getAccessTokenAsync` を提供します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-219">The Office Auth namespace, `Office.context.auth`, provides a method, `getAccessTokenAsync` that enables the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="b70a2-220">これはまた間接的に、ユーザーがもう一度サインインする必要なしに、アドインがサインインしたユーザーの Microsoft Graph データにアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="b70a2-220">Indirectly, enable the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

<span data-ttu-id="b70a2-221">このメソッドは、Azure Active Directory V 2.0 のエンドポイントを呼び出して、アドインの Web アプリケーションへのアクセス トークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="b70a2-221">Calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application.</span></span> <span data-ttu-id="b70a2-222">これにより、アドインを使用してユーザーを識別できます。</span><span class="sxs-lookup"><span data-stu-id="b70a2-222">This enables add-ins to identify users.</span></span> <span data-ttu-id="b70a2-223">["on behalf of" OAuth フロー](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)を使用することにより、サーバー側のコードはこのトークンを使用して、アドインの Web アプリケーションの Microsoft Graph にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="b70a2-223">Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span>

> [!NOTE]
> [!Outlook メモによれば、アドインが Outlook.com または Gmail のメールボックスに読み込まれている場合は、この API はサポートされていません。]

<table><tr><td><span data-ttu-id="b70a2-225">Hosts</span><span class="sxs-lookup"><span data-stu-id="b70a2-225">Hosts</span></span></td><td><span data-ttu-id="b70a2-226">Excel、OneNote、Outlook、PowerPoint、Word</span><span class="sxs-lookup"><span data-stu-id="b70a2-226">Excel, Outlook, PowerPoint, Word</span></span></td></tr>

 <tr><td><span data-ttu-id="b70a2-227">要件セット</span><span class="sxs-lookup"><span data-stu-id="b70a2-227">Requirement sets</span></span></td><td>[<span data-ttu-id="b70a2-228">IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="b70a2-228">IdentityAPI</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td></tr></table>

#### <a name="parameters"></a><span data-ttu-id="b70a2-229">パラメーター</span><span class="sxs-lookup"><span data-stu-id="b70a2-229">Parameters</span></span>

<span data-ttu-id="b70a2-230">`options` - 省略可能。</span><span class="sxs-lookup"><span data-stu-id="b70a2-230">`options` - Optional.</span></span> <span data-ttu-id="b70a2-231">サインオン時の動作を定義するために、`AuthOptions` オブジェクト (下記参照) を受け入れます。</span><span class="sxs-lookup"><span data-stu-id="b70a2-231">Accepts an `AuthOptions` object (see below) to define sign-on behaviors.</span></span>

<span data-ttu-id="b70a2-232">`callback` - 省略可能。</span><span class="sxs-lookup"><span data-stu-id="b70a2-232">`callback` - Optional.</span></span> <span data-ttu-id="b70a2-233">ユーザーの ID のトークンを解析したり、"On-Behalf-Of" フローのトークンを使用して Microsoft Graph へのアクセスを取得することができるコールバック メソッドを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="b70a2-233">Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph.</span></span> <span data-ttu-id="b70a2-234">[AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` が「成功した」場合は、 `AsyncResult.value` は生の AAD v です。</span><span class="sxs-lookup"><span data-stu-id="b70a2-234">If [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v.</span></span> <span data-ttu-id="b70a2-235">2.0 形式のアクセス トークンです。</span><span class="sxs-lookup"><span data-stu-id="b70a2-235">2.0-formatted access token.</span></span>

<span data-ttu-id="b70a2-236">インターフェイスは、Office が AAD v からアドインへのアクセス トークンを取得する場合に、ユーザー エクスペリエンスのオプションを提供します。`AuthOptions`</span><span class="sxs-lookup"><span data-stu-id="b70a2-236">The `AuthOptions` interface provides options for the user experience when Office obtains an access token to the add-in from AAD v.</span></span> <span data-ttu-id="b70a2-237">メソッドを持つ 2.0 です。`getAccessTokenAsync`</span><span class="sxs-lookup"><span data-stu-id="b70a2-237">2.0 with the `getAccessTokenAsync` method.</span></span>

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



