---
title: Office アドインのシングル サインオンを有効にする
description: ''
ms.date: 09/26/2018
ms.openlocfilehash: 05b5088a61df3f77a09b60dbdc3129074d5f8530
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348171"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a><span data-ttu-id="f2dd1-102">Office アドインのシングル サインオンを有効にする (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="f2dd1-102">Enable single sign-on for Office Add-ins (preview)</span></span>

<span data-ttu-id="f2dd1-103">ユーザーは個人用の Microsoft アカウントまたは職場や学校の (Office 365) アカウントのいずれかを使用して、Office (オンライン、モバイル、デスクトップ プラットフォーム) にサインインします。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-103">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account.</span></span> <span data-ttu-id="f2dd1-104">これを利用して、シングル サインオン (SSO) を使用すれば、ユーザーに 2 度目のサインインを求めなくても、ユーザーにアドインの使用を承認できます。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-104">You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>

![アドインのサインイン プロセスを示す画像](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a><span data-ttu-id="f2dd1-106">プレビュー ステータス</span><span class="sxs-lookup"><span data-stu-id="f2dd1-106">Preview Status</span></span>

<span data-ttu-id="f2dd1-107">現時点で、シングル サインオン API をサポートするのはプレビューのみです。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-107">The Single Sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="f2dd1-108">開発者が試すことはできますが、実際に運用するアドインでは使用できません。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-108">It is available to developers for experimentation; but it should not be used in a production add-in.</span></span> <span data-ttu-id="f2dd1-109">さらに、SSO を使用するアドインは、[AppSource](https://appsource.microsoft.com) への掲載を認められません。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-109">In addition, add-ins that use SSO are not accepted in [AppSource](https://appsource.microsoft.com).</span></span>

<span data-ttu-id="f2dd1-110">一部の Office アプリケーションは、SSO プレビューをサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-110">Not all Office applications support the SSO preview.</span></span> <span data-ttu-id="f2dd1-111">Word、Excel、Outlook、PowerPoint では使用できます。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-111">It is available in Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="f2dd1-112">シングル サインオン API に対する現在のサポート状況の詳細については、「[IdentityAPI 要件セット](https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets?view=office-js)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-112">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets?view=office-js).</span></span>

### <a name="requirements-and-best-practices"></a><span data-ttu-id="f2dd1-113">要件とベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="f2dd1-113">Requirements and Best Practices</span></span>

<span data-ttu-id="f2dd1-114">SSO を使用するには、アドインの起動 HTML ページの `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` から、Office の JavaScript ライブラリのベータ版を読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-114">To use SSO, you must load the beta version of the Office JavaScript Library from `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` in the startup HTML page of the add-in.</span></span>

<span data-ttu-id="f2dd1-115">**Outlook** アドインを使用している場合は、Office 365 テナントの先進認証が有効になっていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-115">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="f2dd1-116">確認方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-116">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="f2dd1-117">SSO をアドインの唯一の認証メソッドして使用すべき*ではありません*。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-117">You should *not* rely on SSO as your add-in's only method of authentication.</span></span> <span data-ttu-id="f2dd1-118">特定のエラー状況において、アドインがフォールバックする代替の認証システムを実装する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-118">You should implement an alternate authentication system that your add-in can fall back to in certain error situations.</span></span> <span data-ttu-id="f2dd1-119">ユーザー テーブルと認証のシステムを使用することも、ソーシャル ログイン プロバイダのいずれかを活用することもできます。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-119">You can use a system of user tables and authentication, or you can leverage one of the social login providers.</span></span> <span data-ttu-id="f2dd1-120">Office アドインでこれを行う方法の詳細については、[「Office アドインで外部サービスを承認する」](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-120">For more information about how to do this with an Office add-in, see [Authorize external services in your Office Add-in](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins).</span></span> <span data-ttu-id="f2dd1-121"> \*Outlook* の場合、推奨されるフォールバック システムがあります。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-121">For *Outlook*, there is a recommended fall back system.</span></span> <span data-ttu-id="f2dd1-122">詳細については、「[シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-122">For more details, see [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

### <a name="how-sso-works-at-runtime"></a><span data-ttu-id="f2dd1-123">実行時に SSO が動作する仕組み</span><span class="sxs-lookup"><span data-stu-id="f2dd1-123">How it works at runtime</span></span>

<span data-ttu-id="f2dd1-124">次の図は、SSO の動作の仕組みを示しています。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-124">The following diagram shows how the SSO process works.</span></span>

![SSO プロセスを示す図](../images/sso-overview-diagram.png)

1. <span data-ttu-id="f2dd1-126">アドインでは、JavaScript は新しい Office.js API [getAccessTokenAsync](#sso-api-reference) を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-126">In the add-in, JavaScript calls a new Office.js API [](#sso-api-reference).</span></span> <span data-ttu-id="f2dd1-127">これにより、Office ホスト アプリケーションにアドインへのアクセス トークンを取得するように指示します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-127">This tells the Office host application to obtain an access token to the add-in.</span></span> <span data-ttu-id="f2dd1-128">[アクセス トークンの例](#example-access-token) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-128">See [Example access token](#example-access-token).</span></span>
2. <span data-ttu-id="f2dd1-129">ユーザーがサインインしていない場合、Office ホスト アプリケーションはユーザーにサインインを求めるポップアップ ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-129">If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="f2dd1-130">現在のユーザーが初めてアドインを使用する場合は、そのユーザーの同意を求めるダイアログを表示します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-130">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="f2dd1-131">Office ホスト アプリケーションは、Azure AD v2.0 エンドポイントに現在のユーザーの**アドイン トークン**を要求します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-131">The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="f2dd1-132">Azure AD は、Office ホスト アプリケーションにアドイン トークンを送信します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-132">Azure AD sends the add-in token to the Office host application.</span></span>
6. <span data-ttu-id="f2dd1-133">Office ホスト アプリケーションは、`getAccessTokenAsync` 呼び出しによって返される結果オブジェクトの一部として、アドインに**アドイン トークン**を送信します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-133">The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessTokenAsync` call.</span></span>
7. <span data-ttu-id="f2dd1-134">アドインの JavaScript は、トークンを解析し、必要な情報 (ユーザーの電子メールアドレスなど) を抽出できます。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-134">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span> 
8. <span data-ttu-id="f2dd1-135">アプションで、アドインはサーバー側に HTTP 要求を送信して、ユーザーに関する詳細なデータ (ユーザーの嗜好など) を得ることができます。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-135">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences.</span></span> <span data-ttu-id="f2dd1-136">もしくは、アクセス トークン自体をサーバー側に送信して、解析と検証をすることもできます。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-136">Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span> 

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="f2dd1-137">SSO アドインの開発</span><span class="sxs-lookup"><span data-stu-id="f2dd1-137">Develop an SSO add-in</span></span>

<span data-ttu-id="f2dd1-138">このセクションでは、SSO を使用する Office アドインの作成に関連するタスクについて説明します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-138">This section describes the tasks involved in creating an Office Add-in that uses SSO.</span></span> <span data-ttu-id="f2dd1-139">ここでは、言語とフレームワークに依存しない方法でこれらのタスクを説明します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-139">These tasks are described here in a language- and framework-agnostic way.</span></span> <span data-ttu-id="f2dd1-140">詳細なウォークスルーの例については、次を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-140">For examples of detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="f2dd1-141">シングル サインオンを使用する Node.js Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="f2dd1-141">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="f2dd1-142">シングル サインオンを使用する ASP.NET Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="f2dd1-142">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a><span data-ttu-id="f2dd1-143">サービス アプリケーションを作成する</span><span class="sxs-lookup"><span data-stu-id="f2dd1-143">Create the service application</span></span>

<span data-ttu-id="f2dd1-144">Azure v2.0 エンドポイントの登録ポータルでアドインを登録します。https://apps.dev.microsoft.com。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-144">Register the add-in at the registration portal for the Azure v2.0 endpoint: https://apps.dev.microsoft.com. This is a 5–10 minute process that includes the following tasks:</span></span> <span data-ttu-id="f2dd1-145">これは、次のタスクを含む 5 〜 10 分程度の時間がかかるプロセスです。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-145">This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="f2dd1-146">アドインのクライアント ID とシークレットを取得します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-146">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="f2dd1-147">アドインが AAD v に必要とするアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-147">Specify the permissions that your add-in needs to Microsoft Graph.</span></span> <span data-ttu-id="f2dd1-148">(オプションで Microsoft Graph)。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-148">2.0 endpoint (and optionally to Microsoft Graph).</span></span> <span data-ttu-id="f2dd1-149">[プロファイル] のアクセス許可は、常に必要です。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-149">The "profile" permission is always needed.</span></span>
* <span data-ttu-id="f2dd1-150">Office ホスト アプリケーション信頼をアドインに付与します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-150">Grant the Office host application trust to the add-in.</span></span>
* <span data-ttu-id="f2dd1-151">既定のアクセス許可 *access_as_user* を使用して、Office ホスト アプリケーションのアドインへのアクセスを事前認証します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-151">Preauthorize the Office host application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="f2dd1-152">このプロセスの詳細については、「 [Azure AD v2.0 のエンドポイントで SSO を使用する Office アドインを登録する](register-sso-add-in-aad-v2.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-152">For more information, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="f2dd1-153">アドインを構成する</span><span class="sxs-lookup"><span data-stu-id="f2dd1-153">Configure the add-in</span></span>

<span data-ttu-id="f2dd1-154">アドイン マニフェストに新しいマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-154">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="f2dd1-155">**WebApplicationInfo** - 次の要素の親。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-155">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="f2dd1-156">**Id** - アドインのクライアント ID。これは、アドイン登録の一環として取得するアプリケーション ID です。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-156">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in.</span></span> <span data-ttu-id="f2dd1-157">「[Azure AD v2.0 のエンドポイントで SSO を使用する Office アドインを登録する](register-sso-add-in-aad-v2.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-157">Details are at: [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="f2dd1-158">**Resource** - アドインの URL。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-158">**Resource** - The URL of the add-in.</span></span>
* <span data-ttu-id="f2dd1-159">**Scopes** - 1 つ以上の **Scope** 要素の親。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-159">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="f2dd1-160">**Scope** - アドインが AAD に必要とするアクセス許可を指定します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-160">**Scope** - Specifies a permission that the add-in needs to Microsoft Graph.</span></span> <span data-ttu-id="f2dd1-161">`profile` アクセス許可は常に必要とされます。アドインに Microsoft Graph へのアクセス権がないとき、唯一必要な権限はこの許可である場合があります。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-161">The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph.</span></span> <span data-ttu-id="f2dd1-162">このような場合、必要な Microsoft Graph へのアクセス許可を得るには **Scope** 要素も必要です。たとえば、 `User.Read`、 `Mail.Read` などです。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-162">If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`.</span></span> <span data-ttu-id="f2dd1-163">Microsoft Graph にアクセスするためにコードで使用するライブラリには、アクセス許可がさらに必要な場合があります。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-163">Libraries that you use in your code to access Microsoft Graph may need additional permissions.</span></span> <span data-ttu-id="f2dd1-164">たとえば、Microsoft 認証ライブラリ (MSAL) for .NET は、 `offline_access` アクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-164">For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission.</span></span> <span data-ttu-id="f2dd1-165">詳細については、「 [Office アドインから Microsoft Graph を承認する](authorize-to-microsoft-graph.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-165">For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="f2dd1-p113">Outlook 以外の Office ホストでは、`<VersionOverrides ... xsi:type="VersionOverridesV1_0">` セクションの末尾にマークアップを追加します。Outlook では、`<VersionOverrides ... xsi:type="VersionOverridesV1_1">` セクションの末尾にマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-p113">For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="f2dd1-168">マークアップの例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-168">The following is an example of the markup:</span></span>

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

### <a name="add-client-side-code"></a><span data-ttu-id="f2dd1-169">クライアント側のコードを追加する</span><span class="sxs-lookup"><span data-stu-id="f2dd1-169">Add client-side code</span></span>

<span data-ttu-id="f2dd1-170">アドインに JavaScript を追加します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-170">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="f2dd1-171">[GetAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-171">Call [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span></span>

* <span data-ttu-id="f2dd1-172">アクセストークンを解析するか、アドインのサーバー側コードに渡します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-172">Parse the access token or pass it to the add-in’s server-side code.</span></span> 

<span data-ttu-id="f2dd1-173">ここでは、`getAccessTokenAsync` の呼び出しの簡単な例を示します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-173">Here's a simple example of a call to `getAccessTokenAsync`.</span></span> 

> [!NOTE]
> <span data-ttu-id="f2dd1-174">この例では、1 種類のエラーのみを明示的に処理しています。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-174">This example handles only one kind of error explicitly.</span></span> <span data-ttu-id="f2dd1-175">より詳細なエラー処理の例については、「[Office-Add-in-ASPNET-SSO の Home.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js)」および「[Office-Add-in-NodeJS-SSO の program.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-175">For examples of more elaborate error handling, see [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) and [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span></span> <span data-ttu-id="f2dd1-176">また、「[シングル サインオン (SSO) のエラー メッセージのトラブルシューティング](troubleshoot-sso-in-office-add-ins.md)」 もご覧ください。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-176">Troubleshoot error messages for single sign-on (SSO)</span></span>
 

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

<span data-ttu-id="f2dd1-177">ここでは、アドイン トークンをサーバー側に渡す簡単な例を示します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-177">Here's a simple example of passing the add-in token to the server-side.</span></span> <span data-ttu-id="f2dd1-178">サーバー側に要求を送り返すときは、トークンが `Authorization` ヘッダーとして含まれます。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-178">The token is included as an `Authorization` header when sending a request back to the server-side.</span></span> <span data-ttu-id="f2dd1-179">この例では、JSON データを送信すると想定しています。そのため、`POST` メソッドを使用しますが、サーバーへの書き込みを行っていない場合、アクセス トークンの送信には `GET` で十分です。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-179">This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

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

#### <a name="when-to-call-the-method"></a><span data-ttu-id="f2dd1-180">メソッドを呼び出すタイミング</span><span class="sxs-lookup"><span data-stu-id="f2dd1-180">When to call the method</span></span>

<span data-ttu-id="f2dd1-181">Office にログインしているユーザーがなく、アドインを使用できない場合には、*アドインを起動するときに* `getAccessTokenAsync` を呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-181">If your add-in cannot be used when a no user is logged into Office and Office does not have an access token to your add-in, then you should call `getAccessTokenAsync` *when the add-in launches*.</span></span>

<span data-ttu-id="f2dd1-182">アドインにログイン ユーザーを必要としない機能がある場合、*ユーザーがログイン ユーザーを必要とするアクションを実行するときに* `getAccessTokenAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-182">If the add-in has some functionality that doesn't require access to Microsoft Graph or even a logged in user, then you call `getAccessTokenAsync` *when the user takes an action that requires access to Microsoft Graph or, at least, a logged in user*.</span></span> <span data-ttu-id="f2dd1-183">`getAccessTokenAsync`  の重複呼び出しによってパフォーマンスが大幅に低下することはありません。これは、Office ではアクセス トークンがキャッシュされ、それが期限切れになるまで、AAD V が再度呼び出されることなく、再利用されるためです。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-183">There is no significant performance degradation with redundant calls of `getAccessTokenAsync` because Office caches the access token and will reuse it, until it expires, without making another call to the AAD V. 2.0 endpoint whenever  is called.</span></span> <span data-ttu-id="f2dd1-184">`getAccessTokenAsync` が呼び出されたときに毎回 AAD v. 2.0 エンドポイントを呼び出すことなく再利用するためです。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-184">2.0 endpoint whenever `getAccessTokenAsync` is called.</span></span> <span data-ttu-id="f2dd1-185">このため、`getAccessTokenAsync` の呼び出しを、このトークンが必要とされる場所でアクションを開始するすべての関数とハンドラーに追加します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-185">So you can add calls of `getAccessTokenAsync` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="f2dd1-186">サーバー側のコードを追加する</span><span class="sxs-lookup"><span data-stu-id="f2dd1-186">Add server-side code</span></span>

<span data-ttu-id="f2dd1-187">ほとんどの場合において、アドインが取得しサーバーに渡したアクセス トークンをサーバー側で使用しない場合は、アクセス トークンを取得することに意味はありません。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-187">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="f2dd1-188">アドインで実行できるサーバー側タスクの一部を次に示します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-188">Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="f2dd1-189">トークンから抽出されたユーザーに関する情報を使用する 1 つ以上の Web API メソッドを作成します。たとえば、ホスト型データベース内のユーザーの嗜好を参照するメソッドです。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-189">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base.</span></span> <span data-ttu-id="f2dd1-190">(以下の「 **ID として SSO トークンを使用する** 」を参照してください)。使用している言語とフレームワークによっては、作成する必要があるコードを簡略化できるライブラリを利用できる場合があります。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-190">(See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="f2dd1-191">Microsoft Graph データを取得します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-191">Get Microsoft Graph data.</span></span> <span data-ttu-id="f2dd1-192">サーバー側のコードでは、次に示す操作を実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-192">Your server-side code should do the following:</span></span>

    * <span data-ttu-id="f2dd1-193">アクセストークンを検証する (以下の「**アクセストークンを検証する**」を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-193">Validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="f2dd1-194">アクセス トークン、ユーザーに関するメタデータ、アドインの認証情報 (ID とシークレット) を含む Azure AD v2.0 エンドポイントへの呼び出しを使用して "On-Behalf-Of" フローを開始します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-194">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the add-in access token, some metadata about the user, and the credentials of the add-in (its ID and secret).</span></span> <span data-ttu-id="f2dd1-195">このコンテキストでは、アクセストークンはブートストラップトークンと呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-195">In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="f2dd1-196">"On-Behalf-Of" フローから返された新しいアクセス トークンをキャッシュします。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-196">Cache the new access token that is returned from the on-behalf-of flow.</span></span>
    * <span data-ttu-id="f2dd1-197">新しいトークンを使用して Microsoft Graph からデータを取得します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-197">Get data from Microsoft Graph by using the MSG token.</span></span>

 <span data-ttu-id="f2dd1-198">ユーザーの Microsoft Graph データへの承認済みアクセスを取得する方法の詳細については、「[Office アドインで Microsoft Graph を承認する](authorize-to-microsoft-graph.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-198">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="f2dd1-199">アクセス トークンを検証する</span><span class="sxs-lookup"><span data-stu-id="f2dd1-199">For more information, see Validate the access token.</span></span>

<span data-ttu-id="f2dd1-200">Web API でアクセス トークンを受信したら、そのアクセス トークンを使用する前に検証する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-200">Once the Web API receives the access token, it must validate it before using it.</span></span> <span data-ttu-id="f2dd1-201">このトークンは、JSON Web トークン (JWT) であるため、その検証はほとんどの標準 OAuth のトークン検証と同様に動作します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-201">The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows.</span></span> <span data-ttu-id="f2dd1-202">JWT の検証を処理できるライブラリが多数ありますが、その基本は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-202">There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="f2dd1-203">トークンが整形式であることを確認する</span><span class="sxs-lookup"><span data-stu-id="f2dd1-203">Checking that the token is well-formed</span></span>
- <span data-ttu-id="f2dd1-204">トークンが意図した証明機関から発行されたことを確認する</span><span class="sxs-lookup"><span data-stu-id="f2dd1-204">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="f2dd1-205">トークンが Web API を対象にしていることを確認する</span><span class="sxs-lookup"><span data-stu-id="f2dd1-205">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="f2dd1-206">トークンの検証時には、次のガイドラインに注意してください。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-206">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="f2dd1-207">有効な SSO トークンは Azure 証明機関 `https://login.microsoftonline.com` から発行されます。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-207">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`.</span></span> <span data-ttu-id="f2dd1-208">トークン内の `iss` クレームは、この値で始まっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-208">The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="f2dd1-209">トークンの `aud` パラメータは、アドインの登録のアプリケーション ID に設定します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-209">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="f2dd1-210">トークンの `scp` パラメータは `access_as_user` に設定します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-210">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="f2dd1-211">ID として SSO トークンを使用する</span><span class="sxs-lookup"><span data-stu-id="f2dd1-211">Using the SSO token as an identity</span></span>

<span data-ttu-id="f2dd1-212">アドインでユーザーの ID を検証する必要がある場合、SSO トークンには ID を確定するために使用できる情報が含まれています。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-212">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity.</span></span> <span data-ttu-id="f2dd1-213">ID に関連するトークン内のクレームは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-213">The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="f2dd1-214">`name` - ユーザーの表示名。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-214">`name` - The user's display name.</span></span>
- <span data-ttu-id="f2dd1-215">`preferred_username` - ユーザーの電子メール アドレス。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-215">`preferred_username`The user's email address.</span></span>
- <span data-ttu-id="f2dd1-216">`oid` - Azure Active Directory でユーザーの ID を表す GUID。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-216">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="f2dd1-217">`tid` - Azure Active Directory でユーザーの組織の ID を表す GUID。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-217">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="f2dd1-218">`name` と `preferred_username` の値は変化することがあるため、`oid` と `tid` の値を ID とバックエンドの承認サービスを関連付けるために使用するようにしてください。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-218">Since the `name` and `preferred_username` values could change, it's recommended that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="f2dd1-219">たとえば、サービスでは、これらの値を `{oid-value}@{tid-value}` のような形式にまとめて、内部ユーザー データベースのユーザーのレコードに値として保存できます。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-219">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database.</span></span> <span data-ttu-id="f2dd1-220">その後の要求では、同じ値を使用してユーザーを取得できるようになり、特定のリソースへのアクセスについては既存のアクセス制御メカニズムに基づいて決定できます。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-220">Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="f2dd1-221">アクセス トークンの例</span><span class="sxs-lookup"><span data-stu-id="f2dd1-221">Example access token</span></span>

<span data-ttu-id="f2dd1-222">アクセス トークンの標準的なデコードされたペイロードを次に示します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-222">The following is a typical decoded payload of an access token.</span></span> <span data-ttu-id="f2dd1-223">プロパティの詳細については、「[Azure Active Directory v2.0 トークンリファレンス](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-223">For information about the properties, see [Azure Active Directory v2.0 tokens reference](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).</span></span>


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

## <a name="using-sso-with-an-outlook-add-in"></a><span data-ttu-id="f2dd1-224">Outlook のアドインでの SSO の使用</span><span class="sxs-lookup"><span data-stu-id="f2dd1-224">Using SSO with and Outlook add-in</span></span>

<span data-ttu-id="f2dd1-225">SSO を Outlook アドインとして使用する場合と、Excel、PowerPoint、Word アドインとして使用する場合には、ささやかですが重要な違いがあります。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-225">There are some small, but important differences in using SSO in and Outlook add-in from using it in as Excel, PowerPoint, or Word add-in.</span></span> <span data-ttu-id="f2dd1-226">「[Outlook アドインでシングルサインオン トークンを使用してユーザーを認証する](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)」と「[シナリオ - Outlook アドインでサービスにシングルサインオンを実装する](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)」を必ず読んでください。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-226">Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) and [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

## <a name="sso-api-reference"></a><span data-ttu-id="f2dd1-227">SSO API 参照</span><span class="sxs-lookup"><span data-stu-id="f2dd1-227">SSO API reference</span></span>

### <a name="getaccesstokenasync"></a><span data-ttu-id="f2dd1-228">getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="f2dd1-228">getAccessTokenAsync</span></span>

<span data-ttu-id="f2dd1-229">Office の認証の名前空間 `Office.context.auth` は、Office ホストがアドインの Web アプリケーションへのアクセス トークンを取得できるようにするメソッド `getAccessTokenAsync` を提供します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-229">The Office Auth namespace, `Office.context.auth`, provides a method, `getAccessTokenAsync` that enables the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="f2dd1-230">これはまた間接的に、ユーザーがもう一度サインインする必要なしに、アドインがサインインしたユーザーの Microsoft Graph データにアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-230">Indirectly, enable the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

<span data-ttu-id="f2dd1-231">このメソッドは、Azure Active Directory V 2.0 のエンドポイントを呼び出して、アドインの Web アプリケーションへのアクセス トークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-231">Calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application.</span></span> <span data-ttu-id="f2dd1-232">これにより、アドインを使用してユーザーを識別できます。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-232">This enables add-ins to identify users.</span></span> <span data-ttu-id="f2dd1-233">["on behalf of" OAuth フロー](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)を使用することにより、サーバー側のコードはこのトークンを使用してアドインの Web アプリケーションの Microsoft Graph にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-233">Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span>

> [!NOTE]
> [!Outlook メモによれば、アドインが Outlook.com または Gmail のメールボックスに読み込まれている場合は、この API はサポートされていません。]

<table><tr><td><span data-ttu-id="f2dd1-235">Hosts</span><span class="sxs-lookup"><span data-stu-id="f2dd1-235">Hosts</span></span></td><td><span data-ttu-id="f2dd1-236">Excel、OneNote、Outlook、PowerPoint、Word</span><span class="sxs-lookup"><span data-stu-id="f2dd1-236">Excel, Outlook, PowerPoint, Word</span></span></td></tr>

 <tr><td><span data-ttu-id="f2dd1-237">要件セット</span><span class="sxs-lookup"><span data-stu-id="f2dd1-237">Requirement sets</span></span></td><td>[<span data-ttu-id="f2dd1-238">IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="f2dd1-238">IdentityAPI</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td></tr></table>

#### <a name="parameters"></a><span data-ttu-id="f2dd1-239">パラメータ</span><span class="sxs-lookup"><span data-stu-id="f2dd1-239">Parameters</span></span>

<span data-ttu-id="f2dd1-240">`options` - 省略可能。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-240">`options` - Optional.</span></span> <span data-ttu-id="f2dd1-241">サインオン時の動作を定義するために、`AuthOptions` オブジェクト (下記参照) を受け入れます。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-241">Accepts an `AuthOptions` object (see below) to define sign-on behaviors.</span></span>

<span data-ttu-id="f2dd1-242">`callback` - 省略可能。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-242">`callback` - Optional.</span></span> <span data-ttu-id="f2dd1-243">ユーザーの ID のトークンを解析したり、"On-Behalf-Of" フローのトークンを使用して Microsoft Graph へのアクセスを取得することができるコールバック メソッドを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-243">Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph.</span></span> <span data-ttu-id="f2dd1-244">[AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` が「成功した」場合は、 `AsyncResult.value` は生の AAD v です。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-244">If [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v.</span></span> <span data-ttu-id="f2dd1-245">2.0 形式のアクセス トークンです。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-245">2.0-formatted access token.</span></span>

<span data-ttu-id="f2dd1-246">`AuthOptions` インターフェイスは、Office が AAD v からアドインへのアクセス トークンを取得する場合に、ユーザー エクスペリエンスのオプションを提供します。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-246">The `AuthOptions` interface provides options for the user experience when Office obtains an access token to the add-in from AAD v.</span></span> <span data-ttu-id="f2dd1-247">`getAccessTokenAsync` メソッドを持つ 2.0 です。</span><span class="sxs-lookup"><span data-stu-id="f2dd1-247">2.0 with the `getAccessTokenAsync` method.</span></span>

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



