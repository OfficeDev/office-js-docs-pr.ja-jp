---
title: Office アドインのシングル サインオンを有効化する
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: dc9050d574e0a5e74ae8cae2c63817aa4f952eb9
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/20/2019
ms.locfileid: "30691196"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a><span data-ttu-id="68843-102">Office アドインのシングル サインオンを有効化する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="68843-102">Enable single sign-on for Office Add-ins (preview)</span></span>

<span data-ttu-id="68843-103">ユーザーは個人用の Microsoft アカウントまたは職場や学校の (Office 365) アカウントのいずれかを使用して、Office (オンライン、モバイル、およびデスクトップ プラットフォーム) にサインインします。</span><span class="sxs-lookup"><span data-stu-id="68843-103">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account.</span></span> <span data-ttu-id="68843-104">これとシングル サインオン (SSO) を使用すれば、ユーザーに 2 度目のサインインを求めずに、ご自分のアドインをユーザーに許可できます。</span><span class="sxs-lookup"><span data-stu-id="68843-104">You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>

![アドインのサインイン プロセスを示す画像](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a><span data-ttu-id="68843-106">プレビューの状態</span><span class="sxs-lookup"><span data-stu-id="68843-106">Preview Status</span></span>

<span data-ttu-id="68843-107">現在、シングル サインオン API はプレビューのみでサポートされています。</span><span class="sxs-lookup"><span data-stu-id="68843-107">The Single Sign-on API is currently supported in preview only.</span></span> <span data-ttu-id="68843-108">これは、試験目的のみで開発者に提供されており、運用環境のアドインには使用してはいけません。</span><span class="sxs-lookup"><span data-stu-id="68843-108">It is available to developers for experimentation; but it should not be used in a production add-in.</span></span> <span data-ttu-id="68843-109">また、SSO を使用するアドインは [AppSource](https://appsource.microsoft.com) では許可されていません。</span><span class="sxs-lookup"><span data-stu-id="68843-109">In addition, add-ins that use SSO are not accepted in [AppSource](https://appsource.microsoft.com).</span></span>

<span data-ttu-id="68843-110">SSO には Office 365 (Office のサブスクリプション バージョン) が必要です。</span><span class="sxs-lookup"><span data-stu-id="68843-110">SSO requires Office 365 (the subscription version of Office, also called “Click to Run”).</span></span> <span data-ttu-id="68843-111">Insider チャネルからの最新の月次バージョンとビルドを使ってください。</span><span class="sxs-lookup"><span data-stu-id="68843-111">You should use the latest monthly version and build from the Insiders channel.</span></span> <span data-ttu-id="68843-112">このバージョンを入手するには、Office Insider への参加が必要です。</span><span class="sxs-lookup"><span data-stu-id="68843-112">You need to be an Office Insider to get this version.</span></span> <span data-ttu-id="68843-113">詳細については、「[Office Insider になる](https://products.office.com/office-insider?tab=tab-1)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="68843-113">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span> <span data-ttu-id="68843-114">ビルドが半期チャネルの運用に移行すると、そのビルドで SSO を含むプレビュー機能のサポートはオフになりますので、ご注意ください。</span><span class="sxs-lookup"><span data-stu-id="68843-114">Please note that when a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.</span></span>

<span data-ttu-id="68843-115">SSO のプレビューは、すべての Office アプリケーションではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="68843-115">Not all Office applications support the SSO preview.</span></span> <span data-ttu-id="68843-116">これは、Word、Excel、Outlook、および PowerPoint で利用できます。</span><span class="sxs-lookup"><span data-stu-id="68843-116">It is available in Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="68843-117">シングル サインオン API の現在のサポート状態に関する詳細は、「[IdentityAPI の要件セット](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="68843-117">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets).</span></span>

### <a name="requirements-and-best-practices"></a><span data-ttu-id="68843-118">要件とベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="68843-118">Requirements and Best Practices</span></span>

<span data-ttu-id="68843-119">SSO を使用するには、アドインのスタートアップ HTML ページの `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` から Office JavaScript ライブラリのベータ版を読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="68843-119">To use SSO, you must load the beta version of the Office JavaScript Library from `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` in the startup HTML page of the add-in.</span></span>

<span data-ttu-id="68843-120">**Outlook** アドインで作業している場合は、Office 365 テナントの先進認証が有効になっていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="68843-120">If you are working with an **Outlook** add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="68843-121">この方法の詳細については、「[Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」 (Exchange Online: テナントの先進認証を有効にする方法) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="68843-121">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="68843-122">SSO をアドインの唯一の認証方法と*しない*ようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="68843-122">You should *not* rely on SSO as your add-in's only method of authentication.</span></span> <span data-ttu-id="68843-123">特定のエラー状況でアドインが切り替えることができる、別の認証システムを実装する必要があります。</span><span class="sxs-lookup"><span data-stu-id="68843-123">You should implement an alternate authentication system that your add-in can fall back to in certain error situations.</span></span> <span data-ttu-id="68843-124">ユーザー テーブルと認証のシステムを使用するか、ソーシャル ログイン プロバイダーの 1 つを活用できます。</span><span class="sxs-lookup"><span data-stu-id="68843-124">You can use a system of user tables and authentication, or you can leverage one of the social login providers.</span></span> <span data-ttu-id="68843-125">Office アドインでこれを実行する方法の詳細については、「[Office アドインで外部サービスを承認する](/office/dev/add-ins/develop/auth-external-add-ins)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="68843-125">For more information about how to do this with an Office add-in, see [Authorize external services in your Office Add-in](/office/dev/add-ins/develop/auth-external-add-ins).</span></span> <span data-ttu-id="68843-126">*Outlook* には切り替えることが可能な推奨システムがあります。</span><span class="sxs-lookup"><span data-stu-id="68843-126">For *Outlook*, there is a recommended fall back system.</span></span> <span data-ttu-id="68843-127">詳細については、「[シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](/outlook/add-ins/implement-sso-in-outlook-add-in)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="68843-127">For more information, see [Scenario: Implement single sign-on to your service in an Outlook add-in](/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

### <a name="how-sso-works-at-runtime"></a><span data-ttu-id="68843-128">実行時の SSO の動作のしくみ</span><span class="sxs-lookup"><span data-stu-id="68843-128">How SSO works at runtime</span></span>

<span data-ttu-id="68843-129">次の図は、SSO の動作のしくみを示しています。</span><span class="sxs-lookup"><span data-stu-id="68843-129">The following diagram shows how the SSO process works.</span></span>

![SSO プロセスを示す図](../images/sso-overview-diagram.png)

1. <span data-ttu-id="68843-131">アドインの JavaScript により、新しい Office.js API [getAccessTokenAsync](#sso-api-reference) が呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="68843-131">In the add-in, JavaScript calls a new Office.js API [getAccessTokenAsync](#sso-api-reference).</span></span> <span data-ttu-id="68843-132">これにより、Office ホスト アプリケーションはアドインへのアクセス トークンを取得するよう指示されます。</span><span class="sxs-lookup"><span data-stu-id="68843-132">This tells the Office host application to obtain an access token to the add-in.</span></span> <span data-ttu-id="68843-133">「[アクセス トークンの例](#example-access-token)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="68843-133">See [Example access token](#example-access-token).</span></span>
2. <span data-ttu-id="68843-134">ユーザーがサインインしていない場合、Office ホスト アプリケーションはユーザーにサインインを求めるポップアップ ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="68843-134">If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="68843-135">現在のユーザーが初めてアドインを使用する場合は、そのユーザーに同意を求めるダイアログを表示します。</span><span class="sxs-lookup"><span data-stu-id="68843-135">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="68843-136">Office ホスト アプリケーションは、Azure AD v2.0 エンドポイントから現在のユーザーの**アドイン トークン**を要求します。</span><span class="sxs-lookup"><span data-stu-id="68843-136">The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="68843-137">Azure AD は、Office ホスト アプリケーションにアドイン トークンを送信します。</span><span class="sxs-lookup"><span data-stu-id="68843-137">Azure AD sends the add-in token to the Office host application.</span></span>
6. <span data-ttu-id="68843-138">Office ホスト アプリケーションが、`getAccessTokenAsync` 呼び出しによって返される結果オブジェクトの一部として、アドインに**アドイン トークン**を送信します。</span><span class="sxs-lookup"><span data-stu-id="68843-138">The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessTokenAsync` call.</span></span>
7. <span data-ttu-id="68843-139">アドイン内の JavaScript が、トークンを解析し、ユーザーのメール アドレスなど必要な情報を抽出します。</span><span class="sxs-lookup"><span data-stu-id="68843-139">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span> 
8. <span data-ttu-id="68843-140">オプションで、アドインで HTTP 要求を送信して、ユーザー設定などユーザーに関する情報をさらにサーバー側から求めることができます。</span><span class="sxs-lookup"><span data-stu-id="68843-140">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences.</span></span> <span data-ttu-id="68843-141">または、アクセス トークン自体が解析および検証されるようにサーバー側に送信することができます。</span><span class="sxs-lookup"><span data-stu-id="68843-141">Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span>

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="68843-142">SSO アドインの開発</span><span class="sxs-lookup"><span data-stu-id="68843-142">Develop an SSO add-in</span></span>

<span data-ttu-id="68843-p109">このセクションでは、SSO を使用する Office アドインの作成に関連するタスクについて説明します。 ここでは、これらのタスクについて、言語とフレームワークに依存しない方法で説明しています。 詳細なチュートリアルの例については、次を参照してください。</span><span class="sxs-lookup"><span data-stu-id="68843-p109">This section describes the tasks involved in creating an Office Add-in that uses SSO. These tasks are described here in a language- and framework-agnostic way. For examples of detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="68843-146">シングル サインオンを使用する Node.js Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="68843-146">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="68843-147">シングル サインオンを使用する ASP.NET Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="68843-147">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a><span data-ttu-id="68843-148">サービス アプリケーションを作成する</span><span class="sxs-lookup"><span data-stu-id="68843-148">Create the service application</span></span>

<span data-ttu-id="68843-p110">Azure v2.0 エンドポイントの登録ポータル (https://apps.dev.microsoft.com でアドインを登録します。このプロセスには、次に示すタスクを含めて 5 分から 10 分の時間がかかります。</span><span class="sxs-lookup"><span data-stu-id="68843-p110">Register the add-in at the registration portal for the Azure v2.0 endpoint: https://apps.dev.microsoft.com. This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="68843-151">アドインのクライアント ID とシークレットを取得します。</span><span class="sxs-lookup"><span data-stu-id="68843-151">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="68843-152">アドインが必要とする AAD v.2.0 エンドポイントへのアクセス許可を指定します </span><span class="sxs-lookup"><span data-stu-id="68843-152">Specify the permissions that your add-in needs to AAD v.</span></span> <span data-ttu-id="68843-153">(必要に応じて Microsoft Graph へも指定します)。</span><span class="sxs-lookup"><span data-stu-id="68843-153">2.0 endpoint (and optionally to Microsoft Graph).</span></span> <span data-ttu-id="68843-154">"profile" のアクセス許可は常に必要です。</span><span class="sxs-lookup"><span data-stu-id="68843-154">The "profile" permission is always needed.</span></span>
* <span data-ttu-id="68843-155">Office ホスト アプリケーションにアドインへの信頼を付与します。</span><span class="sxs-lookup"><span data-stu-id="68843-155">Grant the Office host application trust to the add-in.</span></span>
* <span data-ttu-id="68843-156">既定のアクセス許可 *access_as_user* を使用して、Office ホスト アプリケーションのアドインへのアクセスを事前認証します。</span><span class="sxs-lookup"><span data-stu-id="68843-156">Preauthorize the Office host application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="68843-157">この手順の詳細については、「[Azure AD v2.0 のエンドポイントに SSO を使用する Office アドインを登録する](register-sso-add-in-aad-v2.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="68843-157">For more details about this process, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="68843-158">アドインを構成する</span><span class="sxs-lookup"><span data-stu-id="68843-158">Configure the add-in</span></span>

<span data-ttu-id="68843-159">新しいマークアップをアドイン マニフェストに追加します。</span><span class="sxs-lookup"><span data-stu-id="68843-159">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="68843-160">**WebApplicationInfo** - 次の要素の親。</span><span class="sxs-lookup"><span data-stu-id="68843-160">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="68843-161">**Id** -このアドインのクライアント ID。これはアドインを登録する一貫として取得するアプリケーション ID です。</span><span class="sxs-lookup"><span data-stu-id="68843-161">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in.</span></span> <span data-ttu-id="68843-162">詳細については、「[Azure AD v2.0 のエンドポイントに SSO を使用する Office アドインを登録する](register-sso-add-in-aad-v2.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="68843-162">See [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="68843-163">**Resource** - アドインの URL。</span><span class="sxs-lookup"><span data-stu-id="68843-163">**Resource** - The URL of the add-in.</span></span> <span data-ttu-id="68843-164">これは、AAD にアドインを登録したときに使用したのと同じ URI (`api:` プロトコルを含む) です。</span><span class="sxs-lookup"><span data-stu-id="68843-164">This is the same URI (including the `api:` protocol) that you used when registering the add-in in AAD.</span></span> <span data-ttu-id="68843-165">この URI のドメイン部分は、アドインのマニフェストの `<Resources>` のセクションの URL で使用されている、すべてのサブドメインを含むドメインと一致している必要があります。</span><span class="sxs-lookup"><span data-stu-id="68843-165">The domain part of this URI should match the domain, including any subdomains, used in the URLs in the `<Resources>` section of the add-in's manifest.</span></span>
* <span data-ttu-id="68843-166">**Scopes** - 1 つ以上の **Scope** 要素の親。</span><span class="sxs-lookup"><span data-stu-id="68843-166">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="68843-167">**Scope** - アドインが AAD に対して必要なアクセス許可を指定する。</span><span class="sxs-lookup"><span data-stu-id="68843-167">**Scope** - Specifies a permission that the add-in needs to AAD.</span></span> <span data-ttu-id="68843-168">`profile` のアクセス許可は常に必要です。ご使用のアドインが Microsoft Graph にアクセスしない場合、これは必要な唯一のアクセス許可になる場合があります。</span><span class="sxs-lookup"><span data-stu-id="68843-168">The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph.</span></span> <span data-ttu-id="68843-169">アクセスする場合、Microsoft Graph へのアクセスに必要な許可として、`User.Read`、`Mail.Read` など **Scope** 要素も必要になります。</span><span class="sxs-lookup"><span data-stu-id="68843-169">If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`.</span></span> <span data-ttu-id="68843-170">コードで使用している、Microsoft Graph にアクセスするためのライブラリでは、他にもアクセス許可が必要な場合があります。</span><span class="sxs-lookup"><span data-stu-id="68843-170">Libraries that you use in your code to access Microsoft Graph may need additional permissions.</span></span> <span data-ttu-id="68843-171">たとえば、.NET 用の Microsoft 認証ライブラリ (MSAL) では、`offline_access` のアクセス許可が必要です。</span><span class="sxs-lookup"><span data-stu-id="68843-171">For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission.</span></span> <span data-ttu-id="68843-172">詳細については、「[Office アドインで Microsoft Graph へ承認](authorize-to-microsoft-graph.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="68843-172">For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="68843-p115">Outlook 以外の Office ホストでは、`<VersionOverrides ... xsi:type="VersionOverridesV1_0">` セクションの末尾にマークアップを追加します。Outlook では、`<VersionOverrides ... xsi:type="VersionOverridesV1_1">` セクションの末尾にマークアップを追加します。</span><span class="sxs-lookup"><span data-stu-id="68843-p115">For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="68843-175">マークアップの例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="68843-175">The following is an example of the markup:</span></span>

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

### <a name="add-client-side-code"></a><span data-ttu-id="68843-176">クライアント側のコードを追加する</span><span class="sxs-lookup"><span data-stu-id="68843-176">Add client-side code</span></span>

<span data-ttu-id="68843-177">アドインに次のために JavaScript を追加します。</span><span class="sxs-lookup"><span data-stu-id="68843-177">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="68843-178">[getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) の呼び出し。</span><span class="sxs-lookup"><span data-stu-id="68843-178">Call [getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span></span>

* <span data-ttu-id="68843-179">アクセス トークンを解析するか、それをアドインのサーバー側コードに渡す。</span><span class="sxs-lookup"><span data-stu-id="68843-179">Parse the access token or pass it to the add-in’s server-side code.</span></span>

<span data-ttu-id="68843-180">`getAccessTokenAsync` への呼び出しの単純な例を示します。</span><span class="sxs-lookup"><span data-stu-id="68843-180">Here's a simple example of a call to `getAccessTokenAsync`.</span></span>

> [!NOTE]
> <span data-ttu-id="68843-181">この例では、1 種類のエラーのみを明示的に処理します。</span><span class="sxs-lookup"><span data-stu-id="68843-181">This example handles only one kind of error explicitly.</span></span> <span data-ttu-id="68843-182">より複雑なエラー処理の例については、「[Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js)」 (Office-Add-in-ASPNET-SSO の Home.js) と 「[program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)」 (Office-Add-in-NodeJS-SSO の program.js) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="68843-182">For examples of more elaborate error handling, see [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) and [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span></span> <span data-ttu-id="68843-183">また、「[シングル サインオン (SSO) のエラー メッセージのトラブルシューティング](troubleshoot-sso-in-office-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="68843-183">And see [Troubleshoot error messages for single sign-on (SSO)](troubleshoot-sso-in-office-add-ins.md).</span></span>
 

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

<span data-ttu-id="68843-184">サーバー側にアドイン トークンを渡す単純な例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="68843-184">Here's a simple example of passing the add-in token to the server-side.</span></span> <span data-ttu-id="68843-185">このトークンは、サーバー側に要求を戻すときの `Authorization` ヘッダーとして含まれています。</span><span class="sxs-lookup"><span data-stu-id="68843-185">The token is included as an `Authorization` header when sending a request back to the server-side.</span></span> <span data-ttu-id="68843-186">この例では JSON データの送信が想定されているので、`POST` メソッドを使用しています。ただし、サーバーに書き込まない場合は、アクセス トークンの送信に `GET` で十分です。</span><span class="sxs-lookup"><span data-stu-id="68843-186">This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

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

#### <a name="when-to-call-the-method"></a><span data-ttu-id="68843-187">メソッドを呼び出すタイミング</span><span class="sxs-lookup"><span data-stu-id="68843-187">When to call the method</span></span>

<span data-ttu-id="68843-188">Office にログインしているユーザーがいないときにアドインを使用できない場合、*アドインを起動するときに* `getAccessTokenAsync` を呼び出す必要があります。</span><span class="sxs-lookup"><span data-stu-id="68843-188">If your add-in cannot be used when no user is logged into Office, then you should call `getAccessTokenAsync` *when the add-in launches*.</span></span>

<span data-ttu-id="68843-189">アドインに、ログイン ユーザーを必要としない機能がある場合、*ユーザーがログイン ユーザーを必要とするアクションを実行するときに* `getAccessTokenAsync` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="68843-189">If the add-in has some functionality that doesn't require a logged in user, then you call `getAccessTokenAsync` *when the user takes an action that requires a logged in user*.</span></span> <span data-ttu-id="68843-190">`getAccessTokenAsync` の重複呼び出しによってパフォーマンスが大幅に低下することはありません。これは、Office ではアクセス トークンがキャッシュされ、それが期限切れになるまで、</span><span class="sxs-lookup"><span data-stu-id="68843-190">There is no significant performance degradation with redundant calls of `getAccessTokenAsync` because Office caches the access token and will reuse it, until it expires, without making another call to the AAD v.</span></span> <span data-ttu-id="68843-191">`getAccessTokenAsync` が呼び出されても AAD V.2.0 エンドポイントが再度呼び出されずに再利用されるためです。</span><span class="sxs-lookup"><span data-stu-id="68843-191">2.0 endpoint whenever `getAccessTokenAsync` is called.</span></span> <span data-ttu-id="68843-192">このため、`getAccessTokenAsync` の呼び出しを、このトークンが必要とされる場所でアクションを開始するすべての関数とハンドラーに追加できます。</span><span class="sxs-lookup"><span data-stu-id="68843-192">So you can add calls of `getAccessTokenAsync` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="68843-193">サーバー側のコードを追加する</span><span class="sxs-lookup"><span data-stu-id="68843-193">Add server-side code</span></span>

<span data-ttu-id="68843-194">ほとんどの場合、アドインがサーバー側に渡してそこで使用しない場合は、アクセス トークンを取得してもあまり意味はありません。</span><span class="sxs-lookup"><span data-stu-id="68843-194">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="68843-195">アドインでは、次のいくつかのサーバー側のタスクを実行できます。</span><span class="sxs-lookup"><span data-stu-id="68843-195">Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="68843-196">ホストされている使用しているデータベースのユーザー設定を検索するメソッドなど、トークンから抽出されるユーザーに関する情報を使用する 1 つ以上の Web API メソッドを作成します。</span><span class="sxs-lookup"><span data-stu-id="68843-196">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base.</span></span> <span data-ttu-id="68843-197">(以下の「**ID として SSO トークンを使用する**」を参照) 使用する言語とフレームワークによっては、記述する必要のあるコードを簡単に記述できるライブラリが使用できることがあります。</span><span class="sxs-lookup"><span data-stu-id="68843-197">(See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="68843-198">Microsoft Graph データを取得します。</span><span class="sxs-lookup"><span data-stu-id="68843-198">Get Microsoft Graph data.</span></span> <span data-ttu-id="68843-199">サーバー側のコードでは、次に示す操作を実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="68843-199">Your server-side code should do the following:</span></span>

    * <span data-ttu-id="68843-200">アクセス トークンを検証します (以下の「**アクセス トークンを検証する**」を参照)。</span><span class="sxs-lookup"><span data-stu-id="68843-200">Validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="68843-201">Azure AD v2.0 エンドポイントを呼び出して、「代理」フローを開始します。これには、アクセス トークン、ユーザーに関するメタデータ、およびアドインの資格情報 (ID とシークレット) を含めます。</span><span class="sxs-lookup"><span data-stu-id="68843-201">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the access token, some metadata about the user, and the credentials of the add-in (its ID and secret).</span></span> <span data-ttu-id="68843-202">このコンテキストでは、アクセス トークンはブートストラップ トークンと呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="68843-202">In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="68843-203">代理フローから戻される新しいアクセス トークンをキャッシュします。</span><span class="sxs-lookup"><span data-stu-id="68843-203">Cache the new access token that is returned from the on-behalf-of flow.</span></span>
    * <span data-ttu-id="68843-204">新しいトークンを使用して Microsoft Graph からデータを取得します。</span><span class="sxs-lookup"><span data-stu-id="68843-204">Get data from Microsoft Graph by using the new token.</span></span>

 <span data-ttu-id="68843-205">ユーザーの Microsoft Graph のデータへのアクセス許可を取得するには、「[Microsoft Graph への認証](authorize-to-microsoft-graph.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="68843-205">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="68843-206">アクセス トークンを検証する</span><span class="sxs-lookup"><span data-stu-id="68843-206">Validate the access token</span></span>

<span data-ttu-id="68843-207">Web API でアクセス トークンを受信したら、そのアクセス トークンを使用する前に検証する必要があります。</span><span class="sxs-lookup"><span data-stu-id="68843-207">Once the Web API receives the access token, it must validate it before using it.</span></span> <span data-ttu-id="68843-208">このトークンは、JSON Web トークン (JWT) です。そのため、この検証は最も標準的な OAuth でのトークンの検証とまったく同様に動作します。</span><span class="sxs-lookup"><span data-stu-id="68843-208">The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows.</span></span> <span data-ttu-id="68843-209">JWT の検証を処理できるライブラリが複数入手可能ですが、その基本は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="68843-209">There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="68843-210">トークンが整形式であることを確認する</span><span class="sxs-lookup"><span data-stu-id="68843-210">Checking that the token is well-formed</span></span>
- <span data-ttu-id="68843-211">トークンが意図した証明機関から発行されたことを確認する</span><span class="sxs-lookup"><span data-stu-id="68843-211">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="68843-212">トークンが Web API を対象にしていることを確認する</span><span class="sxs-lookup"><span data-stu-id="68843-212">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="68843-213">トークンの検証時には、次のガイドラインに注意してください。</span><span class="sxs-lookup"><span data-stu-id="68843-213">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="68843-214">有効な SSO トークンは Azure 証明機関 `https://login.microsoftonline.com` から発行されます。</span><span class="sxs-lookup"><span data-stu-id="68843-214">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`.</span></span> <span data-ttu-id="68843-215">トークン内の `iss` クレームは、この値で始まっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="68843-215">The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="68843-216">トークンの `aud` パラメーターは、アドインの登録のアプリケーション ID に設定します。</span><span class="sxs-lookup"><span data-stu-id="68843-216">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="68843-217">トークンの `scp` パラメーターは `access_as_user` に設定します。</span><span class="sxs-lookup"><span data-stu-id="68843-217">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="68843-218">ID として SSO トークンを使用する</span><span class="sxs-lookup"><span data-stu-id="68843-218">Using the SSO token as an identity</span></span>

<span data-ttu-id="68843-219">アドインでユーザーの ID を検証する必要がある場合、SSO トークンには ID を確定するために使用できる情報が含まれています。</span><span class="sxs-lookup"><span data-stu-id="68843-219">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity.</span></span> <span data-ttu-id="68843-220">ID に関連するトークン内のクレームは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="68843-220">The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="68843-221">`name`: ユーザーの表示名。</span><span class="sxs-lookup"><span data-stu-id="68843-221">`name` - The user's display name.</span></span>
- <span data-ttu-id="68843-222">`preferred_username`: ユーザーの電子メール アドレス。</span><span class="sxs-lookup"><span data-stu-id="68843-222">`preferred_username` - The user's email address.</span></span>
- <span data-ttu-id="68843-223">`oid`: Azure Active Directory でユーザーの ID を表す GUID。</span><span class="sxs-lookup"><span data-stu-id="68843-223">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="68843-224">`tid`: Azure Active Directory でユーザーの組織の ID を表す GUID。</span><span class="sxs-lookup"><span data-stu-id="68843-224">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="68843-225">`name` と `preferred_username` の値は変化することがあるため、この ID とバックエンドの承認サービスを、`oid` と `tid` の値を使用して関連付けることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="68843-225">Since the `name` and `preferred_username` values could change, we recommend that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="68843-226">たとえば、サービスでは、これらの値を `{oid-value}@{tid-value}` のような形式にまとめて、内部ユーザー データベースのユーザーのレコードに値として保存できます。</span><span class="sxs-lookup"><span data-stu-id="68843-226">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database.</span></span> <span data-ttu-id="68843-227">その後の要求では、同じ値を使用してユーザーを取得できるようになり、特定のリソースへのアクセスについては既存のアクセス制御メカニズムに基づいて決定できます。</span><span class="sxs-lookup"><span data-stu-id="68843-227">Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="68843-228">アクセス トークンの例</span><span class="sxs-lookup"><span data-stu-id="68843-228">Example access token</span></span>

<span data-ttu-id="68843-229">アクセス トークンの標準的なデコードされたペイロードを次に示します。</span><span class="sxs-lookup"><span data-stu-id="68843-229">The following is a typical decoded payload of an access token.</span></span> <span data-ttu-id="68843-230">プロパティの詳細については、[Azure Active Directory v2.0 トークンのリファレンス](/azure/active-directory/develop/active-directory-v2-tokens)に関するページを参照してください。</span><span class="sxs-lookup"><span data-stu-id="68843-230">For information about the properties, see [Azure Active Directory v2.0 tokens reference](/azure/active-directory/develop/active-directory-v2-tokens).</span></span>


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

## <a name="using-sso-with-an-outlook-add-in"></a><span data-ttu-id="68843-231">Outlook のアドインで SSO を使用する</span><span class="sxs-lookup"><span data-stu-id="68843-231">Using SSO with an Outlook add-in</span></span>

<span data-ttu-id="68843-232">Excel、PowerPoint、または Word のアドインで SSO を使用する場合と Outlook のアドインでそれを使用する場合とでは、小さいけれど重要な違いがいくつかあります。</span><span class="sxs-lookup"><span data-stu-id="68843-232">There are some small, but important differences in using SSO in an Outlook add-in from using it in an Excel, PowerPoint, or Word add-in.</span></span> <span data-ttu-id="68843-233">「[Authenticate a user with a single sign-on token in an Outlook add-in](/outlook/add-ins/authenticate-a-user-with-an-sso-token)」 (Outlook アドインでシングル サインオン トークンを使用してユーザーを認証する) と「[シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](/outlook/add-ins/implement-sso-in-outlook-add-in)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="68843-233">Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](/outlook/add-ins/authenticate-a-user-with-an-sso-token) and [Scenario: Implement single sign-on to your service in an Outlook add-in](/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

## <a name="sso-api-reference"></a><span data-ttu-id="68843-234">SSO API リファレンス</span><span class="sxs-lookup"><span data-stu-id="68843-234">SSO API reference</span></span>

### <a name="getaccesstokenasync"></a><span data-ttu-id="68843-235">getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="68843-235">getAccessTokenAsync</span></span>

<span data-ttu-id="68843-236">Office Auth 名前空間 (`Office.context.auth`) には、Office ホストがアドインの Web アプリケーションへのアクセス トークンを取得することを可能にする `getAccessTokenAsync` というメソッドがあります。</span><span class="sxs-lookup"><span data-stu-id="68843-236">The Office Auth namespace, `Office.context.auth`, provides a method, `getAccessTokenAsync` that enables the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="68843-237">これにより、間接的に、サインインしたユーザーの Microsoft Graph データにアドインがアクセスできるようにもなります。ユーザーがもう一度サインインする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="68843-237">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

<span data-ttu-id="68843-238">このメソッドは、Azure Active Directory V 2.0 のエンドポイントを呼び出して、アドインの Web アプリケーションへのアクセス トークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="68843-238">The method calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application.</span></span> <span data-ttu-id="68843-239">これにより、アドインがユーザーを識別できるようになります。</span><span class="sxs-lookup"><span data-stu-id="68843-239">This enables add-ins to identify users.</span></span> <span data-ttu-id="68843-240">["on behalf of" OAuth フロー](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)を使用することにより、サーバー側のコードはこのトークンを使用して、アドインの Web アプリケーションの Microsoft Graph にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="68843-240">Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span>

> [!NOTE]
> <span data-ttu-id="68843-241">Outlook でアドインが Outlook.com または Gmail のメールボックスに読み込まれている場合、この API はサポートされません。</span><span class="sxs-lookup"><span data-stu-id="68843-241">In Outlook, this API is not supported if the add-in is loaded in an Outlook.com or Gmail mailbox.</span></span>

|<span data-ttu-id="68843-242">ホスト</span><span class="sxs-lookup"><span data-stu-id="68843-242">Hosts</span></span>|<span data-ttu-id="68843-243">Excel、OneNote、Outlook、PowerPoint、Word</span><span class="sxs-lookup"><span data-stu-id="68843-243">Excel, OneNote, Outlook, PowerPoint, Word</span></span>|
|---|---|
|[<span data-ttu-id="68843-244">要件セット</span><span class="sxs-lookup"><span data-stu-id="68843-244">Requirement sets</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)|[<span data-ttu-id="68843-245">IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="68843-245">IdentityAPI</span></span>](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)|

#### <a name="parameters"></a><span data-ttu-id="68843-246">パラメーター</span><span class="sxs-lookup"><span data-stu-id="68843-246">Parameters</span></span>

<span data-ttu-id="68843-247">`options` - 省略可能。</span><span class="sxs-lookup"><span data-stu-id="68843-247">`options` - Optional.</span></span> <span data-ttu-id="68843-248">サインイン動作を定義する `AuthOptions` オブジェクトが許可されます (下記参照)。</span><span class="sxs-lookup"><span data-stu-id="68843-248">Accepts an `AuthOptions` object (see below) to define sign-on behaviors.</span></span>

<span data-ttu-id="68843-249">`callback` - 省略可能。</span><span class="sxs-lookup"><span data-stu-id="68843-249">`callback` - Optional.</span></span> <span data-ttu-id="68843-250">ユーザー ID 用のトークンを解析できるコールバック メソッドが許可されます。または、トークンを Microsoft Graph へのアクセスを取得するために、「代理」フローで使用します。</span><span class="sxs-lookup"><span data-stu-id="68843-250">Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph.</span></span> <span data-ttu-id="68843-251">[AsyncResult](/javascript/api/office/office.asyncresult)`.status` が "succeeded" である場合、`AsyncResult.value` が生の AAD v.</span><span class="sxs-lookup"><span data-stu-id="68843-251">If [AsyncResult](/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v.</span></span> <span data-ttu-id="68843-252">2.0 形式のアクセス トークンになります。</span><span class="sxs-lookup"><span data-stu-id="68843-252">2.0-formatted access token.</span></span>

<span data-ttu-id="68843-253">`AuthOptions` インターフェイスには、Office が</span><span class="sxs-lookup"><span data-stu-id="68843-253">The `AuthOptions` interface provides options for the user experience when Office obtains an access token to the add-in from AAD v.</span></span> <span data-ttu-id="68843-254">`getAccessTokenAsync` メソッドを使用して AAD v. 2.0 からアドインに対するアクセス トークンを取得する場合用のユーザー エクスペリエンス用のオプションがあります。</span><span class="sxs-lookup"><span data-stu-id="68843-254">2.0 with the `getAccessTokenAsync` method.</span></span>

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
