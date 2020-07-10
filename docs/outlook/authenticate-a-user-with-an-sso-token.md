---
title: シングル サインオン トークンを使用してユーザーを認証する
description: サービスに SSO を実装するために Outlook アドインが提供するシングル サインオン トークンを使用することについて説明します。
ms.date: 04/28/2020
localization_priority: Normal
ms.openlocfilehash: d53e75faa2d0471b43957cfa71ff6f6a50a0da4f
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093981"
---
# <a name="authenticate-a-user-with-a-single-sign-on-token-in-an-outlook-add-in-preview"></a><span data-ttu-id="f25c7-103">Outlook アドインでシングル サインオン トークンを使用してユーザーを認証する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="f25c7-103">Authenticate a user with a single-sign-on token in an Outlook add-in (preview)</span></span>

<span data-ttu-id="f25c7-104">シングル サインオン (SSO) は、アドインがユーザーを認証する (またオプションでアクセス トークンを認証および取得して [Microsoft Graph API](/graph/overview) を呼び出す) ための、シームレスな方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="f25c7-104">Single sign-on (SSO) provides a seamless way for your add-in to authenticate users (and optionally to obtain access tokens to call the [Microsoft Graph API](/graph/overview)).</span></span>

<span data-ttu-id="f25c7-105">この方法を使用すると、アドインはサーバーのバックエンド API にスコープされたアクセス トークンを取得できます。</span><span class="sxs-lookup"><span data-stu-id="f25c7-105">Using this method, your add-in can obtain an access token scoped to your server back-end API.</span></span> <span data-ttu-id="f25c7-106">アドインはこれを `Authorization` ヘッダーのベアラー トークンとして使用して、API へのコールバックを認証します。</span><span class="sxs-lookup"><span data-stu-id="f25c7-106">The add-in uses this as a bearer token in the `Authorization` header to authenticate a call back to your API.</span></span> <span data-ttu-id="f25c7-107">オプションとして、サーバー側のコードも持つことができます。</span><span class="sxs-lookup"><span data-stu-id="f25c7-107">Optionally, you can also have your server-side code:</span></span>

- <span data-ttu-id="f25c7-108">On-Behalf-Of フローを完了して、Microsoft Graph API にスコープ設定されたアクセス トークンを取得する</span><span class="sxs-lookup"><span data-stu-id="f25c7-108">Complete the On-Behalf-Of flow to obtain an access token scoped to the Microsoft Graph API</span></span>
- <span data-ttu-id="f25c7-109">トークン内の ID 情報を使用して、独自のバックエンド サービスに対するユーザーの識別と認証を確立する</span><span class="sxs-lookup"><span data-stu-id="f25c7-109">Use the identity information in the token to establish the user's identity and authenticate to your own back-end services</span></span>

<span data-ttu-id="f25c7-110">Office アドインの SSO の概要については、[「Office アドインのシングル サインオンを有効化する」](../develop/sso-in-office-add-ins.md) および[「Office アドインの Microsoft Graph への承認」](../develop/authorize-to-microsoft-graph.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f25c7-110">For an overview of SSO in Office Add-ins, see [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md) and [Authorize to Microsoft Graph in your Office Add-in](../develop/authorize-to-microsoft-graph.md).</span></span>

> [!NOTE]
> <span data-ttu-id="f25c7-111">SSO を使用するには、アドインのスタートアップ HTML ページの https://appsforoffice.microsoft.com/lib/beta/hosted/office.js から Office JavaScript ライブラリのベータ版を読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="f25c7-111">To use SSO, you must load the beta version of the Office JavaScript Library from https://appsforoffice.microsoft.com/lib/beta/hosted/office.js in the startup HTML page of the add-in.</span></span> <span data-ttu-id="f25c7-112">ただし、運用環境のアドインではベータ版の Api**を使用しないでください**。</span><span class="sxs-lookup"><span data-stu-id="f25c7-112">However, you should **not** use beta APIs in production add-ins.</span></span>

## <a name="enable-modern-authentication-in-your-microsoft-365-tenancy"></a><span data-ttu-id="f25c7-113">Microsoft 365 テナントで先進認証を有効にする</span><span class="sxs-lookup"><span data-stu-id="f25c7-113">Enable modern authentication in your Microsoft 365 tenancy</span></span>

<span data-ttu-id="f25c7-114">Outlook アドインで SSO を使用するには、Microsoft 365 テナントの先進認証を有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="f25c7-114">To use SSO with an Outlook add-in, you must enable Modern Authentication for the Microsoft 365 tenancy.</span></span> <span data-ttu-id="f25c7-115">この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f25c7-115">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

## <a name="register-your-add-in"></a><span data-ttu-id="f25c7-116">アドインを登録する</span><span class="sxs-lookup"><span data-stu-id="f25c7-116">Register your add-in</span></span>

<span data-ttu-id="f25c7-117">SSO を使用するには、Outlook アドインに Azure Active Directory (AAD) v2.0 を登録したサーバー側ウェブ API が必要です。</span><span class="sxs-lookup"><span data-stu-id="f25c7-117">To use SSO, your Outlook add-in will need to have a server-side web API that is registered with Azure Active Directory (AAD) v2.0.</span></span> <span data-ttu-id="f25c7-118">詳細については、「[Azure AD v2.0 のエンドポイントで SSO を使用する Office アドインを登録する](../develop/register-sso-add-in-aad-v2.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="f25c7-118">For more information, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](../develop/register-sso-add-in-aad-v2.md).</span></span>

### <a name="provide-consent-when-sideloading-an-add-in"></a><span data-ttu-id="f25c7-119">アドインのサイドロード時に同意する</span><span class="sxs-lookup"><span data-stu-id="f25c7-119">Provide consent when sideloading an add-in</span></span>

<span data-ttu-id="f25c7-120">SSO を使用するアドインを AppSource から取得するときに、ストア UI のダイアログが表示され、ユーザーに対して要求される Graph のアクセス許可の同意を求めます。</span><span class="sxs-lookup"><span data-stu-id="f25c7-120">When an add-in that uses SSO is acquired from AppSource, the store UI handles prompting the user for consent to the requested Graph permissions.</span></span> <span data-ttu-id="f25c7-121">ただし、アドインを開発する際には事前に同意を提示する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f25c7-121">However, when you are developing an add-in, you have to provide consent in advance.</span></span> <span data-ttu-id="f25c7-122">詳細については、「[アドインに管理者の同意を付与する](../develop/grant-admin-consent-to-an-add-in.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f25c7-122">For more information, see [Grant administrator consent to the add-in](../develop/grant-admin-consent-to-an-add-in.md)</span></span>

## <a name="update-the-add-in-manifest"></a><span data-ttu-id="f25c7-123">アドイン マニフェストを更新する</span><span class="sxs-lookup"><span data-stu-id="f25c7-123">Update the add-in manifest</span></span>

<span data-ttu-id="f25c7-124">アドインで SSO を使用できるようにするための次の手順では、`VersionOverridesV1_1` の [VersionOverrides](../reference/manifest/versionoverrides.md) 要素の最後に `WebApplicationInfo` 要素を追加します。</span><span class="sxs-lookup"><span data-stu-id="f25c7-124">The next step to enable SSO in the add-in is to add a `WebApplicationInfo` element at the end of the `VersionOverridesV1_1` [VersionOverrides](../reference/manifest/versionoverrides.md) element.</span></span> <span data-ttu-id="f25c7-125">詳細については、「[アドインを構成する](../develop/sso-in-office-add-ins.md#configure-the-add-in)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f25c7-125">For more information, see [Configure the add-in](../develop/sso-in-office-add-ins.md#configure-the-add-in).</span></span>

## <a name="get-the-sso-token"></a><span data-ttu-id="f25c7-126">SSO トークンを取得する</span><span class="sxs-lookup"><span data-stu-id="f25c7-126">Get the SSO token</span></span>

<span data-ttu-id="f25c7-127">アドインがクライアント側スクリプトを含む SSO トークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="f25c7-127">The add-in gets an SSO token with client-side script.</span></span> <span data-ttu-id="f25c7-128">詳細については、[「クライアント側のコードを追加する」](../develop/sso-in-office-add-ins.md#add-client-side-code)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f25c7-128">For more information, see [Add client-side code](../develop/sso-in-office-add-ins.md#add-client-side-code).</span></span>

## <a name="use-the-sso-token-at-the-back-end"></a><span data-ttu-id="f25c7-129">バックエンドで SSO トークンを使用する</span><span class="sxs-lookup"><span data-stu-id="f25c7-129">Use the SSO token at the back-end</span></span>

<span data-ttu-id="f25c7-130">ほとんどの場合、アドインがサーバー側に渡してそこで使用しない場合は、アクセス トークンを取得してもあまり意味はありません。</span><span class="sxs-lookup"><span data-stu-id="f25c7-130">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="f25c7-131">サーバー側で可能となること、また必要な対応に関する詳細は、[「サーバー側のコードを追加する」](../develop/sso-in-office-add-ins.md#add-server-side-code)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f25c7-131">For details on what your server-side could and should do, see [Add server-side code](../develop/sso-in-office-add-ins.md#add-server-side-code).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f25c7-132">ID として SSO トークンを *Outlook* アドインで使用するときには、代替の ID として [Exchange の ID トークンも使用](authenticate-a-user-with-an-identity-token.md)することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="f25c7-132">When using the SSO token as an identity in an *Outlook* add-in, we recommend that you also [use the Exchange identity token](authenticate-a-user-with-an-identity-token.md) as an alternate identity.</span></span> <span data-ttu-id="f25c7-133">アドインのユーザーは、複数のクライアントを使用することがあり、一部のクライアントは SSO トークンの提示をサポートしていないことがあります。</span><span class="sxs-lookup"><span data-stu-id="f25c7-133">Users of your add-in may use multiple clients, and some may not support providing an SSO token.</span></span> <span data-ttu-id="f25c7-134">代わりに Exchange の ID トークン使用すると、そうしたユーザーに資格情報の入力を求めるダイアログを複数回表示しないようにできます。</span><span class="sxs-lookup"><span data-stu-id="f25c7-134">By using the Exchange identity token as an alternate, you can avoid having to prompt these users for credentials multiple times.</span></span> <span data-ttu-id="f25c7-135">詳細については、「[シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](implement-sso-in-outlook-add-in.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f25c7-135">For more information, see [Scenario: Implement single sign-on to your service in an Outlook add-in](implement-sso-in-outlook-add-in.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="f25c7-136">関連項目</span><span class="sxs-lookup"><span data-stu-id="f25c7-136">See also</span></span>

- <span data-ttu-id="f25c7-137">Microsoft Graph API へのアクセスに SSO トークンを使用するサンプル Outlook アドインについては、[「AttachmentsDemo サンプル アドイン」](https://github.com/OfficeDev/outlook-add-in-attachments-demo)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f25c7-137">For a sample Outlook add-in that uses the SSO token to access the Microsoft Graph API, see [AttachmentsDemo Sample Add-in](https://github.com/OfficeDev/outlook-add-in-attachments-demo).</span></span>
- [<span data-ttu-id="f25c7-138">SSO API リファレンス</span><span class="sxs-lookup"><span data-stu-id="f25c7-138">SSO API reference</span></span>](../develop/sso-in-office-add-ins.md#sso-api-reference)
- [<span data-ttu-id="f25c7-139">IdentityAPI 要件セット</span><span class="sxs-lookup"><span data-stu-id="f25c7-139">IdentityAPI requirement set</span></span>](../reference/requirement-sets/identity-api-requirement-sets.md)
