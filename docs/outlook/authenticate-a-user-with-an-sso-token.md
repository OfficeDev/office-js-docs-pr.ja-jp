---
title: シングル サインオン トークンを使用してユーザーを認証する
description: サービスに SSO を実装するために Outlook アドインが提供するシングル サインオン トークンを使用することについて説明します。
ms.date: 11/19/2019
localization_priority: Normal
ms.openlocfilehash: 9ee3ece5929df602a35ddd9883c08e25164d8a22
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42721031"
---
# <a name="authenticate-a-user-with-a-single-sign-on-token-in-an-outlook-add-in-preview"></a><span data-ttu-id="58753-103">Outlook アドインでシングル サインオン トークンを使用してユーザーを認証する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="58753-103">Authenticate a user with a single-sign-on token in an Outlook add-in (preview)</span></span>

<span data-ttu-id="58753-104">シングル サインオン (SSO) は、アドインがユーザーを認証する (またオプションでアクセス トークンを認証および取得して [Microsoft Graph API](/graph/overview) を呼び出す) ための、シームレスな方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="58753-104">Single sign-on (SSO) provides a seamless way for your add-in to authenticate users (and optionally to obtain access tokens to call the [Microsoft Graph API](/graph/overview)).</span></span>

<span data-ttu-id="58753-105">この方法を使用すると、アドインはサーバーのバックエンド API にスコープされたアクセス トークンを取得できます。</span><span class="sxs-lookup"><span data-stu-id="58753-105">Using this method, your add-in can obtain an access token scoped to your server back-end API.</span></span> <span data-ttu-id="58753-106">アドインはこれを `Authorization` ヘッダーのベアラー トークンとして使用して、API へのコールバックを認証します。</span><span class="sxs-lookup"><span data-stu-id="58753-106">The add-in uses this as a bearer token in the `Authorization` header to authenticate a call back to your API.</span></span> <span data-ttu-id="58753-107">オプションとして、サーバー側のコードも持つことができます。</span><span class="sxs-lookup"><span data-stu-id="58753-107">Optionally, you can also have your server-side code:</span></span>

- <span data-ttu-id="58753-108">On-Behalf-Of フローを完了して、Microsoft Graph API にスコープ設定されたアクセス トークンを取得する</span><span class="sxs-lookup"><span data-stu-id="58753-108">Complete the On-Behalf-Of flow to obtain an access token scoped to the Microsoft Graph API</span></span>
- <span data-ttu-id="58753-109">トークン内の ID 情報を使用して、独自のバックエンド サービスに対するユーザーの識別と認証を確立する</span><span class="sxs-lookup"><span data-stu-id="58753-109">Use the identity information in the token to establish the user's identity and authenticate to your own back-end services</span></span>

<span data-ttu-id="58753-110">Office アドインの SSO の概要については、[「Office アドインのシングル サインオンを有効化する」](../develop/sso-in-office-add-ins.md) および[「Office アドインの Microsoft Graph への承認」](../develop/authorize-to-microsoft-graph.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="58753-110">For an overview of SSO in Office Add-ins, see [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md) and [Authorize to Microsoft Graph in your Office Add-in](../develop/authorize-to-microsoft-graph.md).</span></span>

> [!NOTE]
> <span data-ttu-id="58753-111">SSO を使用するには、アドインのスタートアップ HTML ページの https://appsforoffice.microsoft.com/lib/beta/hosted/office.js から Office JavaScript ライブラリのベータ版を読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="58753-111">To use SSO, you must load the beta version of the Office JavaScript Library from https://appsforoffice.microsoft.com/lib/beta/hosted/office.js in the startup HTML page of the add-in.</span></span>

## <a name="enable-modern-authentication-in-your-office-365-tenancy"></a><span data-ttu-id="58753-112">Office 365 テナントで先進認証を有効にする</span><span class="sxs-lookup"><span data-stu-id="58753-112">Enable modern authentication in your Office 365 tenancy</span></span>

<span data-ttu-id="58753-113">Outlook アドインで SSO を使用するには、Office 365 テナントの先進認証を有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="58753-113">To use SSO with an Outlook add-in, you must enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="58753-114">この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="58753-114">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

## <a name="register-your-add-in"></a><span data-ttu-id="58753-115">アドインを登録する</span><span class="sxs-lookup"><span data-stu-id="58753-115">Register your add-in</span></span>

<span data-ttu-id="58753-116">SSO を使用するには、Outlook アドインに Azure Active Directory (AAD) v2.0 を登録したサーバー側ウェブ API が必要です。</span><span class="sxs-lookup"><span data-stu-id="58753-116">To use SSO, your Outlook add-in will need to have a server-side web API that is registered with Azure Active Directory (AAD) v2.0.</span></span> <span data-ttu-id="58753-117">詳細については、「[Azure AD v2.0 のエンドポイントで SSO を使用する Office アドインを登録する](../develop/register-sso-add-in-aad-v2.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="58753-117">For more information, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](../develop/register-sso-add-in-aad-v2.md).</span></span>

### <a name="provide-consent-when-sideloading-an-add-in"></a><span data-ttu-id="58753-118">アドインのサイドロード時に同意する</span><span class="sxs-lookup"><span data-stu-id="58753-118">Provide consent when sideloading an add-in</span></span>

<span data-ttu-id="58753-119">SSO を使用するアドインを AppSource から取得するときに、ストア UI のダイアログが表示され、ユーザーに対して要求される Graph のアクセス許可の同意を求めます。</span><span class="sxs-lookup"><span data-stu-id="58753-119">When an add-in that uses SSO is acquired from AppSource, the store UI handles prompting the user for consent to the requested Graph permissions.</span></span> <span data-ttu-id="58753-120">ただし、アドインを開発する際には事前に同意を提示する必要があります。</span><span class="sxs-lookup"><span data-stu-id="58753-120">However, when you are developing an add-in, you have to provide consent in advance.</span></span> <span data-ttu-id="58753-121">詳細については、「[アドインに管理者の同意を付与する](../develop/grant-admin-consent-to-an-add-in.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="58753-121">For more information, see [Grant administrator consent to the add-in](../develop/grant-admin-consent-to-an-add-in.md)</span></span>

## <a name="update-the-add-in-manifest"></a><span data-ttu-id="58753-122">アドイン マニフェストを更新する</span><span class="sxs-lookup"><span data-stu-id="58753-122">Update the add-in manifest</span></span>

<span data-ttu-id="58753-123">アドインで SSO を使用できるようにするための次の手順では、`VersionOverridesV1_1` の [VersionOverrides](../reference/manifest/versionoverrides.md) 要素の最後に `WebApplicationInfo` 要素を追加します。</span><span class="sxs-lookup"><span data-stu-id="58753-123">The next step to enable SSO in the add-in is to add a `WebApplicationInfo` element at the end of the `VersionOverridesV1_1` [VersionOverrides](../reference/manifest/versionoverrides.md) element.</span></span> <span data-ttu-id="58753-124">詳細については、「[アドインを構成する](../develop/sso-in-office-add-ins.md#configure-the-add-in)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="58753-124">For more information, see [Configure the add-in](../develop/sso-in-office-add-ins.md#configure-the-add-in).</span></span>

## <a name="get-the-sso-token"></a><span data-ttu-id="58753-125">SSO トークンを取得する</span><span class="sxs-lookup"><span data-stu-id="58753-125">Get the SSO token</span></span>

<span data-ttu-id="58753-126">アドインがクライアント側スクリプトを含む SSO トークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="58753-126">The add-in gets an SSO token with client-side script.</span></span> <span data-ttu-id="58753-127">詳細については、[「クライアント側のコードを追加する」](../develop/sso-in-office-add-ins.md#add-client-side-code)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="58753-127">For more information, see [Add client-side code](../develop/sso-in-office-add-ins.md#add-client-side-code).</span></span>

## <a name="use-the-sso-token-at-the-back-end"></a><span data-ttu-id="58753-128">バックエンドで SSO トークンを使用する</span><span class="sxs-lookup"><span data-stu-id="58753-128">Use the SSO token at the back-end</span></span>

<span data-ttu-id="58753-129">ほとんどの場合、アドインがサーバー側に渡してそこで使用しない場合は、アクセス トークンを取得してもあまり意味はありません。</span><span class="sxs-lookup"><span data-stu-id="58753-129">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="58753-130">サーバー側で可能となること、また必要な対応に関する詳細は、[「サーバー側のコードを追加する」](../develop/sso-in-office-add-ins.md#add-server-side-code)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="58753-130">For details on what your server-side could and should do, see [Add server-side code](../develop/sso-in-office-add-ins.md#add-server-side-code).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="58753-131">ID として SSO トークンを *Outlook* アドインで使用するときには、代替の ID として [Exchange の ID トークンも使用](authenticate-a-user-with-an-identity-token.md)することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="58753-131">When using the SSO token as an identity in an *Outlook* add-in, we recommend that you also [use the Exchange identity token](authenticate-a-user-with-an-identity-token.md) as an alternate identity.</span></span> <span data-ttu-id="58753-132">アドインのユーザーは、複数のクライアントを使用することがあり、一部のクライアントは SSO トークンの提示をサポートしていないことがあります。</span><span class="sxs-lookup"><span data-stu-id="58753-132">Users of your add-in may use multiple clients, and some may not support providing an SSO token.</span></span> <span data-ttu-id="58753-133">代わりに Exchange の ID トークン使用すると、そうしたユーザーに資格情報の入力を求めるダイアログを複数回表示しないようにできます。</span><span class="sxs-lookup"><span data-stu-id="58753-133">By using the Exchange identity token as an alternate, you can avoid having to prompt these users for credentials multiple times.</span></span> <span data-ttu-id="58753-134">詳細については、「[シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](implement-sso-in-outlook-add-in.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="58753-134">For more information, see [Scenario: Implement single sign-on to your service in an Outlook add-in](implement-sso-in-outlook-add-in.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="58753-135">関連項目</span><span class="sxs-lookup"><span data-stu-id="58753-135">See also</span></span>

- <span data-ttu-id="58753-136">Microsoft Graph API へのアクセスに SSO トークンを使用するサンプル Outlook アドインについては、[「AttachmentsDemo サンプル アドイン」](https://github.com/OfficeDev/outlook-add-in-attachments-demo)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="58753-136">For a sample Outlook add-in that uses the SSO token to access the Microsoft Graph API, see [AttachmentsDemo Sample Add-in](https://github.com/OfficeDev/outlook-add-in-attachments-demo).</span></span>
- [<span data-ttu-id="58753-137">SSO API リファレンス</span><span class="sxs-lookup"><span data-stu-id="58753-137">SSO API reference</span></span>](../develop/sso-in-office-add-ins.md#sso-api-reference)
- [<span data-ttu-id="58753-138">IdentityAPI 要件セット</span><span class="sxs-lookup"><span data-stu-id="58753-138">IdentityAPI requirement set</span></span>](../reference/requirement-sets/identity-api-requirement-sets.md)
