---
title: シングル サインオン トークンを使用してユーザーを認証する
description: サービスに SSO を実装するために Outlook アドインが提供するシングル サインオン トークンを使用することについて説明します。
ms.date: 08/20/2020
localization_priority: Normal
ms.openlocfilehash: e0925979d26f6b3145658d71b1edaf30431e0c7e
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293983"
---
# <a name="authenticate-a-user-with-a-single-sign-on-token-in-an-outlook-add-in"></a><span data-ttu-id="8d6c1-103">Outlook アドインでシングルサインオントークンを使用してユーザーを認証する</span><span class="sxs-lookup"><span data-stu-id="8d6c1-103">Authenticate a user with a single-sign-on token in an Outlook add-in</span></span>

<span data-ttu-id="8d6c1-104">シングル サインオン (SSO) は、アドインがユーザーを認証する (またオプションでアクセス トークンを認証および取得して [Microsoft Graph API](/graph/overview) を呼び出す) ための、シームレスな方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-104">Single sign-on (SSO) provides a seamless way for your add-in to authenticate users (and optionally to obtain access tokens to call the [Microsoft Graph API](/graph/overview)).</span></span>

<span data-ttu-id="8d6c1-105">この方法を使用すると、アドインはサーバーのバックエンド API にスコープされたアクセス トークンを取得できます。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-105">Using this method, your add-in can obtain an access token scoped to your server back-end API.</span></span> <span data-ttu-id="8d6c1-106">アドインはこれを `Authorization` ヘッダーのベアラー トークンとして使用して、API へのコールバックを認証します。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-106">The add-in uses this as a bearer token in the `Authorization` header to authenticate a call back to your API.</span></span> <span data-ttu-id="8d6c1-107">オプションとして、サーバー側のコードも持つことができます。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-107">Optionally, you can also have your server-side code:</span></span>

- <span data-ttu-id="8d6c1-108">On-Behalf-Of フローを完了して、Microsoft Graph API にスコープ設定されたアクセス トークンを取得する</span><span class="sxs-lookup"><span data-stu-id="8d6c1-108">Complete the On-Behalf-Of flow to obtain an access token scoped to the Microsoft Graph API</span></span>
- <span data-ttu-id="8d6c1-109">トークン内の ID 情報を使用して、独自のバックエンド サービスに対するユーザーの識別と認証を確立する</span><span class="sxs-lookup"><span data-stu-id="8d6c1-109">Use the identity information in the token to establish the user's identity and authenticate to your own back-end services</span></span>

<span data-ttu-id="8d6c1-110">Office アドインの SSO の概要については、[「Office アドインのシングル サインオンを有効化する」](../develop/sso-in-office-add-ins.md) および[「Office アドインの Microsoft Graph への承認」](../develop/authorize-to-microsoft-graph.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-110">For an overview of SSO in Office Add-ins, see [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md) and [Authorize to Microsoft Graph in your Office Add-in](../develop/authorize-to-microsoft-graph.md).</span></span>

## <a name="enable-modern-authentication-in-your-microsoft-365-tenancy"></a><span data-ttu-id="8d6c1-111">Microsoft 365 テナントで先進認証を有効にする</span><span class="sxs-lookup"><span data-stu-id="8d6c1-111">Enable modern authentication in your Microsoft 365 tenancy</span></span>

<span data-ttu-id="8d6c1-112">Outlook アドインで SSO を使用するには、Microsoft 365 テナントの先進認証を有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-112">To use SSO with an Outlook add-in, you must enable Modern Authentication for the Microsoft 365 tenancy.</span></span> <span data-ttu-id="8d6c1-113">この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-113">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

## <a name="register-your-add-in"></a><span data-ttu-id="8d6c1-114">アドインを登録する</span><span class="sxs-lookup"><span data-stu-id="8d6c1-114">Register your add-in</span></span>

<span data-ttu-id="8d6c1-115">SSO を使用するには、Outlook アドインに Azure Active Directory (AAD) v2.0 を登録したサーバー側ウェブ API が必要です。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-115">To use SSO, your Outlook add-in will need to have a server-side web API that is registered with Azure Active Directory (AAD) v2.0.</span></span> <span data-ttu-id="8d6c1-116">詳細については、「[Azure AD v2.0 のエンドポイントで SSO を使用する Office アドインを登録する](../develop/register-sso-add-in-aad-v2.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-116">For more information, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](../develop/register-sso-add-in-aad-v2.md).</span></span>

### <a name="provide-consent-when-sideloading-an-add-in"></a><span data-ttu-id="8d6c1-117">アドインのサイドロード時に同意する</span><span class="sxs-lookup"><span data-stu-id="8d6c1-117">Provide consent when sideloading an add-in</span></span>

<span data-ttu-id="8d6c1-118">アドインを開発している場合は、事前に同意を得る必要があります。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-118">When you are developing an add-in, you will have to provide consent in advance.</span></span> <span data-ttu-id="8d6c1-119">詳細については、「 [アドインに管理者の同意を付与する](../develop/grant-admin-consent-to-an-add-in.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-119">For more information, see [Grant administrator consent to the add-in](../develop/grant-admin-consent-to-an-add-in.md).</span></span>

## <a name="update-the-add-in-manifest"></a><span data-ttu-id="8d6c1-120">アドイン マニフェストを更新する</span><span class="sxs-lookup"><span data-stu-id="8d6c1-120">Update the add-in manifest</span></span>

<span data-ttu-id="8d6c1-121">アドインで SSO を使用できるようにするための次の手順では、`VersionOverridesV1_1` の [VersionOverrides](../reference/manifest/versionoverrides.md) 要素の最後に `WebApplicationInfo` 要素を追加します。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-121">The next step to enable SSO in the add-in is to add a `WebApplicationInfo` element at the end of the `VersionOverridesV1_1` [VersionOverrides](../reference/manifest/versionoverrides.md) element.</span></span> <span data-ttu-id="8d6c1-122">詳細については、「[アドインを構成する](../develop/sso-in-office-add-ins.md#configure-the-add-in)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-122">For more information, see [Configure the add-in](../develop/sso-in-office-add-ins.md#configure-the-add-in).</span></span>

## <a name="get-the-sso-token"></a><span data-ttu-id="8d6c1-123">SSO トークンを取得する</span><span class="sxs-lookup"><span data-stu-id="8d6c1-123">Get the SSO token</span></span>

<span data-ttu-id="8d6c1-124">アドインがクライアント側スクリプトを含む SSO トークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-124">The add-in gets an SSO token with client-side script.</span></span> <span data-ttu-id="8d6c1-125">詳細については、[「クライアント側のコードを追加する」](../develop/sso-in-office-add-ins.md#add-client-side-code)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-125">For more information, see [Add client-side code](../develop/sso-in-office-add-ins.md#add-client-side-code).</span></span>

## <a name="use-the-sso-token-at-the-back-end"></a><span data-ttu-id="8d6c1-126">バックエンドで SSO トークンを使用する</span><span class="sxs-lookup"><span data-stu-id="8d6c1-126">Use the SSO token at the back-end</span></span>

<span data-ttu-id="8d6c1-127">ほとんどの場合、アドインがサーバー側に渡してそこで使用しない場合は、アクセス トークンを取得してもあまり意味はありません。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-127">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="8d6c1-128">サーバー側で可能となること、また必要な対応に関する詳細は、[「サーバー側のコードを追加する」](../develop/sso-in-office-add-ins.md#add-server-side-code)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-128">For details on what your server-side could and should do, see [Add server-side code](../develop/sso-in-office-add-ins.md#add-server-side-code).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8d6c1-129">ID として SSO トークンを *Outlook* アドインで使用するときには、代替の ID として [Exchange の ID トークンも使用](authenticate-a-user-with-an-identity-token.md)することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-129">When using the SSO token as an identity in an *Outlook* add-in, we recommend that you also [use the Exchange identity token](authenticate-a-user-with-an-identity-token.md) as an alternate identity.</span></span> <span data-ttu-id="8d6c1-130">アドインのユーザーは、複数のクライアントを使用することがあり、一部のクライアントは SSO トークンの提示をサポートしていないことがあります。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-130">Users of your add-in may use multiple clients, and some may not support providing an SSO token.</span></span> <span data-ttu-id="8d6c1-131">代わりに Exchange の ID トークン使用すると、そうしたユーザーに資格情報の入力を求めるダイアログを複数回表示しないようにできます。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-131">By using the Exchange identity token as an alternate, you can avoid having to prompt these users for credentials multiple times.</span></span> <span data-ttu-id="8d6c1-132">詳細については、「[シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](implement-sso-in-outlook-add-in.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-132">For more information, see [Scenario: Implement single sign-on to your service in an Outlook add-in](implement-sso-in-outlook-add-in.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="8d6c1-133">関連項目</span><span class="sxs-lookup"><span data-stu-id="8d6c1-133">See also</span></span>

- <span data-ttu-id="8d6c1-134">SSO トークンを使用して Microsoft Graph API にアクセスする Outlook アドインのサンプルについては、「 [Outlook アドイン SSO](https://github.com/OfficeDev/Outlook-Add-in-SSO)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8d6c1-134">For a sample Outlook add-in that uses the SSO token to access the Microsoft Graph API, see [Outlook Add-in SSO](https://github.com/OfficeDev/Outlook-Add-in-SSO).</span></span>
- [<span data-ttu-id="8d6c1-135">SSO API リファレンス</span><span class="sxs-lookup"><span data-stu-id="8d6c1-135">SSO API reference</span></span>](../develop/sso-in-office-add-ins.md#sso-api-reference)
- [<span data-ttu-id="8d6c1-136">IdentityAPI 要件セット</span><span class="sxs-lookup"><span data-stu-id="8d6c1-136">IdentityAPI requirement set</span></span>](../reference/requirement-sets/identity-api-requirement-sets.md)
