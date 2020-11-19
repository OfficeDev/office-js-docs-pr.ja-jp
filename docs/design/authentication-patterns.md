---
title: Office アドインの認証設計ガイドライン
ms.date: 07/30/2020
description: Office アドインでサインオンまたはサインアップページを視覚的に設計する方法について説明します。
localization_priority: Normal
ms.openlocfilehash: e75ab26f820e53ca0379467fba559ebcb727ee3a
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132145"
---
# <a name="authentication-patterns"></a><span data-ttu-id="f4bf7-103">認証パターン</span><span class="sxs-lookup"><span data-stu-id="f4bf7-103">Authentication patterns</span></span>

<span data-ttu-id="f4bf7-104">アドインの機能にユーザーがアクセスするには、サインインまたはサインアップする必要があります。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-104">Add-ins may require users to sign-in or sign-up in order to access features and functionality.</span></span> <span data-ttu-id="f4bf7-105">認証時の典型的なインターフェイス コントロールには、ユーザー名とパスワードの入力ボックスやサードパーティの資格情報フローを開始するボタンがあります。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-105">Input boxes for username and password or buttons that start third party credential flows are common interface controls in authentication experiences.</span></span> <span data-ttu-id="f4bf7-106">ユーザーにアドインの使用を開始してもらうには、簡単で効率的に認証を導入することが重要な最初の一歩となります。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-106">A simple and efficient authentication experience is an important first step to getting users started with your add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="f4bf7-107">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="f4bf7-107">Best practices</span></span>

|<span data-ttu-id="f4bf7-108">するべきこと</span><span class="sxs-lookup"><span data-stu-id="f4bf7-108">Do</span></span>|<span data-ttu-id="f4bf7-109">してはいけないこと</span><span class="sxs-lookup"><span data-stu-id="f4bf7-109">Don't</span></span>|
|:----|:----|
|<span data-ttu-id="f4bf7-110">サインインの前に、アドインの価値について説明し、アカウントを要求せずにこの機能を実際に使用します。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-110">Prior to sign-in, describe the value of your add-in or demonstrate functionality without requiring an account.</span></span> |<span data-ttu-id="f4bf7-111">アドインの価値と長所を理解せずにサインインすることをユーザーに期待します。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-111">Expect users to sign-in without understanding the value and benefits of your add-in.</span></span>|
|<span data-ttu-id="f4bf7-112">各画面に目立つ第 1 のボタンを配置し、ユーザーに認証フローを段階的に説明します。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-112">Guide users through authentication flows with a primary, highly visible button on each screen.</span></span> |<span data-ttu-id="f4bf7-113">ボタンや行動喚起が競合する第 2 のタスクや第 3 のタスクに注意を向けさせます。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-113">Draw attention to secondary and tertiary tasks with competing buttons and calls to action.</span></span>|
|<span data-ttu-id="f4bf7-114">"サインイン" や "アカウント作成" など、特定のタスクを説明するわかりやすいボタン ラベルを使用します。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-114">Use clear button labels that describe specific tasks like "Sign in" or "Create account".</span></span> |<span data-ttu-id="f4bf7-115">認証フローでユーザーを誘導するとき、"送信" や "開始" のようなあいまいなボタン ラベルを使用します。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-115">Use vague button labels like "Submit" or "Get started" to guide users through authentication flows.</span></span>|
|<span data-ttu-id="f4bf7-116">ダイアログを使用し、ユーザーの注意を認証フォームに向けさせます。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-116">Use a dialog to focus users' attention on authentication forms.</span></span> |<span data-ttu-id="f4bf7-117">最初の実行エクスペリエンスと認証フォームで作業ウィンドウをあふれさせます。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-117">Overcrowd your task pane with a first run experience and authentication forms.</span></span>|
|<span data-ttu-id="f4bf7-118">入力ボックスの自動フォーカスなど、フローの中に小さな効率性を見つけます。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-118">Find small efficiencies in the flow like auto-focusing on input boxes.</span></span> |<span data-ttu-id="f4bf7-119">クリックしてフォーム フィールドに入るようにユーザーに要求するなど、操作に不要な手順を追加します。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-119">Add unnecessary steps to the interaction like requiring users to click into form fields.</span></span>|
|<span data-ttu-id="f4bf7-120">ユーザーがサインアウトして再認証する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-120">Provide a way for users to sign out and reauthenticate.</span></span> |<span data-ttu-id="f4bf7-121">ID を切り替える際、アンインストールをユーザーに強制します。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-121">Force users to uninstall to switch identities.</span></span>|

## <a name="authentication-flow"></a><span data-ttu-id="f4bf7-122">認証フロー</span><span class="sxs-lookup"><span data-stu-id="f4bf7-122">Authentication flow</span></span>

1. <span data-ttu-id="f4bf7-123">初回実行プレースマット - アドインの最初の実行エクスペリエンス内にわかりやすい行動喚起としてサインイン ボタンを配置します。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-123">First Run Placemat - Place your sign-in button as a clear call-to action inside your add-in's first run experience.</span></span>

    ![Office アプリケーションのアドイン作業ウィンドウを示すスクリーンショット](../images/add-in-fre-value-placemat.png)

1. <span data-ttu-id="f4bf7-125">ID プロバイダーの選択肢ダイアログ - ID プロバイダーのわかりやすい一覧を表示します。該当する場合、ユーザー名やパスワードのフォームも含めます。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-125">Identity Provider Choices Dialog - Display a clear list of identity providers including a username and password form if applicable.</span></span> <span data-ttu-id="f4bf7-126">認証ダイアログが開いているとき、アドイン UI はブロックされることがあります。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-126">Your add-in UI may be blocked while the authentication dialog is open.</span></span>

    ![Office アプリケーションの Id プロバイダーの選択肢ダイアログボックスを示すスクリーンショット](../images/add-in-auth-choices-dialog.png)

1. <span data-ttu-id="f4bf7-128">ID プロバイダーのサインイン - ID プロバイダーによって独自の UI が提供されます。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-128">Identity Provider Sign-in - The identity provider will have their own UI.</span></span> <span data-ttu-id="f4bf7-129">Microsoft Azure Active Directory を使用すると、サインインページとアクセスパネルページをカスタマイズして、サービスに一貫性のあるルックアンドフィールを持たせることができます。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-129">Microsoft Azure Active Directory allows customization of sign-in and access panel pages for consistent look and feel with your service.</span></span> <span data-ttu-id="f4bf7-130">[詳細を参照](/azure/active-directory/fundamentals/customize-branding)してください。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-130">[Learn More](/azure/active-directory/fundamentals/customize-branding).</span></span>

    ![Office アプリケーションの Id プロバイダーのサインインダイアログを示すスクリーンショット](../images/add-in-auth-identity-sign-in.png)

1. <span data-ttu-id="f4bf7-132">進捗状況 - 設定や UI の読み込みの進行状況を示します。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-132">Progress - Indicate progress while settings and UI load.</span></span>

    ![Office アプリケーションで進行状況インジケーターが表示されているダイアログを示すスクリーンショット](../images/add-in-auth-modal-interstitial.png)

> [!NOTE]
> <span data-ttu-id="f4bf7-134">Microsoft の ID サービスを使用すると、商標付きのサインイン ボタンを使用できます。このボタンは淡色テーマまたは濃色テーマにカスタマイズできます。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-134">When using Microsoft's Identity service you'll have the opportunity to use a branded sign-in button that is customizable to light and dark themes.</span></span> <span data-ttu-id="f4bf7-135">詳細情報。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-135">Learn more.</span></span>

## <a name="single-sign-on-authentication-flow"></a><span data-ttu-id="f4bf7-136">単一の Sign-On 認証フロー</span><span class="sxs-lookup"><span data-stu-id="f4bf7-136">Single Sign-On authentication flow</span></span>

> [!NOTE]
> <span data-ttu-id="f4bf7-137">現在、シングルサインオン API は Word、Excel、Outlook、および PowerPoint でサポートされています。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-137">The single sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="f4bf7-138">シングルサインオンのサポートの詳細については、「identity [api の要件セット](../reference/requirement-sets/identity-api-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-138">For more information about single sign-on support, see [IdentityAPI requirement sets](../reference/requirement-sets/identity-api-requirement-sets.md).</span></span> <span data-ttu-id="f4bf7-139">Outlook アドインで作業している場合は、Office 365 テナントの先進認証が有効になっていることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-139">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="f4bf7-140">この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-140">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="f4bf7-141">エンドユーザーの利便性を向上させるためにシングルサインオンを使用します。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-141">Use single sign-on for a smoother end-user experience.</span></span> <span data-ttu-id="f4bf7-142">Office 内のユーザーの id (Microsoft アカウントまたは Microsoft 365 identity) は、アドインにサインインするために使用されます。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-142">The user's identity within Office (either a Microsoft Account or a Microsoft 365 identity) is used to sign in to your add-in.</span></span> <span data-ttu-id="f4bf7-143">その結果、ユーザーは一度だけサインインします。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-143">As a result users only sign in once.</span></span> <span data-ttu-id="f4bf7-144">お客様は途中で止められることなく、簡単に利用を開始できます。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-144">This removes friction in the experience making it easier for your customers to get started.</span></span>

1. <span data-ttu-id="f4bf7-145">アドインがインストールされると、ユーザーには次のような同意ウィンドウが表示されます。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-145">As an add-in is being installed, a user will see a consent window similar to the one following:</span></span>

    ![アドインのインストール時に Office アプリケーションの同意ウィンドウを示すスクリーンショット](../images/add-in-auth-SSO-consent-dialog.png)

    > [!NOTE]
    > <span data-ttu-id="f4bf7-147">この同意ウィンドウに含まれるロゴ、文字列、アクセス許可の範囲については、アドインの発行元が制御します。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-147">The add-in publisher will have control over the logo, strings and permission scopes included in the consent window.</span></span> <span data-ttu-id="f4bf7-148">UI は Microsoft が事前に構成したものです。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-148">The UI is pre-configured by Microsoft.</span></span>

1. <span data-ttu-id="f4bf7-149">アドインはユーザーが同意した後で読み込まれます。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-149">The add-in will load after the user consents.</span></span> <span data-ttu-id="f4bf7-150">ユーザーがカスタマイズした情報が必要であれば、それを抽出し、表示できます。</span><span class="sxs-lookup"><span data-stu-id="f4bf7-150">It can extract and display any necessary user customized information.</span></span>

    ![リボンにアドインボタンが表示されている Office アプリケーションを示すスクリーンショット](../images/add-in-ribbon.png)

## <a name="see-also"></a><span data-ttu-id="f4bf7-152">関連項目</span><span class="sxs-lookup"><span data-stu-id="f4bf7-152">See also</span></span>

- <span data-ttu-id="f4bf7-153">[SSO アドインの開発](../develop/sso-in-office-add-ins.md)に関する詳細情報</span><span class="sxs-lookup"><span data-stu-id="f4bf7-153">Learn more about [developing SSO Add-ins](../develop/sso-in-office-add-ins.md)</span></span>
