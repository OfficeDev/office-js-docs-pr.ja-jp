---
title: Azure AD v2.0 のエンドポイントに SSO を使用する Office アドインを登録する
description: ''
ms.date: 04/10/2019
localization_priority: Priority
ms.openlocfilehash: a98fb7e9f073024804f577057fde83d1bdc83273
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914250"
---
# <a name="register-an-office-add-in-that-uses-sso-with-the-azure-ad-v20-endpoint"></a><span data-ttu-id="ebdc6-102">Azure AD v2.0 のエンドポイントに SSO を使用する Office アドインを登録する</span><span class="sxs-lookup"><span data-stu-id="ebdc6-102">Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint</span></span>

<span data-ttu-id="ebdc6-103">この記事では、Azure AD v2.0 のエンドポイントに Office アドインを登録する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="ebdc6-103">This article explains how to register an Office Add-in with the Azure AD v2.0 endpoint.</span></span> <span data-ttu-id="ebdc6-104">開発を開始する前に、アドインを登録する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ebdc6-104">You need to register the add-in when you begin developing it.</span></span> <span data-ttu-id="ebdc6-105">テストまたは運用環境に進んだ場合、既存の登録を変更するか、アドインの開発、テスト、および運用バージョン用に別の登録を作成できます。</span><span class="sxs-lookup"><span data-stu-id="ebdc6-105">When you progress to testing or production, you can change the existing registration or create separate registrations for development, testing, and production versions of the add-in.</span></span>

<span data-ttu-id="ebdc6-106">次の表では、この手順を実行するために必要な情報と、指示に表示される対応するプレースホルダーが項目ごとに分類されています。</span><span class="sxs-lookup"><span data-stu-id="ebdc6-106">The following table itemizes the information that you need to carry out this procedure and the corresponding placeholders that appear in the instructions.</span></span>

|<span data-ttu-id="ebdc6-107">情報</span><span class="sxs-lookup"><span data-stu-id="ebdc6-107">Information</span></span>  |<span data-ttu-id="ebdc6-108">例</span><span class="sxs-lookup"><span data-stu-id="ebdc6-108">Examples</span></span>  |<span data-ttu-id="ebdc6-109">プレースホルダー</span><span class="sxs-lookup"><span data-stu-id="ebdc6-109">Placeholder</span></span>  |
|---------|---------|---------|
|<span data-ttu-id="ebdc6-110">人間が判読できるアドインの名前です </span><span class="sxs-lookup"><span data-stu-id="ebdc6-110">A human readable name for the add-in.</span></span> <span data-ttu-id="ebdc6-111">(一意であることが推奨されますが、必須ではありません)。</span><span class="sxs-lookup"><span data-stu-id="ebdc6-111">(Uniqueness recommended, but not required.)</span></span>|`Contoso Marketing Excel Add-in (Prod)`|<span data-ttu-id="ebdc6-112">**$ADD-IN-NAME$**</span><span class="sxs-lookup"><span data-stu-id="ebdc6-112">**$ADD-IN-NAME$**</span></span>|
|<span data-ttu-id="ebdc6-113">アドインの完全修飾ドメイン名 (プロトコルを除く) です。</span><span class="sxs-lookup"><span data-stu-id="ebdc6-113">The fully qualified domain name (except for protocol) of the add-in.</span></span> <span data-ttu-id="ebdc6-114">*所有しているドメインを使用する必要があります。*</span><span class="sxs-lookup"><span data-stu-id="ebdc6-114">*You must use a domain that you own.*</span></span> <span data-ttu-id="ebdc6-115">この理由から、`azurewebsites.net` または `cloudapp.net` などのよく知られている特定のドメインは使用できません。</span><span class="sxs-lookup"><span data-stu-id="ebdc6-115">For this reason, you cannot use certain well-known domains such as `azurewebsites.net` or `cloudapp.net`.</span></span> <span data-ttu-id="ebdc6-116">このドメインは、アドインのマニフェストの `<Resources>` のセクションにある URL で使用されている、すべてのサブドメインを含むドメインと一致している必要があります。</span><span class="sxs-lookup"><span data-stu-id="ebdc6-116">The domain must be the same, including any subdomains, as is used in the URLs in the `<Resources>` section of the add-in's manifest.</span></span>|<span data-ttu-id="ebdc6-117">`localhost:6789`, `addins.contoso.com`</span><span class="sxs-lookup"><span data-stu-id="ebdc6-117"></span></span>|<span data-ttu-id="ebdc6-118">**$FQDN-WITHOUT-PROTOCOL$**</span><span class="sxs-lookup"><span data-stu-id="ebdc6-118">**$FQDN-WITHOUT-PROTOCOL$**</span></span>|
|<span data-ttu-id="ebdc6-119">ご使用のアドインに必要な AAD および Microsoft Graph へのアクセス許可です </span><span class="sxs-lookup"><span data-stu-id="ebdc6-119">The permissions to AAD and Microsoft Graph that your add-in needs.</span></span> <span data-ttu-id="ebdc6-120">(`profile` は常に必須です)。</span><span class="sxs-lookup"><span data-stu-id="ebdc6-120">(`profile` is always required.)</span></span>|<span data-ttu-id="ebdc6-121">`profile`, `Files.Read.All`</span><span class="sxs-lookup"><span data-stu-id="ebdc6-121"></span></span>|<span data-ttu-id="ebdc6-122">N/A</span><span class="sxs-lookup"><span data-stu-id="ebdc6-122">N/A</span></span>|

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]
