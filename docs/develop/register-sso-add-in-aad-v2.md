---
title: Azure AD v2.0 のエンドポイントに SSO を使用する Office アドインを登録する
description: Azure AD v2.0 エンドポイントを使用して Office アドインを登録する方法について説明します。
ms.date: 04/10/2019
localization_priority: Normal
ms.openlocfilehash: 8bcd72bd6f2d56c5f97d2d4f153d6791d111452e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609377"
---
# <a name="register-an-office-add-in-that-uses-sso-with-the-azure-ad-v20-endpoint"></a><span data-ttu-id="4c0b2-103">Azure AD v2.0 のエンドポイントに SSO を使用する Office アドインを登録する</span><span class="sxs-lookup"><span data-stu-id="4c0b2-103">Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint</span></span>

<span data-ttu-id="4c0b2-104">この記事では、Azure AD v2.0 のエンドポイントに Office アドインを登録する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="4c0b2-104">This article explains how to register an Office Add-in with the Azure AD v2.0 endpoint.</span></span> <span data-ttu-id="4c0b2-105">開発を開始する前に、アドインを登録する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4c0b2-105">You need to register the add-in when you begin developing it.</span></span> <span data-ttu-id="4c0b2-106">テストまたは運用環境に進んだ場合、既存の登録を変更するか、アドインの開発、テスト、および運用バージョン用に別の登録を作成できます。</span><span class="sxs-lookup"><span data-stu-id="4c0b2-106">When you progress to testing or production, you can change the existing registration or create separate registrations for development, testing, and production versions of the add-in.</span></span>

<span data-ttu-id="4c0b2-107">次の表では、この手順を実行するために必要な情報と、指示に表示される対応するプレースホルダーが項目ごとに分類されています。</span><span class="sxs-lookup"><span data-stu-id="4c0b2-107">The following table itemizes the information that you need to carry out this procedure and the corresponding placeholders that appear in the instructions.</span></span>

|<span data-ttu-id="4c0b2-108">情報</span><span class="sxs-lookup"><span data-stu-id="4c0b2-108">Information</span></span>  |<span data-ttu-id="4c0b2-109">例</span><span class="sxs-lookup"><span data-stu-id="4c0b2-109">Examples</span></span>  |<span data-ttu-id="4c0b2-110">プレースホルダー</span><span class="sxs-lookup"><span data-stu-id="4c0b2-110">Placeholder</span></span>  |
|---------|---------|---------|
|<span data-ttu-id="4c0b2-111">人間が判読できるアドインの名前です </span><span class="sxs-lookup"><span data-stu-id="4c0b2-111">A human readable name for the add-in.</span></span> <span data-ttu-id="4c0b2-112">(一意であることが推奨されますが、必須ではありません)。</span><span class="sxs-lookup"><span data-stu-id="4c0b2-112">(Uniqueness recommended, but not required.)</span></span>|`Contoso Marketing Excel Add-in (Prod)`|<span data-ttu-id="4c0b2-113">**$ADD-IN-NAME$**</span><span class="sxs-lookup"><span data-stu-id="4c0b2-113">**$ADD-IN-NAME$**</span></span>|
|<span data-ttu-id="4c0b2-114">アドインの完全修飾ドメイン名 (プロトコルを除く) です。</span><span class="sxs-lookup"><span data-stu-id="4c0b2-114">The fully qualified domain name (except for protocol) of the add-in.</span></span> <span data-ttu-id="4c0b2-115">*所有しているドメインを使用する必要があります。*</span><span class="sxs-lookup"><span data-stu-id="4c0b2-115">*You must use a domain that you own.*</span></span> <span data-ttu-id="4c0b2-116">この理由から、`azurewebsites.net` または `cloudapp.net` などのよく知られている特定のドメインは使用できません。</span><span class="sxs-lookup"><span data-stu-id="4c0b2-116">For this reason, you cannot use certain well-known domains such as `azurewebsites.net` or `cloudapp.net`.</span></span> <span data-ttu-id="4c0b2-117">このドメインは、アドインのマニフェストの `<Resources>` のセクションにある URL で使用されている、すべてのサブドメインを含むドメインと一致している必要があります。</span><span class="sxs-lookup"><span data-stu-id="4c0b2-117">The domain must be the same, including any subdomains, as is used in the URLs in the `<Resources>` section of the add-in's manifest.</span></span>|<span data-ttu-id="4c0b2-118">`localhost:6789`, `addins.contoso.com`</span><span class="sxs-lookup"><span data-stu-id="4c0b2-118">`localhost:6789`, `addins.contoso.com`</span></span>|<span data-ttu-id="4c0b2-119">**$FQDN-WITHOUT-PROTOCOL$**</span><span class="sxs-lookup"><span data-stu-id="4c0b2-119">**$FQDN-WITHOUT-PROTOCOL$**</span></span>|
|<span data-ttu-id="4c0b2-120">ご使用のアドインに必要な AAD および Microsoft Graph へのアクセス許可です </span><span class="sxs-lookup"><span data-stu-id="4c0b2-120">The permissions to AAD and Microsoft Graph that your add-in needs.</span></span> <span data-ttu-id="4c0b2-121">(`profile` は常に必須です)。</span><span class="sxs-lookup"><span data-stu-id="4c0b2-121">(`profile` is always required.)</span></span>|<span data-ttu-id="4c0b2-122">`profile`, `Files.Read.All`</span><span class="sxs-lookup"><span data-stu-id="4c0b2-122">`profile`, `Files.Read.All`</span></span>|<span data-ttu-id="4c0b2-123">N/A</span><span class="sxs-lookup"><span data-stu-id="4c0b2-123">N/A</span></span>|

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]
