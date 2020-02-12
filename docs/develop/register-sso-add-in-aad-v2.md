---
title: Azure AD v2.0 のエンドポイントに SSO を使用する Office アドインを登録する
description: ''
ms.date: 04/10/2019
localization_priority: Normal
ms.openlocfilehash: 3594b1e1b22f7a4341b5fd9a5b6774f3d21d8c26
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2020
ms.locfileid: "41949673"
---
# <a name="register-an-office-add-in-that-uses-sso-with-the-azure-ad-v20-endpoint"></a><span data-ttu-id="cddf0-102">Azure AD v2.0 のエンドポイントに SSO を使用する Office アドインを登録する</span><span class="sxs-lookup"><span data-stu-id="cddf0-102">Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint</span></span>

<span data-ttu-id="cddf0-103">この記事では、Azure AD v2.0 のエンドポイントに Office アドインを登録する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="cddf0-103">This article explains how to register an Office Add-in with the Azure AD v2.0 endpoint.</span></span> <span data-ttu-id="cddf0-104">開発を開始する前に、アドインを登録する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cddf0-104">You need to register the add-in when you begin developing it.</span></span> <span data-ttu-id="cddf0-105">テストまたは運用環境に進んだ場合、既存の登録を変更するか、アドインの開発、テスト、および運用バージョン用に別の登録を作成できます。</span><span class="sxs-lookup"><span data-stu-id="cddf0-105">When you progress to testing or production, you can change the existing registration or create separate registrations for development, testing, and production versions of the add-in.</span></span>

<span data-ttu-id="cddf0-106">次の表では、この手順を実行するために必要な情報と、指示に表示される対応するプレースホルダーが項目ごとに分類されています。</span><span class="sxs-lookup"><span data-stu-id="cddf0-106">The following table itemizes the information that you need to carry out this procedure and the corresponding placeholders that appear in the instructions.</span></span>

|<span data-ttu-id="cddf0-107">情報</span><span class="sxs-lookup"><span data-stu-id="cddf0-107">Information</span></span>  |<span data-ttu-id="cddf0-108">例</span><span class="sxs-lookup"><span data-stu-id="cddf0-108">Examples</span></span>  |<span data-ttu-id="cddf0-109">プレースホルダー</span><span class="sxs-lookup"><span data-stu-id="cddf0-109">Placeholder</span></span>  |
|---------|---------|---------|
|<span data-ttu-id="cddf0-110">人間が判読できるアドインの名前です </span><span class="sxs-lookup"><span data-stu-id="cddf0-110">A human readable name for the add-in.</span></span> <span data-ttu-id="cddf0-111">(一意であることが推奨されますが、必須ではありません)。</span><span class="sxs-lookup"><span data-stu-id="cddf0-111">(Uniqueness recommended, but not required.)</span></span>|`Contoso Marketing Excel Add-in (Prod)`|<span data-ttu-id="cddf0-112">**$ADD-IN-NAME$**</span><span class="sxs-lookup"><span data-stu-id="cddf0-112">**$ADD-IN-NAME$**</span></span>|
|<span data-ttu-id="cddf0-113">アドインの完全修飾ドメイン名 (プロトコルを除く) です。</span><span class="sxs-lookup"><span data-stu-id="cddf0-113">The fully qualified domain name (except for protocol) of the add-in.</span></span> <span data-ttu-id="cddf0-114">*所有しているドメインを使用する必要があります。*</span><span class="sxs-lookup"><span data-stu-id="cddf0-114">*You must use a domain that you own.*</span></span> <span data-ttu-id="cddf0-115">この理由から、`azurewebsites.net` または `cloudapp.net` などのよく知られている特定のドメインは使用できません。</span><span class="sxs-lookup"><span data-stu-id="cddf0-115">For this reason, you cannot use certain well-known domains such as `azurewebsites.net` or `cloudapp.net`.</span></span> <span data-ttu-id="cddf0-116">このドメインは、アドインのマニフェストの `<Resources>` のセクションにある URL で使用されている、すべてのサブドメインを含むドメインと一致している必要があります。</span><span class="sxs-lookup"><span data-stu-id="cddf0-116">The domain must be the same, including any subdomains, as is used in the URLs in the `<Resources>` section of the add-in's manifest.</span></span>|<span data-ttu-id="cddf0-117">`localhost:6789`, `addins.contoso.com`</span><span class="sxs-lookup"><span data-stu-id="cddf0-117">`localhost:6789`, `addins.contoso.com`</span></span>|<span data-ttu-id="cddf0-118">**$FQDN-WITHOUT-PROTOCOL$**</span><span class="sxs-lookup"><span data-stu-id="cddf0-118">**$FQDN-WITHOUT-PROTOCOL$**</span></span>|
|<span data-ttu-id="cddf0-119">ご使用のアドインに必要な AAD および Microsoft Graph へのアクセス許可です </span><span class="sxs-lookup"><span data-stu-id="cddf0-119">The permissions to AAD and Microsoft Graph that your add-in needs.</span></span> <span data-ttu-id="cddf0-120">(`profile` は常に必須です)。</span><span class="sxs-lookup"><span data-stu-id="cddf0-120">(`profile` is always required.)</span></span>|<span data-ttu-id="cddf0-121">`profile`, `Files.Read.All`</span><span class="sxs-lookup"><span data-stu-id="cddf0-121">`profile`, `Files.Read.All`</span></span>|<span data-ttu-id="cddf0-122">N/A</span><span class="sxs-lookup"><span data-stu-id="cddf0-122">N/A</span></span>|

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]
