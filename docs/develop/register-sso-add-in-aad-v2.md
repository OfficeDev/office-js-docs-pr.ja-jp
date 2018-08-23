---
title: Azure AD v2.0 のエンドポイントに SSO を使用する Office アドインを登録する
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 95b690e21bddf7f2754cc308c8b771e629bbc630
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437256"
---
# <a name="register-an-office-add-in-that-uses-sso-with-the-azure-ad-v20-endpoint"></a><span data-ttu-id="c4858-102">Azure AD v2.0 のエンドポイントに SSO を使用する Office アドインを登録する</span><span class="sxs-lookup"><span data-stu-id="c4858-102">Details are at: Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint.</span></span>

<span data-ttu-id="c4858-103">この記事では、Azure AD v2.0 のエンドポイントに Office アドインを登録する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="c4858-103">This article explains how to register an Office Add-in with the Azure AD v2.0 endpoint.</span></span> <span data-ttu-id="c4858-104">アドインの開発を開始するときには、アドインを登録する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c4858-104">You need to register the add-in when you begin developing it.</span></span> <span data-ttu-id="c4858-105">テストまたは運用に進むと、アドインの開発、テスト、および運用バージョン用に、既存の登録を変更するか、別々の登録を作成するかできます。</span><span class="sxs-lookup"><span data-stu-id="c4858-105">When you progress to testing or production, you can change the existing registration or create separate registrations for development, testing, and production versions of the add-in.</span></span> 

<span data-ttu-id="c4858-106">次の表には、このプロシージャを実行するために必要な情報と、指示に表示される対応するプレースホルダーを挙げてあります。</span><span class="sxs-lookup"><span data-stu-id="c4858-106">The following table itemizes the information that you need to carry out this procedure and the corresponding placeholders that appear in the instructions.</span></span> 

|<span data-ttu-id="c4858-107">情報</span><span class="sxs-lookup"><span data-stu-id="c4858-107">Information</span></span>  |<span data-ttu-id="c4858-108">例</span><span class="sxs-lookup"><span data-stu-id="c4858-108">Examples</span></span>  |<span data-ttu-id="c4858-109">プレースホルダー</span><span class="sxs-lookup"><span data-stu-id="c4858-109">Placeholder</span></span>  |
|---------|---------|---------|
|<span data-ttu-id="c4858-110">アドインの読みやすい名前。</span><span class="sxs-lookup"><span data-stu-id="c4858-110">A human readable name for the add-in.</span></span> <span data-ttu-id="c4858-111">(一意であることが推奨されます。ただし、一意でなくてもかまいません。)</span><span class="sxs-lookup"><span data-stu-id="c4858-111">(Uniqueness recommended, but not required.)</span></span>    |`Contoso Marketing Excel Add-in (Prod)`        |<span data-ttu-id="c4858-112">**$ADD-IN-NAME$**</span><span class="sxs-lookup"><span data-stu-id="c4858-112">**$ADD-IN-NAME$**</span></span>         |
|<span data-ttu-id="c4858-113">アドインの完全修飾ドメイン名（プロトコルを除く）。</span><span class="sxs-lookup"><span data-stu-id="c4858-113">The fully qualified domain name (except for protocol) of the add-in.</span></span> <span data-ttu-id="c4858-114">*自分が所有するドメインを使用する必要があります。*</span><span class="sxs-lookup"><span data-stu-id="c4858-114">*You must use a domain that you own.*</span></span> <span data-ttu-id="c4858-115">このため、 `azurewebsites.net` や `cloudapp.net` のような特定のよく知られたドメインは使用できません。</span><span class="sxs-lookup"><span data-stu-id="c4858-115">For this reason, you cannot use certain well-known domains such as `azurewebsites.net` or `cloudapp.net`.</span></span>   |<span data-ttu-id="c4858-116">`localhost:6789`, `addins.contoso.com`</span><span class="sxs-lookup"><span data-stu-id="c4858-116"></span></span>         |<span data-ttu-id="c4858-117">**$FQDN-WITHOUT-PROTOCOL$**</span><span class="sxs-lookup"><span data-stu-id="c4858-117">**$FQDN-WITHOUT-PROTOCOL$**</span></span>         |
|<span data-ttu-id="c4858-118">アドインに必要な AAD と Microsoft Graph へのアクセス許可。</span><span class="sxs-lookup"><span data-stu-id="c4858-118">The permissions to AAD and Microsoft Graph that your add-in needs.</span></span> <span data-ttu-id="c4858-119">（`profile` が常に必要です。）</span><span class="sxs-lookup"><span data-stu-id="c4858-119">(`profile` is always required.)</span></span>    |<span data-ttu-id="c4858-120">`profile`, `Files.Read.All`</span><span class="sxs-lookup"><span data-stu-id="c4858-120"></span></span>         |<span data-ttu-id="c4858-121">該当なし</span><span class="sxs-lookup"><span data-stu-id="c4858-121">N/A</span></span>         |

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]