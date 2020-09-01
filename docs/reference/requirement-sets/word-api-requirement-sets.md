---
title: Word JavaScript API の要件セット
description: Word ビルド用の Office アドイン要件セットの情報。
ms.date: 07/10/2020
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: 23db807df0c47aaab4c579d17e4fbd28bb809fed
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293437"
---
# <a name="word-javascript-api-requirement-sets"></a><span data-ttu-id="79094-103">Word JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="79094-103">Word JavaScript API requirement sets</span></span>

<span data-ttu-id="79094-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="79094-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="requirement-set-availability"></a><span data-ttu-id="79094-107">要件セットの可用性</span><span class="sxs-lookup"><span data-stu-id="79094-107">Requirement set availability</span></span>

<span data-ttu-id="79094-108">Word アドインは、Windows での Office 2016 以降、Office on the web、iPad、および Mac など、複数のバージョンの Office で機能します。</span><span class="sxs-lookup"><span data-stu-id="79094-108">Word add-ins run across multiple versions of Office, including Office 2016 or later on Windows, and Office on the web, iPad, and Mac.</span></span> <span data-ttu-id="79094-109">次の表は、Word の要件セット、その要件セットをサポートする Office クライアント アプリケーション、およびそれらのアプリケーションのビルド番号またはバージョン番号の一覧です。</span><span class="sxs-lookup"><span data-stu-id="79094-109">The following table lists the Word requirement sets, the Office client applications that support that requirement set, and the build or version numbers for those applications.</span></span>

> [!NOTE]
> <span data-ttu-id="79094-110">番号付きの要件セットで API を使用するには、CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js で**実稼働**ライブラリを参照してください。</span><span class="sxs-lookup"><span data-stu-id="79094-110">To use APIs in any of the numbered requirement sets, you should reference the **production** library on the CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.</span></span>
>
> <span data-ttu-id="79094-111">プレビューの API の使用に関する詳細については、記事「[Excel JavaScript プレビュー API](word-preview-apis.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="79094-111">For information about using preview APIs, see the [Excel JavaScript preview APIs](word-preview-apis.md) article.</span></span>

|  <span data-ttu-id="79094-112">要件セット</span><span class="sxs-lookup"><span data-stu-id="79094-112">Requirement set</span></span>  |   <span data-ttu-id="79094-113">Windows での Office\*</span><span class="sxs-lookup"><span data-stu-id="79094-113">Office on Windows\*</span></span><br><span data-ttu-id="79094-114">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="79094-114">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="79094-115">Office on iPad</span><span class="sxs-lookup"><span data-stu-id="79094-115">Office on iPad</span></span><br><span data-ttu-id="79094-116">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="79094-116">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="79094-117">Office on Mac</span><span class="sxs-lookup"><span data-stu-id="79094-117">Office on Mac</span></span><br><span data-ttu-id="79094-118">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="79094-118">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="79094-119">Office on the web</span><span class="sxs-lookup"><span data-stu-id="79094-119">Office on the web</span></span>  |
|:-----|-----|:-----|:-----|:-----|
| [<span data-ttu-id="79094-120">プレビュー</span><span class="sxs-lookup"><span data-stu-id="79094-120">Preview</span></span>](word-preview-apis.md) | <span data-ttu-id="79094-121">プレビュー API を試すには、最新版 Office を使用してください (場合によっては、[Office Insider プログラム](https://insider.office.com)に参加する必要があります)</span><span class="sxs-lookup"><span data-stu-id="79094-121">Please use the latest Office version to try preview APIs (you may need to join the [Office Insider program](https://insider.office.com))</span></span> |
| [<span data-ttu-id="79094-122">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="79094-122">WordApi 1.3</span></span>](word-api-1-3-requirement-set.md) | <span data-ttu-id="79094-123">バージョン 1612 (ビルド 7668.1000) 以降</span><span class="sxs-lookup"><span data-stu-id="79094-123">Version 1612 (Build 7668.1000) or later</span></span>| <span data-ttu-id="79094-124">2017 年 3 月、2.22 以降</span><span class="sxs-lookup"><span data-stu-id="79094-124">March 2017, 2.22 or later</span></span> | <span data-ttu-id="79094-125">2017 年 3 月、15.32 以降</span><span class="sxs-lookup"><span data-stu-id="79094-125">March 2017, 15.32 or later</span></span>| <span data-ttu-id="79094-126">2017 年 3 月</span><span class="sxs-lookup"><span data-stu-id="79094-126">March 2017</span></span> |
| [<span data-ttu-id="79094-127">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="79094-127">WordApi 1.2</span></span>](word-api-1-2-requirement-set.md) | <span data-ttu-id="79094-128">2015年 12 月更新プログラム、バージョン 1601 (ビルド 6568.1000) 以降</span><span class="sxs-lookup"><span data-stu-id="79094-128">December 2015 update, Version 1601 (Build 6568.1000) or later</span></span> | <span data-ttu-id="79094-129">2016 年 1 月、1.18 以降</span><span class="sxs-lookup"><span data-stu-id="79094-129">January 2016, 1.18 or later</span></span> | <span data-ttu-id="79094-130">2016 年 1 月、15.19 以降</span><span class="sxs-lookup"><span data-stu-id="79094-130">January 2016, 15.19 or later</span></span>| <span data-ttu-id="79094-131">2016 年 9 月</span><span class="sxs-lookup"><span data-stu-id="79094-131">September 2016</span></span> |
| [<span data-ttu-id="79094-132">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="79094-132">WordApi 1.1</span></span>](word-api-1-1-requirement-set.md) | <span data-ttu-id="79094-133">バージョン 1509 (ビルド 4266.1001) 以降</span><span class="sxs-lookup"><span data-stu-id="79094-133">Version 1509 (Build 4266.1001) or later</span></span>| <span data-ttu-id="79094-134">2016 年 1 月、1.18 以降</span><span class="sxs-lookup"><span data-stu-id="79094-134">January 2016, 1.18 or later</span></span> | <span data-ttu-id="79094-135">2016 年 1 月、15.19 以降</span><span class="sxs-lookup"><span data-stu-id="79094-135">January 2016, 15.19 or later</span></span>| <span data-ttu-id="79094-136">2016 年 9 月</span><span class="sxs-lookup"><span data-stu-id="79094-136">September 2016</span></span> |

> [!NOTE]
> <span data-ttu-id="79094-137">永続ライセンス版 Office でサポートされる要件セットは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="79094-137">Perpetual versions of Office support requirement sets as follows:</span></span>
>
> - <span data-ttu-id="79094-138">Office 2019 では WordApi 1.3 以前がサポートされています。</span><span class="sxs-lookup"><span data-stu-id="79094-138">Office 2019 supports WordApi 1.3 and earlier.</span></span>
> - <span data-ttu-id="79094-139">Office 2016 では WordApi 1.1 要求セットのみがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="79094-139">Office 2016 only supports the WordApi 1.1 requirement set.</span></span>

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="79094-140">Office のバージョンとビルド番号</span><span class="sxs-lookup"><span data-stu-id="79094-140">Office versions and build numbers</span></span>

<span data-ttu-id="79094-141">Office のバージョンとビルド番号の詳細については、次を参照してください。</span><span class="sxs-lookup"><span data-stu-id="79094-141">For more information about Office versions and build numbers, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="see-also"></a><span data-ttu-id="79094-142">関連項目</span><span class="sxs-lookup"><span data-stu-id="79094-142">See also</span></span>

- [<span data-ttu-id="79094-143">Word JavaScript API リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="79094-143">Word JavaScript API Reference Documentation</span></span>](/javascript/api/word)
- [<span data-ttu-id="79094-144">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="79094-144">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="79094-145">Office アプリケーションと API 要件を指定する</span><span class="sxs-lookup"><span data-stu-id="79094-145">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="79094-146">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="79094-146">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
