---
title: Word JavaScript API の要件セット
description: Word ビルド用の Office アドイン要件セットの情報
ms.date: 01/06/2020
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: c90daafe46d301b404ee902b38bb7417562adc44
ms.sourcegitcommit: abe8188684b55710261c69e206de83d3a6bd2ed3
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/08/2020
ms.locfileid: "40969532"
---
# <a name="word-javascript-api-requirement-sets"></a><span data-ttu-id="f9182-103">Word JavaScript API の要件セット</span><span class="sxs-lookup"><span data-stu-id="f9182-103">Word JavaScript API requirement sets</span></span>

<span data-ttu-id="f9182-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="f9182-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

## <a name="requirement-set-availability"></a><span data-ttu-id="f9182-107">要件セットの可用性</span><span class="sxs-lookup"><span data-stu-id="f9182-107">Requirement set availability</span></span>

<span data-ttu-id="f9182-108">Word アドインは、Windows での Office 2016 以降、Office on the web、iPad、および Mac など、複数のバージョンの Office で機能します。</span><span class="sxs-lookup"><span data-stu-id="f9182-108">Word add-ins run across multiple versions of Office, including Office 2016 or later on Windows, and Office on the web, iPad, and Mac.</span></span> <span data-ttu-id="f9182-109">次の表は、Word の要件セット、その要件セットをサポートする Office ホスト アプリケーション、およびそれらのアプリケーションのビルド番号またはバージョン番号の一覧です。</span><span class="sxs-lookup"><span data-stu-id="f9182-109">The following table lists the Word requirement sets, the Office host applications that support that requirement set, and the build or version numbers for those applications.</span></span>

> [!NOTE]
> <span data-ttu-id="f9182-110">番号付きの要件セットで API を使用するには、CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js で**実稼働**ライブラリを参照してください。</span><span class="sxs-lookup"><span data-stu-id="f9182-110">To use APIs in any of the numbered requirement sets, you should reference the **production** library on the CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.</span></span>
>
> <span data-ttu-id="f9182-111">プレビューの API の使用に関する詳細については、記事「[Excel JavaScript プレビュー API](word-preview-apis.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f9182-111">For information about using preview APIs, see the [Excel JavaScript preview APIs](word-preview-apis.md) article.</span></span>

|  <span data-ttu-id="f9182-112">要件セット</span><span class="sxs-lookup"><span data-stu-id="f9182-112">Requirement set</span></span>  |   <span data-ttu-id="f9182-113">Windows での Office\*</span><span class="sxs-lookup"><span data-stu-id="f9182-113">Office on Windows\*</span></span><br><span data-ttu-id="f9182-114">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="f9182-114">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="f9182-115">Office on iPad</span><span class="sxs-lookup"><span data-stu-id="f9182-115">Office on iPad</span></span><br><span data-ttu-id="f9182-116">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="f9182-116">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="f9182-117">Office on Mac</span><span class="sxs-lookup"><span data-stu-id="f9182-117">Office on Mac</span></span><br><span data-ttu-id="f9182-118">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="f9182-118">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="f9182-119">Office on the web</span><span class="sxs-lookup"><span data-stu-id="f9182-119">Office on the web</span></span>  |
|:-----|-----|:-----|:-----|:-----|
| [<span data-ttu-id="f9182-120">プレビュー</span><span class="sxs-lookup"><span data-stu-id="f9182-120">Preview</span></span>](word-preview-apis.md) | <span data-ttu-id="f9182-121">プレビュー API を試すには、最新版 Office を使用してください (場合によっては、[Office Insider プログラム](https://products.office.com/office-insider)に参加する必要があります)</span><span class="sxs-lookup"><span data-stu-id="f9182-121">Please use the latest Office version to try preview APIs (you may need to join the [Office Insider program](https://products.office.com/office-insider))</span></span> |
| [<span data-ttu-id="f9182-122">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="f9182-122">WordApi 1.3</span></span>](word-api-1-3-requirement-set.md) | <span data-ttu-id="f9182-123">バージョン 1612 (ビルド 7668.1000) 以降</span><span class="sxs-lookup"><span data-stu-id="f9182-123">Version 1612 (Build 7668.1000) or later</span></span>| <span data-ttu-id="f9182-124">2017 年 3 月、2.22 以降</span><span class="sxs-lookup"><span data-stu-id="f9182-124">March 2017, 2.22 or later</span></span> | <span data-ttu-id="f9182-125">2017 年 3 月、15.32 以降</span><span class="sxs-lookup"><span data-stu-id="f9182-125">March 2017, 15.32 or later</span></span>| <span data-ttu-id="f9182-126">2017 年 3 月</span><span class="sxs-lookup"><span data-stu-id="f9182-126">March 2017</span></span> |
| [<span data-ttu-id="f9182-127">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="f9182-127">WordApi 1.2</span></span>](word-api-1-2-requirement-set.md) | <span data-ttu-id="f9182-128">2015年 12 月更新プログラム、バージョン 1601 (ビルド 6568.1000) 以降</span><span class="sxs-lookup"><span data-stu-id="f9182-128">December 2015 update, Version 1601 (Build 6568.1000) or later</span></span> | <span data-ttu-id="f9182-129">2016 年 1 月、1.18 以降</span><span class="sxs-lookup"><span data-stu-id="f9182-129">January 2016, 1.18 or later</span></span> | <span data-ttu-id="f9182-130">2016 年 1 月、15.19 以降</span><span class="sxs-lookup"><span data-stu-id="f9182-130">January 2016, 15.19 or later</span></span>| <span data-ttu-id="f9182-131">2016 年 9 月</span><span class="sxs-lookup"><span data-stu-id="f9182-131">September 2016</span></span> |
| [<span data-ttu-id="f9182-132">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f9182-132">WordApi 1.1</span></span>](word-api-1-1-requirement-set.md) | <span data-ttu-id="f9182-133">バージョン 1509 (ビルド 4266.1001) 以降</span><span class="sxs-lookup"><span data-stu-id="f9182-133">Version 1509 (Build 4266.1001) or later</span></span>| <span data-ttu-id="f9182-134">2016 年 1 月、1.18 以降</span><span class="sxs-lookup"><span data-stu-id="f9182-134">January 2016, 1.18 or later</span></span> | <span data-ttu-id="f9182-135">2016 年 1 月、15.19 以降</span><span class="sxs-lookup"><span data-stu-id="f9182-135">January 2016, 15.19 or later</span></span>| <span data-ttu-id="f9182-136">2016 年 9 月</span><span class="sxs-lookup"><span data-stu-id="f9182-136">September 2016</span></span> |

> [!NOTE]
> <span data-ttu-id="f9182-137">永続ライセンス版 Office でサポートされる要件セットは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="f9182-137">Perpetual versions of Office support requirement sets as follows:</span></span>
>
> - <span data-ttu-id="f9182-138">Office 2019 では WordApi 1.3 以前がサポートされています。</span><span class="sxs-lookup"><span data-stu-id="f9182-138">Office 2019 supports WordApi 1.3 and earlier.</span></span>
> - <span data-ttu-id="f9182-139">Office 2016 では WordApi 1.1 要求セットのみがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="f9182-139">Office 2016 only supports the WordApi 1.1 requirement set.</span></span>

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="f9182-140">Office のバージョンとビルド番号</span><span class="sxs-lookup"><span data-stu-id="f9182-140">Office versions and build numbers</span></span>

<span data-ttu-id="f9182-141">Office のバージョンとビルド番号の詳細については、次を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f9182-141">For more information about Office versions and build numbers, see:</span></span>

- [<span data-ttu-id="f9182-142">Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号</span><span class="sxs-lookup"><span data-stu-id="f9182-142">Version and build numbers of update channel releases for Office 365 clients</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="f9182-143">使用している Office のバージョンを確認する方法</span><span class="sxs-lookup"><span data-stu-id="f9182-143">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [<span data-ttu-id="f9182-144">Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。</span><span class="sxs-lookup"><span data-stu-id="f9182-144">Where you can find the version and build number for an Office 365 client application</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)

## <a name="see-also"></a><span data-ttu-id="f9182-145">関連項目</span><span class="sxs-lookup"><span data-stu-id="f9182-145">See also</span></span>

- [<span data-ttu-id="f9182-146">Word JavaScript API リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="f9182-146">Word JavaScript API Reference Documentation</span></span>](/javascript/api/word)
- [<span data-ttu-id="f9182-147">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="f9182-147">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="f9182-148">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="f9182-148">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="f9182-149">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="f9182-149">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
