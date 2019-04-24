---
title: ダイアログ API の要件セット
description: ''
ms.date: 03/19/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: ebbd10e65894a7d038e54ffbaac20c973adf4a9f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450087"
---
# <a name="dialog-api-requirement-sets"></a><span data-ttu-id="e6ffd-102">ダイアログ API の要件セット</span><span class="sxs-lookup"><span data-stu-id="e6ffd-102">Dialog API requirement sets</span></span>

<span data-ttu-id="e6ffd-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="e6ffd-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="e6ffd-p102">Office アドインは Office の複数のバージョンで機能します。次の表は、ダイアログ API の要件セット、その要件セットをサポートする Office ホスト アプリケーション、Office アプリケーションのビルド番号またはバージョン番号の一覧です。</span><span class="sxs-lookup"><span data-stu-id="e6ffd-p102">Office Add-ins run across multiple versions of Office. The following table lists the Dialog API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="e6ffd-108">要件セット</span><span class="sxs-lookup"><span data-stu-id="e6ffd-108">Requirement set</span></span>  | <span data-ttu-id="e6ffd-109">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="e6ffd-109">Office 2013 for Windows</span></span> | <span data-ttu-id="e6ffd-110">Outlook 2016 for Windows 以降</span><span class="sxs-lookup"><span data-stu-id="e6ffd-110">Office 2016 or later for Windows</span></span>   | <span data-ttu-id="e6ffd-111">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="e6ffd-111">Office 365 for Windows</span></span> |  <span data-ttu-id="e6ffd-112">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="e6ffd-112">Office 365 for iPad</span></span>  |  <span data-ttu-id="e6ffd-113">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="e6ffd-113">Office 365 for Mac</span></span>  | <span data-ttu-id="e6ffd-114">Office Online</span><span class="sxs-lookup"><span data-stu-id="e6ffd-114">Office Online</span></span>  |  <span data-ttu-id="e6ffd-115">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="e6ffd-115">Office Online Server</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="e6ffd-116">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="e6ffd-116">DialogApi 1.1</span></span>  | <span data-ttu-id="e6ffd-117">ビルド 15.0.4855.1000 以降</span><span class="sxs-lookup"><span data-stu-id="e6ffd-117">Build 15.0.4855.1000 or later</span></span> | <span data-ttu-id="e6ffd-118">ビルド 16.0.4390.1000 以降</span><span class="sxs-lookup"><span data-stu-id="e6ffd-118">Build 16.0.4390.1000 or later</span></span> | <span data-ttu-id="e6ffd-119">バージョン 1602 (ビルド 6741.0000) 以降</span><span class="sxs-lookup"><span data-stu-id="e6ffd-119">Version 1602 (Build 6741.0000) or later</span></span> | <span data-ttu-id="e6ffd-120">1.22 以降</span><span class="sxs-lookup"><span data-stu-id="e6ffd-120">1.22 or later</span></span> | <span data-ttu-id="e6ffd-121">15.20 以降</span><span class="sxs-lookup"><span data-stu-id="e6ffd-121">15.20 or later</span></span>| <span data-ttu-id="e6ffd-122">2017 年 1 月</span><span class="sxs-lookup"><span data-stu-id="e6ffd-122">January 2017</span></span> | <span data-ttu-id="e6ffd-123">バージョン 1608 (ビルド 7601.6800) 以降</span><span class="sxs-lookup"><span data-stu-id="e6ffd-123">Version 1608 (Build 7601.6800) or later</span></span>|

<span data-ttu-id="e6ffd-124">バージョン、ビルド番号、および Office Online Server の詳細については以下を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e6ffd-124">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

- [<span data-ttu-id="e6ffd-125">Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号</span><span class="sxs-lookup"><span data-stu-id="e6ffd-125">Version and build numbers of update channel releases for Office 365 clients</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="e6ffd-126">使用している Office のバージョンを確認する方法</span><span class="sxs-lookup"><span data-stu-id="e6ffd-126">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [<span data-ttu-id="e6ffd-127">Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。</span><span class="sxs-lookup"><span data-stu-id="e6ffd-127">Where you can find the version and build number for an Office 365 client application</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="e6ffd-128">Office Online Server 概要</span><span class="sxs-lookup"><span data-stu-id="e6ffd-128">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="e6ffd-129">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="e6ffd-129">Office Common API requirement sets</span></span>

<span data-ttu-id="e6ffd-130">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="e6ffd-130">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="dialog-api-11"></a><span data-ttu-id="e6ffd-131">ダイアログ API 1.1</span><span class="sxs-lookup"><span data-stu-id="e6ffd-131">Dialog API 1.1</span></span>

<span data-ttu-id="e6ffd-132">ダイアログ API 1.1 は、API の最初のバージョンです。</span><span class="sxs-lookup"><span data-stu-id="e6ffd-132">The Dialog API 1.1 is the first version of the API.</span></span> <span data-ttu-id="e6ffd-133">API について詳しくは、[ダイアログ API](/javascript/api/office/office.ui) リファレンスのトピックをご覧ください。</span><span class="sxs-lookup"><span data-stu-id="e6ffd-133">For details about the API, see the [Dialog API ](/javascript/api/office/office.ui) reference topic.</span></span>

## <a name="see-also"></a><span data-ttu-id="e6ffd-134">関連項目</span><span class="sxs-lookup"><span data-stu-id="e6ffd-134">See also</span></span>

- [<span data-ttu-id="e6ffd-135">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="e6ffd-135">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="e6ffd-136">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="e6ffd-136">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="e6ffd-137">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="e6ffd-137">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
