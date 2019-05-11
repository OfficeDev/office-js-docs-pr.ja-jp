---
title: ダイアログ API の要件セット
description: ''
ms.date: 05/08/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: f6f0b0184736bfd0f6b417198ade4c621d8d8b6b
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952202"
---
# <a name="dialog-api-requirement-sets"></a><span data-ttu-id="5fb2a-102">ダイアログ API の要件セット</span><span class="sxs-lookup"><span data-stu-id="5fb2a-102">Dialog API requirement sets</span></span>

<span data-ttu-id="5fb2a-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="5fb2a-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="5fb2a-p102">Office アドインは Office の複数のバージョンで機能します。次の表は、ダイアログ API の要件セット、その要件セットをサポートする Office ホスト アプリケーション、Office アプリケーションのビルド番号またはバージョン番号の一覧です。</span><span class="sxs-lookup"><span data-stu-id="5fb2a-p102">Office Add-ins run across multiple versions of Office. The following table lists the Dialog API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="5fb2a-108">要件セット</span><span class="sxs-lookup"><span data-stu-id="5fb2a-108">Requirement set</span></span>  | <span data-ttu-id="5fb2a-109">Windows の Office 2013</span><span class="sxs-lookup"><span data-stu-id="5fb2a-109">Office 2013 on Windows</span></span><br><span data-ttu-id="5fb2a-110">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5fb2a-110">(one-time purchase)</span></span> | <span data-ttu-id="5fb2a-111">Office 2016 以降 (Windows)</span><span class="sxs-lookup"><span data-stu-id="5fb2a-111">Office 2016 or later on Windows</span></span><br><span data-ttu-id="5fb2a-112">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5fb2a-112">(one-time purchase)</span></span>   | <span data-ttu-id="5fb2a-113">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="5fb2a-113">Office on Windows</span></span><br><span data-ttu-id="5fb2a-114">(Office 365 に接続)</span><span class="sxs-lookup"><span data-stu-id="5fb2a-114">(connected to Office 365)</span></span> |  <span data-ttu-id="5fb2a-115">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="5fb2a-115">Office for iPad</span></span><br><span data-ttu-id="5fb2a-116">(Office 365 に接続)</span><span class="sxs-lookup"><span data-stu-id="5fb2a-116">(connected to Office 365)</span></span>  |  <span data-ttu-id="5fb2a-117">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="5fb2a-117">Office for Mac</span></span><br><span data-ttu-id="5fb2a-118">(Office 365 に接続)</span><span class="sxs-lookup"><span data-stu-id="5fb2a-118">(connected to Office 365)</span></span>  | <span data-ttu-id="5fb2a-119">Office Online</span><span class="sxs-lookup"><span data-stu-id="5fb2a-119">Office Online</span></span>  |  <span data-ttu-id="5fb2a-120">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="5fb2a-120">Office Online Server</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="5fb2a-121">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="5fb2a-121">DialogApi 1.1</span></span>  | <span data-ttu-id="5fb2a-122">ビルド 15.0.4855.1000 以降</span><span class="sxs-lookup"><span data-stu-id="5fb2a-122">Build 15.0.4855.1000 or later</span></span> | <span data-ttu-id="5fb2a-123">ビルド 16.0.4390.1000 以降</span><span class="sxs-lookup"><span data-stu-id="5fb2a-123">Build 16.0.4390.1000 or later</span></span> | <span data-ttu-id="5fb2a-124">バージョン 1602 (ビルド 6741.0000) 以降</span><span class="sxs-lookup"><span data-stu-id="5fb2a-124">Version 1602 (Build 6741.0000) or later</span></span> | <span data-ttu-id="5fb2a-125">1.22 以降</span><span class="sxs-lookup"><span data-stu-id="5fb2a-125">1.22 or later</span></span> | <span data-ttu-id="5fb2a-126">15.20 以降</span><span class="sxs-lookup"><span data-stu-id="5fb2a-126">15.20 or later</span></span>| <span data-ttu-id="5fb2a-127">2017 年 1 月</span><span class="sxs-lookup"><span data-stu-id="5fb2a-127">January 2017</span></span> | <span data-ttu-id="5fb2a-128">バージョン 1608 (ビルド 7601.6800) 以降</span><span class="sxs-lookup"><span data-stu-id="5fb2a-128">Version 1608 (Build 7601.6800) or later</span></span>|

<span data-ttu-id="5fb2a-129">バージョン、ビルド番号、および Office Online Server の詳細については以下を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5fb2a-129">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

- [<span data-ttu-id="5fb2a-130">Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号</span><span class="sxs-lookup"><span data-stu-id="5fb2a-130">Version and build numbers of update channel releases for Office 365 clients</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="5fb2a-131">使用している Office のバージョンを確認する方法</span><span class="sxs-lookup"><span data-stu-id="5fb2a-131">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [<span data-ttu-id="5fb2a-132">Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。</span><span class="sxs-lookup"><span data-stu-id="5fb2a-132">Where you can find the version and build number for an Office 365 client application</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="5fb2a-133">Office Online Server 概要</span><span class="sxs-lookup"><span data-stu-id="5fb2a-133">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="5fb2a-134">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="5fb2a-134">Office Common API requirement sets</span></span>

<span data-ttu-id="5fb2a-135">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="5fb2a-135">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="dialog-api-11"></a><span data-ttu-id="5fb2a-136">ダイアログ API 1.1</span><span class="sxs-lookup"><span data-stu-id="5fb2a-136">Dialog API 1.1</span></span>

<span data-ttu-id="5fb2a-137">ダイアログ API 1.1 は、API の最初のバージョンです。</span><span class="sxs-lookup"><span data-stu-id="5fb2a-137">The Dialog API 1.1 is the first version of the API.</span></span> <span data-ttu-id="5fb2a-138">API について詳しくは、[ダイアログ API](/javascript/api/office/office.ui) リファレンスのトピックをご覧ください。</span><span class="sxs-lookup"><span data-stu-id="5fb2a-138">For details about the API, see the [Dialog API ](/javascript/api/office/office.ui) reference topic.</span></span>

## <a name="see-also"></a><span data-ttu-id="5fb2a-139">関連項目</span><span class="sxs-lookup"><span data-stu-id="5fb2a-139">See also</span></span>

- [<span data-ttu-id="5fb2a-140">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="5fb2a-140">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="5fb2a-141">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="5fb2a-141">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="5fb2a-142">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="5fb2a-142">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
