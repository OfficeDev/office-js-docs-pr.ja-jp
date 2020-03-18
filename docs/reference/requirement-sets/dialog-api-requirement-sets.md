---
title: ダイアログ API の要件セット
description: ダイアログ API の要件セットの詳細情報
ms.date: 03/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: aeae2f140b158f3343c9812db8e9f27ea7608a3e
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719911"
---
# <a name="dialog-api-requirement-sets"></a><span data-ttu-id="0692b-103">ダイアログ API の要件セット</span><span class="sxs-lookup"><span data-stu-id="0692b-103">Dialog API requirement sets</span></span>

<span data-ttu-id="0692b-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="0692b-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="0692b-p102">Office アドインは Office の複数のバージョンで機能します。次の表は、ダイアログ API の要件セット、その要件セットをサポートする Office ホスト アプリケーション、Office アプリケーションのビルド番号またはバージョン番号の一覧です。</span><span class="sxs-lookup"><span data-stu-id="0692b-p102">Office Add-ins run across multiple versions of Office. The following table lists the Dialog API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="0692b-109">要件セット</span><span class="sxs-lookup"><span data-stu-id="0692b-109">Requirement set</span></span>  | <span data-ttu-id="0692b-110">Windows 版 Office 2013\*</span><span class="sxs-lookup"><span data-stu-id="0692b-110">Office 2013 on Windows\*</span></span><br><span data-ttu-id="0692b-111">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="0692b-111">(one-time purchase)</span></span> | <span data-ttu-id="0692b-112">Office 2016 以降 (Windows)\*</span><span class="sxs-lookup"><span data-stu-id="0692b-112">Office 2016 or later on Windows\*</span></span><br><span data-ttu-id="0692b-113">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="0692b-113">(one-time purchase)</span></span>   | <span data-ttu-id="0692b-114">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="0692b-114">Office on Windows</span></span><br><span data-ttu-id="0692b-115">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="0692b-115">(connected to Office 365 subscription)</span></span> |  <span data-ttu-id="0692b-116">Office on iPad</span><span class="sxs-lookup"><span data-stu-id="0692b-116">Office on iPad</span></span><br><span data-ttu-id="0692b-117">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="0692b-117">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="0692b-118">Office on Mac</span><span class="sxs-lookup"><span data-stu-id="0692b-118">Office on Mac</span></span><br><span data-ttu-id="0692b-119">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="0692b-119">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="0692b-120">Office on the web</span><span class="sxs-lookup"><span data-stu-id="0692b-120">Office on the web</span></span>  |  <span data-ttu-id="0692b-121">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="0692b-121">Office Online Server</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="0692b-122">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="0692b-122">DialogApi 1.1</span></span>  | <span data-ttu-id="0692b-123">ビルド 15.0.4855.1000 以降</span><span class="sxs-lookup"><span data-stu-id="0692b-123">Build 15.0.4855.1000 or later</span></span> | <span data-ttu-id="0692b-124">ビルド 16.0.4390.1000 以降</span><span class="sxs-lookup"><span data-stu-id="0692b-124">Build 16.0.4390.1000 or later</span></span> | <span data-ttu-id="0692b-125">バージョン 1602 (ビルド 6741.0000) 以降</span><span class="sxs-lookup"><span data-stu-id="0692b-125">Version 1602 (Build 6741.0000) or later</span></span> | <span data-ttu-id="0692b-126">1.22 以降</span><span class="sxs-lookup"><span data-stu-id="0692b-126">1.22 or later</span></span> | <span data-ttu-id="0692b-127">15.20 以降</span><span class="sxs-lookup"><span data-stu-id="0692b-127">15.20 or later</span></span>| <span data-ttu-id="0692b-128">2017 年 1 月</span><span class="sxs-lookup"><span data-stu-id="0692b-128">January 2017</span></span> | <span data-ttu-id="0692b-129">バージョン 1608 (ビルド 7601.6800) 以降</span><span class="sxs-lookup"><span data-stu-id="0692b-129">Version 1608 (Build 7601.6800) or later</span></span>|

><span data-ttu-id="0692b-130">\*ワンタイム購入オフィスのユーザーは、すべての修正プログラムと更新を承諾していない場合があります。</span><span class="sxs-lookup"><span data-stu-id="0692b-130">\* Users of the one-time purchase Office may not have accepted all patches and updates.</span></span> <span data-ttu-id="0692b-131">その場合、Office が UI でそのバージョンを報告するために使用する DLL が、ユーザーのコンピューターにインストールされていない更新された Dll がインストールされていない場合でも、ここにリストされているバージョンよりも大きくなる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="0692b-131">If so, the DLL that Office uses to report its version in the UI may be greater than the versions listed here even if the updated DLLs needed to support DialogApi have not be installed on the user's computer.</span></span> <span data-ttu-id="0692b-132">必要な修正プログラムがインストールされていることを確認するには、ユーザーは Office 更新プログラムの一覧 ([office 2013 リスト](/officeupdates/msp-files-office-2013)または[office 2016 の一覧](/officeupdates/msp-files-office-2016)) に移動し、 **osfclient**を検索して、一覧に記載されている修正プログラムをインストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="0692b-132">To ensure that the needed patch is installed, the user must go to the Office update list ([Office 2013 list](/officeupdates/msp-files-office-2013) or [Office 2016 list](/officeupdates/msp-files-office-2016)), search for **osfclient-x-none**, and install the listed patch.</span></span>

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="0692b-133">Office のバージョンとビルド番号</span><span class="sxs-lookup"><span data-stu-id="0692b-133">Office versions and build numbers</span></span>

<span data-ttu-id="0692b-134">バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0692b-134">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [<span data-ttu-id="0692b-135">Office Online Server 概要</span><span class="sxs-lookup"><span data-stu-id="0692b-135">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="0692b-136">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="0692b-136">Office Common API requirement sets</span></span>

<span data-ttu-id="0692b-137">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="0692b-137">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="dialog-api-11"></a><span data-ttu-id="0692b-138">ダイアログ API 1.1</span><span class="sxs-lookup"><span data-stu-id="0692b-138">Dialog API 1.1</span></span>

<span data-ttu-id="0692b-139">ダイアログ API 1.1 は、API の最初のバージョンです。</span><span class="sxs-lookup"><span data-stu-id="0692b-139">The Dialog API 1.1 is the first version of the API.</span></span> <span data-ttu-id="0692b-140">API の詳細については、「 [DIALOG api](/javascript/api/office/office.ui)リファレンス」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0692b-140">For details about the API, see the [Dialog API](/javascript/api/office/office.ui) reference topic.</span></span>

## <a name="see-also"></a><span data-ttu-id="0692b-141">関連項目</span><span class="sxs-lookup"><span data-stu-id="0692b-141">See also</span></span>

- [<span data-ttu-id="0692b-142">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="0692b-142">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="0692b-143">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="0692b-143">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="0692b-144">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="0692b-144">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
