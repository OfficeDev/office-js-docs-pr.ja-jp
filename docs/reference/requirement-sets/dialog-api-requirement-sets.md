---
title: ダイアログ API の要件セット
description: ''
ms.date: 03/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: daa16235101a32adda8d462056a2db94a0da2d4f
ms.sourcegitcommit: 05b73cdec5f4db7f0b8d48a5a552ee296a0332ca
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/11/2020
ms.locfileid: "42600719"
---
# <a name="dialog-api-requirement-sets"></a><span data-ttu-id="5e9d8-102">ダイアログ API の要件セット</span><span class="sxs-lookup"><span data-stu-id="5e9d8-102">Dialog API requirement sets</span></span>

<span data-ttu-id="5e9d8-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="5e9d8-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="5e9d8-p102">Office アドインは Office の複数のバージョンで機能します。次の表は、ダイアログ API の要件セット、その要件セットをサポートする Office ホスト アプリケーション、Office アプリケーションのビルド番号またはバージョン番号の一覧です。</span><span class="sxs-lookup"><span data-stu-id="5e9d8-p102">Office Add-ins run across multiple versions of Office. The following table lists the Dialog API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="5e9d8-108">要件セット</span><span class="sxs-lookup"><span data-stu-id="5e9d8-108">Requirement set</span></span>  | <span data-ttu-id="5e9d8-109">Windows 版 Office 2013\*</span><span class="sxs-lookup"><span data-stu-id="5e9d8-109">Office 2013 on Windows\*</span></span><br><span data-ttu-id="5e9d8-110">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5e9d8-110">(one-time purchase)</span></span> | <span data-ttu-id="5e9d8-111">Office 2016 以降 (Windows)\*</span><span class="sxs-lookup"><span data-stu-id="5e9d8-111">Office 2016 or later on Windows\*</span></span><br><span data-ttu-id="5e9d8-112">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5e9d8-112">(one-time purchase)</span></span>   | <span data-ttu-id="5e9d8-113">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="5e9d8-113">Office on Windows</span></span><br><span data-ttu-id="5e9d8-114">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="5e9d8-114">(connected to Office 365 subscription)</span></span> |  <span data-ttu-id="5e9d8-115">Office on iPad</span><span class="sxs-lookup"><span data-stu-id="5e9d8-115">Office on iPad</span></span><br><span data-ttu-id="5e9d8-116">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="5e9d8-116">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="5e9d8-117">Office on Mac</span><span class="sxs-lookup"><span data-stu-id="5e9d8-117">Office on Mac</span></span><br><span data-ttu-id="5e9d8-118">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="5e9d8-118">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="5e9d8-119">Office on the web</span><span class="sxs-lookup"><span data-stu-id="5e9d8-119">Office on the web</span></span>  |  <span data-ttu-id="5e9d8-120">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="5e9d8-120">Office Online Server</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="5e9d8-121">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="5e9d8-121">DialogApi 1.1</span></span>  | <span data-ttu-id="5e9d8-122">ビルド 15.0.4855.1000 以降</span><span class="sxs-lookup"><span data-stu-id="5e9d8-122">Build 15.0.4855.1000 or later</span></span> | <span data-ttu-id="5e9d8-123">ビルド 16.0.4390.1000 以降</span><span class="sxs-lookup"><span data-stu-id="5e9d8-123">Build 16.0.4390.1000 or later</span></span> | <span data-ttu-id="5e9d8-124">バージョン 1602 (ビルド 6741.0000) 以降</span><span class="sxs-lookup"><span data-stu-id="5e9d8-124">Version 1602 (Build 6741.0000) or later</span></span> | <span data-ttu-id="5e9d8-125">1.22 以降</span><span class="sxs-lookup"><span data-stu-id="5e9d8-125">1.22 or later</span></span> | <span data-ttu-id="5e9d8-126">15.20 以降</span><span class="sxs-lookup"><span data-stu-id="5e9d8-126">15.20 or later</span></span>| <span data-ttu-id="5e9d8-127">2017 年 1 月</span><span class="sxs-lookup"><span data-stu-id="5e9d8-127">January 2017</span></span> | <span data-ttu-id="5e9d8-128">バージョン 1608 (ビルド 7601.6800) 以降</span><span class="sxs-lookup"><span data-stu-id="5e9d8-128">Version 1608 (Build 7601.6800) or later</span></span>|

><span data-ttu-id="5e9d8-129">\*ワンタイム購入オフィスのユーザーは、すべての修正プログラムと更新を承諾していない場合があります。</span><span class="sxs-lookup"><span data-stu-id="5e9d8-129">\* Users of the one-time purchase Office may not have accepted all patches and updates.</span></span> <span data-ttu-id="5e9d8-130">その場合、Office が UI でそのバージョンを報告するために使用する DLL が、ユーザーのコンピューターにインストールされていない更新された Dll がインストールされていない場合でも、ここにリストされているバージョンよりも大きくなる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="5e9d8-130">If so, the DLL that Office uses to report its version in the UI may be greater than the versions listed here even if the updated DLLs needed to support DialogApi have not be installed on the user's computer.</span></span> <span data-ttu-id="5e9d8-131">必要な修正プログラムがインストールされていることを確認するには、ユーザーは Office 更新プログラムの一覧 ([office 2013 リスト](/officeupdates/msp-files-office-2013)または[office 2016 の一覧](/officeupdates/msp-files-office-2016)) に移動し、 **osfclient**を検索して、一覧に記載されている修正プログラムをインストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="5e9d8-131">To ensure that the needed patch is installed, the user must go to the Office update list ([Office 2013 list](/officeupdates/msp-files-office-2013) or [Office 2016 list](/officeupdates/msp-files-office-2016)), search for **osfclient-x-none**, and install the listed patch.</span></span>

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="5e9d8-132">Office のバージョンとビルド番号</span><span class="sxs-lookup"><span data-stu-id="5e9d8-132">Office versions and build numbers</span></span>

<span data-ttu-id="5e9d8-133">バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5e9d8-133">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [<span data-ttu-id="5e9d8-134">Office Online Server 概要</span><span class="sxs-lookup"><span data-stu-id="5e9d8-134">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="5e9d8-135">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="5e9d8-135">Office Common API requirement sets</span></span>

<span data-ttu-id="5e9d8-136">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="5e9d8-136">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="dialog-api-11"></a><span data-ttu-id="5e9d8-137">ダイアログ API 1.1</span><span class="sxs-lookup"><span data-stu-id="5e9d8-137">Dialog API 1.1</span></span>

<span data-ttu-id="5e9d8-138">ダイアログ API 1.1 は、API の最初のバージョンです。</span><span class="sxs-lookup"><span data-stu-id="5e9d8-138">The Dialog API 1.1 is the first version of the API.</span></span> <span data-ttu-id="5e9d8-139">API の詳細については、「 [DIALOG api](/javascript/api/office/office.ui)リファレンス」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5e9d8-139">For details about the API, see the [Dialog API](/javascript/api/office/office.ui) reference topic.</span></span>

## <a name="see-also"></a><span data-ttu-id="5e9d8-140">関連項目</span><span class="sxs-lookup"><span data-stu-id="5e9d8-140">See also</span></span>

- [<span data-ttu-id="5e9d8-141">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="5e9d8-141">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="5e9d8-142">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="5e9d8-142">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="5e9d8-143">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="5e9d8-143">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
