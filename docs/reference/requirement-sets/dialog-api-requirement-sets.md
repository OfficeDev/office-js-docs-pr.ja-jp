---
title: ダイアログ API の要件セット
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: ad0d472ebdcbdb9d61e78f6bdc9bfe7c08311cd7
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432656"
---
# <a name="dialog-api-requirement-sets"></a><span data-ttu-id="94755-102">ダイアログ API の要件セット</span><span class="sxs-lookup"><span data-stu-id="94755-102">Dialog API requirement sets</span></span>

<span data-ttu-id="94755-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="94755-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="94755-p102">Office アドインは Office の複数のバージョンで機能します。次の表は、ダイアログ API の要件セット、その要件セットをサポートする Office ホスト アプリケーション、Office アプリケーションのビルド番号またはバージョン番号の一覧です。</span><span class="sxs-lookup"><span data-stu-id="94755-p102">Office Add-ins run across multiple versions of Office. The following table lists the Dialog API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="94755-108">要件セット</span><span class="sxs-lookup"><span data-stu-id="94755-108">Requirement set</span></span>  | <span data-ttu-id="94755-109">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="94755-109">Office 2013 for Windows</span></span> | <span data-ttu-id="94755-110">Office 2016 for Windows (MSI インストール)</span><span class="sxs-lookup"><span data-stu-id="94755-110">Office 2016 for Windows (MSI Installs)</span></span>   | <span data-ttu-id="94755-111">Office 365 for Windows (C2R インストール)</span><span class="sxs-lookup"><span data-stu-id="94755-111">Office 2016 for Windows (C2R Installs)</span></span>   |  <span data-ttu-id="94755-112">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="94755-112">Office 365 for iPad</span></span>  |  <span data-ttu-id="94755-113">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="94755-113">Office 365 for Mac</span></span>  | <span data-ttu-id="94755-114">Office Online</span><span class="sxs-lookup"><span data-stu-id="94755-114">Office Online</span></span>  |  <span data-ttu-id="94755-115">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="94755-115">Office Online Server</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="94755-116">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="94755-116">DialogApi 1.1</span></span>  | <span data-ttu-id="94755-117">ビルド 15.0.4855.1000 以降</span><span class="sxs-lookup"><span data-stu-id="94755-117">Build 15.0.4855.1000 or later</span></span> | <span data-ttu-id="94755-118">ビルド 16.0.4390.1000 以降</span><span class="sxs-lookup"><span data-stu-id="94755-118">Build 16.0.4390.1000 or later</span></span> | <span data-ttu-id="94755-119">バージョン 1602 (ビルド 6741.0000) 以降</span><span class="sxs-lookup"><span data-stu-id="94755-119">Version 1602 (Build 6741.0000) or later</span></span> | <span data-ttu-id="94755-120">1.22 以降</span><span class="sxs-lookup"><span data-stu-id="94755-120">1.22 or later</span></span> | <span data-ttu-id="94755-121">15.20 以降</span><span class="sxs-lookup"><span data-stu-id="94755-121">15.20 or later</span></span>| <span data-ttu-id="94755-122">2017 年 1 月</span><span class="sxs-lookup"><span data-stu-id="94755-122">January 2017</span></span> | <span data-ttu-id="94755-123">バージョン 1608 (ビルド 7601.6800) 以降</span><span class="sxs-lookup"><span data-stu-id="94755-123">Version 1608 (Build 7601.6800) or later</span></span>|

<span data-ttu-id="94755-124">バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。</span><span class="sxs-lookup"><span data-stu-id="94755-124">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

- [<span data-ttu-id="94755-125">Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号</span><span class="sxs-lookup"><span data-stu-id="94755-125">Version and build numbers of update channel releases for Office 365 clients</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="94755-126">使用している Office のバージョンを確認する方法</span><span class="sxs-lookup"><span data-stu-id="94755-126">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [<span data-ttu-id="94755-127">Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。</span><span class="sxs-lookup"><span data-stu-id="94755-127">Where you can find the version and build number for an Office 365 client application</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="94755-128">Office Online Server 概要</span><span class="sxs-lookup"><span data-stu-id="94755-128">Office Online Server overview</span></span>](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="94755-129">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="94755-129">Office common API requirement sets</span></span>

<span data-ttu-id="94755-130">共通 API の要件セットについて詳しくは、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="94755-130">For information about common API requirement sets, see [Office common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="dialog-api-11"></a><span data-ttu-id="94755-131">ダイアログ API 1.1</span><span class="sxs-lookup"><span data-stu-id="94755-131">Dialog API 1.1</span></span> 

<span data-ttu-id="94755-132">ダイアログ API 1.1 は、API の最初のバージョンです。</span><span class="sxs-lookup"><span data-stu-id="94755-132">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="94755-133">API について詳しくは、[ダイアログ API](/javascript/api/office/office.ui) リファレンスのトピックをご覧ください。</span><span class="sxs-lookup"><span data-stu-id="94755-133">For details about the API, see the [getAccessTokenAsync](/javascript/api/office/office.ui) reference topic.</span></span>

## <a name="see-also"></a><span data-ttu-id="94755-134">関連項目</span><span class="sxs-lookup"><span data-stu-id="94755-134">See also</span></span>

- [<span data-ttu-id="94755-135">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="94755-135">Office versions and requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="94755-136">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="94755-136">Specify Office hosts and API requirements</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="94755-137">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="94755-137">Office Add-ins XML manifest</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
