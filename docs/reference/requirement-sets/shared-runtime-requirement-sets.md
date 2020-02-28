---
title: 共有ランタイム要件セット
description: SharedRuntime Api をサポートするプラットフォームと Office ホストを指定します。
ms.date: 02/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: dbb9d908154da074eaff6901c778adea168504a9
ms.sourcegitcommit: 7464eac3b54a6a6b65e27549a3ad603af6ee1011
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42315883"
---
# <a name="shared-runtime-requirement-sets"></a><span data-ttu-id="33698-103">共有ランタイム要件セット</span><span class="sxs-lookup"><span data-stu-id="33698-103">Shared runtime requirement sets</span></span>

<span data-ttu-id="33698-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="33698-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="33698-107">JavaScript コードを実行する Office アドインの部分 (作業ウィンドウ、アドインコマンドから起動される関数ファイル、Excel カスタム関数) は、1つの JavaScript ランタイムを共有できます。</span><span class="sxs-lookup"><span data-stu-id="33698-107">Parts of an Office Add-in that run JavaScript code, such as task panes, function files launched from add-in commands, and Excel custom functions, can share a single JavaScript runtime.</span></span> <span data-ttu-id="33698-108">これにより、すべてのパーツが一連のグローバル変数を共有し、読み込まれたライブラリセットを共有して、永続的なストレージを介してメッセージを渡さずに相互に通信できるようになります。</span><span class="sxs-lookup"><span data-stu-id="33698-108">This enables all the parts to share a set of global variables, to share a set of loaded libraries, and to communicate with each other without having to pass messages through a persisted storage.</span></span>

<span data-ttu-id="33698-109">次の表に、SharedRuntime 1.1 の要件セット、その要件セットをサポートする Office ホストアプリケーション、Office アプリケーションのビルド番号またはバージョン番号を示します。</span><span class="sxs-lookup"><span data-stu-id="33698-109">The following table lists the SharedRuntime 1.1 requirement set, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="33698-110">要件セット</span><span class="sxs-lookup"><span data-stu-id="33698-110">Requirement set</span></span>  |  <span data-ttu-id="33698-111">Windows での Office 2013 (またはそれ以降のバージョン)</span><span class="sxs-lookup"><span data-stu-id="33698-111">Office 2013 (or later) on Windows</span></span><br><span data-ttu-id="33698-112">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="33698-112">(one-time purchase)</span></span> | <span data-ttu-id="33698-113">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="33698-113">Office on Windows</span></span><br><span data-ttu-id="33698-114">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="33698-114">(connected to Office 365 subscription)</span></span>   |  <span data-ttu-id="33698-115">Office on iPad</span><span class="sxs-lookup"><span data-stu-id="33698-115">Office on iPad</span></span><br><span data-ttu-id="33698-116">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="33698-116">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="33698-117">Office on Mac</span><span class="sxs-lookup"><span data-stu-id="33698-117">Office on Mac</span></span><br><span data-ttu-id="33698-118">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="33698-118">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="33698-119">Office on the web</span><span class="sxs-lookup"><span data-stu-id="33698-119">Office on the web</span></span>  | <span data-ttu-id="33698-120">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="33698-120">Office Online Server</span></span> |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="33698-121">SharedRuntime 1.1</span><span class="sxs-lookup"><span data-stu-id="33698-121">SharedRuntime 1.1</span></span>  | <span data-ttu-id="33698-122">該当なし</span><span class="sxs-lookup"><span data-stu-id="33698-122">N/A</span></span> | <span data-ttu-id="33698-123">バージョン 2002 (ビルド 12527.20092) 以降</span><span class="sxs-lookup"><span data-stu-id="33698-123">Version 2002 (Build 12527.20092) or later</span></span> | <span data-ttu-id="33698-124">該当なし</span><span class="sxs-lookup"><span data-stu-id="33698-124">N/A</span></span> | <span data-ttu-id="33698-125">16.35 以降</span><span class="sxs-lookup"><span data-stu-id="33698-125">16.35 or later</span></span> | <span data-ttu-id="33698-126">2020 年 2 月</span><span class="sxs-lookup"><span data-stu-id="33698-126">February 2020</span></span> | <span data-ttu-id="33698-127">該当なし</span><span class="sxs-lookup"><span data-stu-id="33698-127">N/A</span></span> |

<span data-ttu-id="33698-128">バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。</span><span class="sxs-lookup"><span data-stu-id="33698-128">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

- [<span data-ttu-id="33698-129">Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号</span><span class="sxs-lookup"><span data-stu-id="33698-129">Version and build numbers of update channel releases for Office 365 clients</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="33698-130">使用している Office のバージョンを確認する方法</span><span class="sxs-lookup"><span data-stu-id="33698-130">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [<span data-ttu-id="33698-131">Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。</span><span class="sxs-lookup"><span data-stu-id="33698-131">Where you can find the version and build number for an Office 365 client application</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="33698-132">Office Online Server 概要</span><span class="sxs-lookup"><span data-stu-id="33698-132">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="33698-133">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="33698-133">Office Common API requirement sets</span></span>

<span data-ttu-id="33698-134">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="33698-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="33698-135">関連項目</span><span class="sxs-lookup"><span data-stu-id="33698-135">See also</span></span>

- [<span data-ttu-id="33698-136">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="33698-136">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="33698-137">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="33698-137">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="33698-138">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="33698-138">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
