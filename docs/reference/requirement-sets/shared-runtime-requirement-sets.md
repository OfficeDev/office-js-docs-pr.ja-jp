---
title: 共有ランタイム要件セット
description: SharedRuntime Api をサポートするプラットフォームと Office ホストを指定します。
ms.date: 07/10/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 37ab904242a07a5ae7f1f580332f709ac409c6be
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159270"
---
# <a name="shared-runtime-requirement-sets"></a><span data-ttu-id="085b7-103">共有ランタイム要件セット</span><span class="sxs-lookup"><span data-stu-id="085b7-103">Shared runtime requirement sets</span></span>

<span data-ttu-id="085b7-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="085b7-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="085b7-107">JavaScript コードを実行する Office アドインの部分 (作業ウィンドウ、アドインコマンドから起動される関数ファイル、Excel カスタム関数) は、1つの JavaScript ランタイムを共有できます。</span><span class="sxs-lookup"><span data-stu-id="085b7-107">Parts of an Office Add-in that run JavaScript code, such as task panes, function files launched from add-in commands, and Excel custom functions, can share a single JavaScript runtime.</span></span> <span data-ttu-id="085b7-108">これにより、すべてのパーツが一連のグローバル変数を共有し、読み込まれたライブラリセットを共有して、永続的なストレージを介してメッセージを渡さずに相互に通信できるようになります。</span><span class="sxs-lookup"><span data-stu-id="085b7-108">This enables all the parts to share a set of global variables, to share a set of loaded libraries, and to communicate with each other without having to pass messages through a persisted storage.</span></span>

<span data-ttu-id="085b7-109">次の表に、SharedRuntime 1.1 の要件セット、その要件セットをサポートする Office ホストアプリケーション、Office アプリケーションのビルド番号またはバージョン番号を示します。</span><span class="sxs-lookup"><span data-stu-id="085b7-109">The following table lists the SharedRuntime 1.1 requirement set, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="085b7-110">要件セット</span><span class="sxs-lookup"><span data-stu-id="085b7-110">Requirement set</span></span>  |  <span data-ttu-id="085b7-111">Windows での Office 2013 (またはそれ以降のバージョン)</span><span class="sxs-lookup"><span data-stu-id="085b7-111">Office 2013 (or later) on Windows</span></span><br><span data-ttu-id="085b7-112">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="085b7-112">(one-time purchase)</span></span> | <span data-ttu-id="085b7-113">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="085b7-113">Office on Windows</span></span><br><span data-ttu-id="085b7-114">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="085b7-114">(connected to a Microsoft 365 subscription)</span></span>   |  <span data-ttu-id="085b7-115">Office on iPad</span><span class="sxs-lookup"><span data-stu-id="085b7-115">Office on iPad</span></span><br><span data-ttu-id="085b7-116">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="085b7-116">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="085b7-117">Office on Mac</span><span class="sxs-lookup"><span data-stu-id="085b7-117">Office on Mac</span></span><br><span data-ttu-id="085b7-118">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="085b7-118">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="085b7-119">Office on the web</span><span class="sxs-lookup"><span data-stu-id="085b7-119">Office on the web</span></span>  | <span data-ttu-id="085b7-120">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="085b7-120">Office Online Server</span></span> |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="085b7-121">SharedRuntime 1.1</span><span class="sxs-lookup"><span data-stu-id="085b7-121">SharedRuntime 1.1</span></span>  | <span data-ttu-id="085b7-122">該当なし</span><span class="sxs-lookup"><span data-stu-id="085b7-122">N/A</span></span> | <span data-ttu-id="085b7-123">バージョン 2002 (ビルド 12527.20092) 以降</span><span class="sxs-lookup"><span data-stu-id="085b7-123">Version 2002 (Build 12527.20092) or later</span></span> | <span data-ttu-id="085b7-124">該当なし</span><span class="sxs-lookup"><span data-stu-id="085b7-124">N/A</span></span> | <span data-ttu-id="085b7-125">16.35 以降</span><span class="sxs-lookup"><span data-stu-id="085b7-125">16.35 or later</span></span> | <span data-ttu-id="085b7-126">2020 年 2 月</span><span class="sxs-lookup"><span data-stu-id="085b7-126">February 2020</span></span> | <span data-ttu-id="085b7-127">該当なし</span><span class="sxs-lookup"><span data-stu-id="085b7-127">N/A</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="085b7-128">Office のバージョンとビルド番号</span><span class="sxs-lookup"><span data-stu-id="085b7-128">Office versions and build numbers</span></span>

<span data-ttu-id="085b7-129">バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。</span><span class="sxs-lookup"><span data-stu-id="085b7-129">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [<span data-ttu-id="085b7-130">Office Online Server 概要</span><span class="sxs-lookup"><span data-stu-id="085b7-130">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="085b7-131">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="085b7-131">Office Common API requirement sets</span></span>

<span data-ttu-id="085b7-132">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="085b7-132">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="085b7-133">関連項目</span><span class="sxs-lookup"><span data-stu-id="085b7-133">See also</span></span>

- [<span data-ttu-id="085b7-134">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="085b7-134">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="085b7-135">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="085b7-135">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="085b7-136">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="085b7-136">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
