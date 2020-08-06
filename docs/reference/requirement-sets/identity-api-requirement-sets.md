---
title: Identity API の要件セット
description: Id API の要件 Office アドインの情報を設定します。
ms.date: 07/30/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 05805451f17cc70597a61e55d1ecacbb81c383c5
ms.sourcegitcommit: 8fdd7369bfd97a273e222a0404e337ba2b8807b0
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/05/2020
ms.locfileid: "46573218"
---
# <a name="identity-api-requirement-sets"></a><span data-ttu-id="b54e1-103">Identity API の要件セット</span><span class="sxs-lookup"><span data-stu-id="b54e1-103">Identity API requirement sets</span></span>

<span data-ttu-id="b54e1-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="b54e1-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="b54e1-107">Office アドインは Office の複数のバージョンで機能します。</span><span class="sxs-lookup"><span data-stu-id="b54e1-107">Office Add-ins run across multiple versions of Office.</span></span> <span data-ttu-id="b54e1-108">次の表は、Identity API の要件セット、その要件セットをサポートする Office ホスト アプリケーション、Office アプリケーションのビルド番号またはバージョン番号の一覧です。</span><span class="sxs-lookup"><span data-stu-id="b54e1-108">The following table lists the Identity API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="b54e1-109">要件セット</span><span class="sxs-lookup"><span data-stu-id="b54e1-109">Requirement set</span></span>  | <span data-ttu-id="b54e1-110">Windows での Office 2013 以降</span><span class="sxs-lookup"><span data-stu-id="b54e1-110">Office 2013 or later on Windows</span></span><br><span data-ttu-id="b54e1-111">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="b54e1-111">(one-time purchase)</span></span> | <span data-ttu-id="b54e1-112">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="b54e1-112">Office on Windows</span></span><br><span data-ttu-id="b54e1-113">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="b54e1-113">(connected to a Microsoft 365 subscription)</span></span> |  <span data-ttu-id="b54e1-114">Office on iPad</span><span class="sxs-lookup"><span data-stu-id="b54e1-114">Office on iPad</span></span><br><span data-ttu-id="b54e1-115">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="b54e1-115">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="b54e1-116">Office on Mac</span><span class="sxs-lookup"><span data-stu-id="b54e1-116">Office on Mac</span></span><br><span data-ttu-id="b54e1-117">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="b54e1-117">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="b54e1-118">Office on the web</span><span class="sxs-lookup"><span data-stu-id="b54e1-118">Office on the web</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="b54e1-119">Identity Api 1.3</span><span class="sxs-lookup"><span data-stu-id="b54e1-119">IdentityAPI 1.3</span></span>  | <span data-ttu-id="b54e1-120">該当なし</span><span class="sxs-lookup"><span data-stu-id="b54e1-120">N/A</span></span> | <span data-ttu-id="b54e1-121">2008 (ビルド 13127.20000) 以降</span><span class="sxs-lookup"><span data-stu-id="b54e1-121">2008 (build 13127.20000) or later</span></span> | <span data-ttu-id="b54e1-122">近日対応予定</span><span class="sxs-lookup"><span data-stu-id="b54e1-122">Coming soon</span></span> | <span data-ttu-id="b54e1-123">16.40 以降</span><span class="sxs-lookup"><span data-stu-id="b54e1-123">16.40 or later</span></span> | <span data-ttu-id="b54e1-124">8月、2020 \*</span><span class="sxs-lookup"><span data-stu-id="b54e1-124">August, 2020\*</span></span> |

> <span data-ttu-id="b54e1-125">\*最初は、web 上の Office で要件セットがサポートされているのは、SharePoint Online および OneDrive.com から開かれたドキュメントのみです。</span><span class="sxs-lookup"><span data-stu-id="b54e1-125">\* Initially, the requirement set is supported in Office on the web only for documents that are opened from SharePoint Online and OneDrive.com.</span></span> <span data-ttu-id="b54e1-126">他のドキュメントのサポートは、2020の後の方に web 上の Office に送られます。</span><span class="sxs-lookup"><span data-stu-id="b54e1-126">Support for other documents will come to Office on the web later in 2020.</span></span>

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="b54e1-127">Office のバージョンとビルド番号</span><span class="sxs-lookup"><span data-stu-id="b54e1-127">Office versions and build numbers</span></span>

<span data-ttu-id="b54e1-128">バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b54e1-128">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [<span data-ttu-id="b54e1-129">Office Online Server 概要</span><span class="sxs-lookup"><span data-stu-id="b54e1-129">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="b54e1-130">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="b54e1-130">Office Common API requirement sets</span></span>

<span data-ttu-id="b54e1-131">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="b54e1-131">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="identityapi-preview"></a><span data-ttu-id="b54e1-132">Identity Api プレビュー</span><span class="sxs-lookup"><span data-stu-id="b54e1-132">IdentityAPI Preview</span></span>

<span data-ttu-id="b54e1-133">この API の詳細については、「 [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-)での約束を使用するバージョン」または[getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-)でコールバックを使用するバージョンのいずれかを参照してください。</span><span class="sxs-lookup"><span data-stu-id="b54e1-133">For details about this API, see either the version that uses Promises at [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) or the version that uses callbacks at [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-).</span></span>

## <a name="see-also"></a><span data-ttu-id="b54e1-134">関連項目</span><span class="sxs-lookup"><span data-stu-id="b54e1-134">See also</span></span>

- [<span data-ttu-id="b54e1-135">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="b54e1-135">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="b54e1-136">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="b54e1-136">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="b54e1-137">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="b54e1-137">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
