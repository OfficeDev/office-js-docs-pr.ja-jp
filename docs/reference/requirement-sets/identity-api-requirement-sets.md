---
title: Identity API の要件セット
description: ID API 要件は、アドインOffice情報を設定します。
ms.date: 01/26/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: c662e7a5306692fd75de51acc7cadfd1df3e7406
ms.sourcegitcommit: 85b4839be743059bf155ff44e49d64968444d80a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/31/2021
ms.locfileid: "51471725"
---
# <a name="identity-api-requirement-sets"></a><span data-ttu-id="d5d24-103">ID API の要件セット</span><span class="sxs-lookup"><span data-stu-id="d5d24-103">Identity API requirement sets</span></span>

<span data-ttu-id="d5d24-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="d5d24-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="d5d24-107">Office アドインは Office の複数のバージョンで機能します。</span><span class="sxs-lookup"><span data-stu-id="d5d24-107">Office Add-ins run across multiple versions of Office.</span></span> <span data-ttu-id="d5d24-108">次の表に、Identity API 要件セット、その要件セットをサポートする Office クライアント アプリケーション、およびアプリケーションのビルドまたはバージョン番号をOfficeします。</span><span class="sxs-lookup"><span data-stu-id="d5d24-108">The following table lists the Identity API requirement sets, the Office client applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="d5d24-109">要件セット</span><span class="sxs-lookup"><span data-stu-id="d5d24-109">Requirement set</span></span>  | <span data-ttu-id="d5d24-110">Windows での Office 2013 以降</span><span class="sxs-lookup"><span data-stu-id="d5d24-110">Office 2013 or later on Windows</span></span><br><span data-ttu-id="d5d24-111">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="d5d24-111">(one-time purchase)</span></span> | <span data-ttu-id="d5d24-112">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="d5d24-112">Office on Windows</span></span><br><span data-ttu-id="d5d24-113">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="d5d24-113">(connected to a Microsoft 365 subscription)</span></span> |  <span data-ttu-id="d5d24-114">Office on iPad</span><span class="sxs-lookup"><span data-stu-id="d5d24-114">Office on iPad</span></span><br><span data-ttu-id="d5d24-115">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="d5d24-115">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="d5d24-116">Office on Mac</span><span class="sxs-lookup"><span data-stu-id="d5d24-116">Office on Mac</span></span><br><span data-ttu-id="d5d24-117">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="d5d24-117">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="d5d24-118">Office on the web</span><span class="sxs-lookup"><span data-stu-id="d5d24-118">Office on the web</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="d5d24-119">IdentityAPI 1.3</span><span class="sxs-lookup"><span data-stu-id="d5d24-119">IdentityAPI 1.3</span></span>  | <span data-ttu-id="d5d24-120">該当なし</span><span class="sxs-lookup"><span data-stu-id="d5d24-120">N/A</span></span> | <span data-ttu-id="d5d24-121">2008 (ビルド 13127.20000) 以降</span><span class="sxs-lookup"><span data-stu-id="d5d24-121">2008 (build 13127.20000) or later</span></span> | <span data-ttu-id="d5d24-122">近日対応予定</span><span class="sxs-lookup"><span data-stu-id="d5d24-122">Coming soon</span></span> | <span data-ttu-id="d5d24-123">16.40 以降</span><span class="sxs-lookup"><span data-stu-id="d5d24-123">16.40 or later</span></span> | <span data-ttu-id="d5d24-124">Microsoft SharePoint Online と OneDrive\*</span><span class="sxs-lookup"><span data-stu-id="d5d24-124">Microsoft SharePoint Online and OneDrive\*</span></span> |

<span data-ttu-id="d5d24-125">\* 現在、要件セットは、Microsoft SharePoint Online Office OneDrive から開いたドキュメントに対してのみ、Web 上でサポートされています。</span><span class="sxs-lookup"><span data-stu-id="d5d24-125">\* Currently, the requirement set is supported in Office on the web only for documents that are opened from Microsoft SharePoint Online and OneDrive.</span></span>

> [!NOTE]
> <span data-ttu-id="d5d24-126">Outlook: アドイン コードで Identity API セット 1.3 を要求するには、呼び出しでサポートされていないか確認します `isSetSupported('IdentityAPI', '1.3')` 。</span><span class="sxs-lookup"><span data-stu-id="d5d24-126">Outlook: To require the Identity API set 1.3 in your add-in code, check if it's supported by calling `isSetSupported('IdentityAPI', '1.3')`.</span></span> <span data-ttu-id="d5d24-127">Outlook アドインのマニフェストで宣言はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d5d24-127">Declaring it in the Outlook add-in's manifest isn't supported.</span></span> <span data-ttu-id="d5d24-128">`undefined` ではないことを確認することで、API がサポートされているかどうかを判断することもできます。</span><span class="sxs-lookup"><span data-stu-id="d5d24-128">You can also determine if the API is supported by checking that it's not `undefined`.</span></span> <span data-ttu-id="d5d24-129">詳細については、「[後続の要件セットからの API の使用](outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d5d24-129">For further details, see [Using APIs from later requirement sets](outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets).</span></span>

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="d5d24-130">Office のバージョンとビルド番号</span><span class="sxs-lookup"><span data-stu-id="d5d24-130">Office versions and build numbers</span></span>

<span data-ttu-id="d5d24-131">バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d5d24-131">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [<span data-ttu-id="d5d24-132">Office Online Server 概要</span><span class="sxs-lookup"><span data-stu-id="d5d24-132">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="d5d24-133">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="d5d24-133">Office Common API requirement sets</span></span>

<span data-ttu-id="d5d24-134">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="d5d24-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="identityapi-preview"></a><span data-ttu-id="d5d24-135">IdentityAPI プレビュー</span><span class="sxs-lookup"><span data-stu-id="d5d24-135">IdentityAPI Preview</span></span>

<span data-ttu-id="d5d24-136">この API の詳細については [、getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) で Promises を使用するバージョン、または [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-)でコールバックを使用するバージョンを参照してください。</span><span class="sxs-lookup"><span data-stu-id="d5d24-136">For details about this API, see either the version that uses Promises at [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) or the version that uses callbacks at [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-).</span></span>

## <a name="see-also"></a><span data-ttu-id="d5d24-137">関連項目</span><span class="sxs-lookup"><span data-stu-id="d5d24-137">See also</span></span>

- [<span data-ttu-id="d5d24-138">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="d5d24-138">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="d5d24-139">Office アプリケーションと API 要件を指定する</span><span class="sxs-lookup"><span data-stu-id="d5d24-139">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="d5d24-140">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="d5d24-140">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
