---
title: 画像強制型変換要件セット
description: Excel、PowerPoint、Word で Office アドインを使用して、画像の強制型変換の要件セットをサポートします。
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 7140099757c6e4b5ad405723d5fed95fded6d919
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293549"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="db0cd-103">画像強制型変換要件セット</span><span class="sxs-lookup"><span data-stu-id="db0cd-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="db0cd-104">要件セットは、API メンバーの名前付きグループです。</span><span class="sxs-lookup"><span data-stu-id="db0cd-104">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="db0cd-105">Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイムチェックを使用して、Office アプリケーションがアドインに必要な Api をサポートしているかどうかを判断します。</span><span class="sxs-lookup"><span data-stu-id="db0cd-105">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs.</span></span> <span data-ttu-id="db0cd-106">詳細については、「 [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="db0cd-106">For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="db0cd-107">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="db0cd-107">ImageCoercion 1.1</span></span>

<span data-ttu-id="db0cd-108">ImageCoercion 1.1 `Office.CoercionType.Image` は、メソッドを使用してデータを書き込むときに、image () への変換を有効に [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) します。</span><span class="sxs-lookup"><span data-stu-id="db0cd-108">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="db0cd-109">サポートされているアプリケーションは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="db0cd-109">The following applications are supported:</span></span>

- <span data-ttu-id="db0cd-110">Excel 2013 以降</span><span class="sxs-lookup"><span data-stu-id="db0cd-110">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="db0cd-111">Excel 2016 以降 (Mac)</span><span class="sxs-lookup"><span data-stu-id="db0cd-111">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="db0cd-112">Excel on iPad</span><span class="sxs-lookup"><span data-stu-id="db0cd-112">Excel on iPad</span></span>
- <span data-ttu-id="db0cd-113">OneNote on the web</span><span class="sxs-lookup"><span data-stu-id="db0cd-113">OneNote on the web</span></span>
- <span data-ttu-id="db0cd-114">PowerPoint 2013 以降</span><span class="sxs-lookup"><span data-stu-id="db0cd-114">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="db0cd-115">PowerPoint 2016 以降の Mac</span><span class="sxs-lookup"><span data-stu-id="db0cd-115">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="db0cd-116">PowerPoint on the web</span><span class="sxs-lookup"><span data-stu-id="db0cd-116">PowerPoint on the web</span></span>
- <span data-ttu-id="db0cd-117">PowerPoint on iPad</span><span class="sxs-lookup"><span data-stu-id="db0cd-117">PowerPoint on iPad</span></span>
- <span data-ttu-id="db0cd-118">Word on Windows (Word 2013 以降)</span><span class="sxs-lookup"><span data-stu-id="db0cd-118">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="db0cd-119">Word on Mac (Word 2016 以降)</span><span class="sxs-lookup"><span data-stu-id="db0cd-119">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="db0cd-120">Word on the web</span><span class="sxs-lookup"><span data-stu-id="db0cd-120">Word on the web</span></span>
- <span data-ttu-id="db0cd-121">Word on iPad</span><span class="sxs-lookup"><span data-stu-id="db0cd-121">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="db0cd-122">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="db0cd-122">ImageCoercion 1.2</span></span>

<span data-ttu-id="db0cd-123">ImageCoercion 1.2 `Office.CoercionType.XmlSvg` は、メソッドを使用してデータを書き込むときに SVG 形式 () への変換を有効にし [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) ます。</span><span class="sxs-lookup"><span data-stu-id="db0cd-123">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="db0cd-124">サポートされているアプリケーションは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="db0cd-124">The following applications are supported:</span></span>

- <span data-ttu-id="db0cd-125">Windows 上の Excel (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="db0cd-125">Excel on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="db0cd-126">Mac 上の Excel (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="db0cd-126">Excel on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="db0cd-127">Windows 上の PowerPoint (Microsoft 365 サブスクリプションに接続されています)</span><span class="sxs-lookup"><span data-stu-id="db0cd-127">PowerPoint on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="db0cd-128">PowerPoint on Mac (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="db0cd-128">PowerPoint on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="db0cd-129">PowerPoint on the web</span><span class="sxs-lookup"><span data-stu-id="db0cd-129">PowerPoint on the web</span></span>
- <span data-ttu-id="db0cd-130">Windows 上の Word (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="db0cd-130">Word on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="db0cd-131">Mac 上の Word (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="db0cd-131">Word on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="db0cd-132">Word on the web</span><span class="sxs-lookup"><span data-stu-id="db0cd-132">Word on the web</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="db0cd-133">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="db0cd-133">Office Common API requirement sets</span></span>

<span data-ttu-id="db0cd-134">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="db0cd-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="db0cd-135">関連項目</span><span class="sxs-lookup"><span data-stu-id="db0cd-135">See also</span></span>

- [<span data-ttu-id="db0cd-136">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="db0cd-136">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="db0cd-137">Office アプリケーションと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="db0cd-137">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="db0cd-138">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="db0cd-138">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
