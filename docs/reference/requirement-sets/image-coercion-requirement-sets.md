---
title: 画像強制型変換要件セット
description: Excel、PowerPoint、Word で Office アドインを使用して、画像の強制型変換の要件セットをサポートします。
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 59f6891182f47bed1b7e3b6aa69a30e941bce7cb
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094355"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="69de2-103">画像強制型変換要件セット</span><span class="sxs-lookup"><span data-stu-id="69de2-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="69de2-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="69de2-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="69de2-107">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="69de2-107">ImageCoercion 1.1</span></span>

<span data-ttu-id="69de2-108">ImageCoercion 1.1 `Office.CoercionType.Image` は、メソッドを使用してデータを書き込むときに、image () への変換を有効に [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) します。</span><span class="sxs-lookup"><span data-stu-id="69de2-108">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="69de2-109">次のホストがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="69de2-109">The following hosts are supported:</span></span>

- <span data-ttu-id="69de2-110">Excel 2013 以降</span><span class="sxs-lookup"><span data-stu-id="69de2-110">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="69de2-111">Excel 2016 以降 (Mac)</span><span class="sxs-lookup"><span data-stu-id="69de2-111">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="69de2-112">Excel on iPad</span><span class="sxs-lookup"><span data-stu-id="69de2-112">Excel on iPad</span></span>
- <span data-ttu-id="69de2-113">OneNote on the web</span><span class="sxs-lookup"><span data-stu-id="69de2-113">OneNote on the web</span></span>
- <span data-ttu-id="69de2-114">PowerPoint 2013 以降</span><span class="sxs-lookup"><span data-stu-id="69de2-114">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="69de2-115">PowerPoint 2016 以降の Mac</span><span class="sxs-lookup"><span data-stu-id="69de2-115">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="69de2-116">PowerPoint on the web</span><span class="sxs-lookup"><span data-stu-id="69de2-116">PowerPoint on the web</span></span>
- <span data-ttu-id="69de2-117">PowerPoint on iPad</span><span class="sxs-lookup"><span data-stu-id="69de2-117">PowerPoint on iPad</span></span>
- <span data-ttu-id="69de2-118">Word on Windows (Word 2013 以降)</span><span class="sxs-lookup"><span data-stu-id="69de2-118">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="69de2-119">Word on Mac (Word 2016 以降)</span><span class="sxs-lookup"><span data-stu-id="69de2-119">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="69de2-120">Word on the web</span><span class="sxs-lookup"><span data-stu-id="69de2-120">Word on the web</span></span>
- <span data-ttu-id="69de2-121">Word on iPad</span><span class="sxs-lookup"><span data-stu-id="69de2-121">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="69de2-122">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="69de2-122">ImageCoercion 1.2</span></span>

<span data-ttu-id="69de2-123">ImageCoercion 1.2 `Office.CoercionType.XmlSvg` は、メソッドを使用してデータを書き込むときに SVG 形式 () への変換を有効にし [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) ます。</span><span class="sxs-lookup"><span data-stu-id="69de2-123">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="69de2-124">次のホストがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="69de2-124">The following hosts are supported:</span></span>

- <span data-ttu-id="69de2-125">Windows 上の Excel (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="69de2-125">Excel on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="69de2-126">Mac 上の Excel (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="69de2-126">Excel on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="69de2-127">Windows 上の PowerPoint (Microsoft 365 サブスクリプションに接続されています)</span><span class="sxs-lookup"><span data-stu-id="69de2-127">PowerPoint on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="69de2-128">PowerPoint on Mac (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="69de2-128">PowerPoint on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="69de2-129">PowerPoint on the web</span><span class="sxs-lookup"><span data-stu-id="69de2-129">PowerPoint on the web</span></span>
- <span data-ttu-id="69de2-130">Windows 上の Word (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="69de2-130">Word on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="69de2-131">Mac 上の Word (Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="69de2-131">Word on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="69de2-132">Word on the web</span><span class="sxs-lookup"><span data-stu-id="69de2-132">Word on the web</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="69de2-133">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="69de2-133">Office Common API requirement sets</span></span>

<span data-ttu-id="69de2-134">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="69de2-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="69de2-135">関連項目</span><span class="sxs-lookup"><span data-stu-id="69de2-135">See also</span></span>

- [<span data-ttu-id="69de2-136">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="69de2-136">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="69de2-137">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="69de2-137">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="69de2-138">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="69de2-138">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
