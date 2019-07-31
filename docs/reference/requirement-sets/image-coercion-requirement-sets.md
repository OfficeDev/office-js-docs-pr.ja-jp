---
title: 画像強制の要件セット
description: Excel、PowerPoint、Word で Office アドインを使用して、画像の強制型変換の要件セットをサポートします。
ms.date: 07/11/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: bffe6c074d9e0734299d0087f2488524875931ed
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940851"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="fc68a-103">画像強制の要件セット</span><span class="sxs-lookup"><span data-stu-id="fc68a-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="fc68a-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="fc68a-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="fc68a-107">Office アドインは Office の複数のバージョンで機能します。</span><span class="sxs-lookup"><span data-stu-id="fc68a-107">Office Add-ins run across multiple versions of Office.</span></span> <span data-ttu-id="fc68a-108">次の表に、イメージ強制の要件セット、その要件セットをサポートする Office ホストアプリケーション、Office アプリケーションのビルド番号またはバージョン番号を示します。</span><span class="sxs-lookup"><span data-stu-id="fc68a-108">The following table lists the Image Coercion requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="fc68a-109">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="fc68a-109">ImageCoercion 1.1</span></span>

<span data-ttu-id="fc68a-110">ImageCoercion 1.1 は、メソッドを使用し`Office.CoercionType.Image`てデータを書き込むときに[`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 、image () への変換を有効にします。</span><span class="sxs-lookup"><span data-stu-id="fc68a-110">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="fc68a-111">次のホストがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="fc68a-111">The following hosts are supported:</span></span>

- <span data-ttu-id="fc68a-112">Excel 2013 以降</span><span class="sxs-lookup"><span data-stu-id="fc68a-112">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="fc68a-113">Excel 2016 以降 (Mac)</span><span class="sxs-lookup"><span data-stu-id="fc68a-113">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="fc68a-114">Excel on the web</span><span class="sxs-lookup"><span data-stu-id="fc68a-114">Excel on the web</span></span>
- <span data-ttu-id="fc68a-115">IPad の Excel</span><span class="sxs-lookup"><span data-stu-id="fc68a-115">Excel on iPad</span></span>
- <span data-ttu-id="fc68a-116">Web 上の OneNote</span><span class="sxs-lookup"><span data-stu-id="fc68a-116">OneNote on the web</span></span>
- <span data-ttu-id="fc68a-117">PowerPoint 2013 以降</span><span class="sxs-lookup"><span data-stu-id="fc68a-117">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="fc68a-118">PowerPoint 2016 以降の Mac</span><span class="sxs-lookup"><span data-stu-id="fc68a-118">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="fc68a-119">PowerPoint on the web</span><span class="sxs-lookup"><span data-stu-id="fc68a-119">PowerPoint on the web</span></span>
- <span data-ttu-id="fc68a-120">IPad の PowerPoint</span><span class="sxs-lookup"><span data-stu-id="fc68a-120">PowerPoint on iPad</span></span>
- <span data-ttu-id="fc68a-121">Word 2013 以降 (Windows)</span><span class="sxs-lookup"><span data-stu-id="fc68a-121">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="fc68a-122">Word 2016 以降の Mac</span><span class="sxs-lookup"><span data-stu-id="fc68a-122">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="fc68a-123">Web 上の Word</span><span class="sxs-lookup"><span data-stu-id="fc68a-123">Word on the web</span></span>
- <span data-ttu-id="fc68a-124">iPad の Word</span><span class="sxs-lookup"><span data-stu-id="fc68a-124">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="fc68a-125">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="fc68a-125">ImageCoercion 1.2</span></span>

<span data-ttu-id="fc68a-126">ImageCoercion 1.2 は、メソッドを使用し`Office.CoercionType.XmlSvg`てデータを書き込むときに[`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) SVG 形式 () への変換を有効にします。</span><span class="sxs-lookup"><span data-stu-id="fc68a-126">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="fc68a-127">次のホストがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="fc68a-127">The following hosts are supported:</span></span>

- <span data-ttu-id="fc68a-128">Windows 上の Excel (Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="fc68a-128">Excel on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="fc68a-129">Excel on Mac (Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="fc68a-129">Excel on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="fc68a-130">Excel on the web</span><span class="sxs-lookup"><span data-stu-id="fc68a-130">Excel on the web</span></span>
- <span data-ttu-id="fc68a-131">Windows 上の PowerPoint (Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="fc68a-131">PowerPoint on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="fc68a-132">PowerPoint on Mac (Office 365 サブスクリプションに接続されている)</span><span class="sxs-lookup"><span data-stu-id="fc68a-132">PowerPoint on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="fc68a-133">PowerPoint on the web</span><span class="sxs-lookup"><span data-stu-id="fc68a-133">PowerPoint on the web</span></span>
- <span data-ttu-id="fc68a-134">Windows 上の Word (Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="fc68a-134">Word on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="fc68a-135">Mac 上の Word (Office 365 サブスクリプションに接続されている)</span><span class="sxs-lookup"><span data-stu-id="fc68a-135">Word on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="fc68a-136">Web 上の Word</span><span class="sxs-lookup"><span data-stu-id="fc68a-136">Word on the web</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="fc68a-137">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="fc68a-137">Office Common API requirement sets</span></span>

<span data-ttu-id="fc68a-138">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="fc68a-138">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="fc68a-139">関連項目</span><span class="sxs-lookup"><span data-stu-id="fc68a-139">See also</span></span>

- [<span data-ttu-id="fc68a-140">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="fc68a-140">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="fc68a-141">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="fc68a-141">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="fc68a-142">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="fc68a-142">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
