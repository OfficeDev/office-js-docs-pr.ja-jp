---
title: 画像強制型変換要件セット
description: Excel、PowerPoint、Word で Office アドインを使用して、画像の強制型変換の要件セットをサポートします。
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 83817bfc7cf8a193138a805b0e90b4357d605801
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596971"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="0aafe-103">画像強制型変換要件セット</span><span class="sxs-lookup"><span data-stu-id="0aafe-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="0aafe-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="0aafe-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="0aafe-107">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="0aafe-107">ImageCoercion 1.1</span></span>

<span data-ttu-id="0aafe-108">ImageCoercion 1.1 は、メソッドを使用し`Office.CoercionType.Image`てデータを書き込むときに[`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 、image () への変換を有効にします。</span><span class="sxs-lookup"><span data-stu-id="0aafe-108">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="0aafe-109">次のホストがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="0aafe-109">The following hosts are supported:</span></span>

- <span data-ttu-id="0aafe-110">Excel 2013 以降</span><span class="sxs-lookup"><span data-stu-id="0aafe-110">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="0aafe-111">Excel 2016 以降 (Mac)</span><span class="sxs-lookup"><span data-stu-id="0aafe-111">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="0aafe-112">Excel on iPad</span><span class="sxs-lookup"><span data-stu-id="0aafe-112">Excel on iPad</span></span>
- <span data-ttu-id="0aafe-113">OneNote on the web</span><span class="sxs-lookup"><span data-stu-id="0aafe-113">OneNote on the web</span></span>
- <span data-ttu-id="0aafe-114">PowerPoint 2013 以降</span><span class="sxs-lookup"><span data-stu-id="0aafe-114">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="0aafe-115">PowerPoint 2016 以降の Mac</span><span class="sxs-lookup"><span data-stu-id="0aafe-115">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="0aafe-116">PowerPoint on the web</span><span class="sxs-lookup"><span data-stu-id="0aafe-116">PowerPoint on the web</span></span>
- <span data-ttu-id="0aafe-117">PowerPoint on iPad</span><span class="sxs-lookup"><span data-stu-id="0aafe-117">PowerPoint on iPad</span></span>
- <span data-ttu-id="0aafe-118">Word on Windows (Word 2013 以降)</span><span class="sxs-lookup"><span data-stu-id="0aafe-118">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="0aafe-119">Word on Mac (Word 2016 以降)</span><span class="sxs-lookup"><span data-stu-id="0aafe-119">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="0aafe-120">Word on the web</span><span class="sxs-lookup"><span data-stu-id="0aafe-120">Word on the web</span></span>
- <span data-ttu-id="0aafe-121">Word on iPad</span><span class="sxs-lookup"><span data-stu-id="0aafe-121">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="0aafe-122">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="0aafe-122">ImageCoercion 1.2</span></span>

<span data-ttu-id="0aafe-123">ImageCoercion 1.2 は、メソッドを使用し`Office.CoercionType.XmlSvg`てデータを書き込むときに[`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) SVG 形式 () への変換を有効にします。</span><span class="sxs-lookup"><span data-stu-id="0aafe-123">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="0aafe-124">次のホストがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="0aafe-124">The following hosts are supported:</span></span>

- <span data-ttu-id="0aafe-125">Windows 上の Excel (Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="0aafe-125">Excel on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="0aafe-126">Excel on Mac (Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="0aafe-126">Excel on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="0aafe-127">Windows 上の PowerPoint (Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="0aafe-127">PowerPoint on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="0aafe-128">PowerPoint on Mac (Office 365 サブスクリプションに接続されている)</span><span class="sxs-lookup"><span data-stu-id="0aafe-128">PowerPoint on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="0aafe-129">PowerPoint on the web</span><span class="sxs-lookup"><span data-stu-id="0aafe-129">PowerPoint on the web</span></span>
- <span data-ttu-id="0aafe-130">Windows 上の Word (Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="0aafe-130">Word on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="0aafe-131">Mac 上の Word (Office 365 サブスクリプションに接続されている)</span><span class="sxs-lookup"><span data-stu-id="0aafe-131">Word on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="0aafe-132">Word on the web</span><span class="sxs-lookup"><span data-stu-id="0aafe-132">Word on the web</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="0aafe-133">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="0aafe-133">Office Common API requirement sets</span></span>

<span data-ttu-id="0aafe-134">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="0aafe-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="0aafe-135">関連項目</span><span class="sxs-lookup"><span data-stu-id="0aafe-135">See also</span></span>

- [<span data-ttu-id="0aafe-136">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="0aafe-136">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="0aafe-137">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="0aafe-137">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="0aafe-138">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="0aafe-138">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
