---
title: 画像強制の要件セット
description: Excel、PowerPoint、Word で Office アドインを使用して、画像の強制型変換の要件セットをサポートします。
ms.date: 07/11/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 046a3f1f16d8b48cddbd64bddf80a31ed1e50583
ms.sourcegitcommit: 61f8f02193ce05da957418d938f0d94cb12c468d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/11/2019
ms.locfileid: "35633992"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="1b246-103">画像強制の要件セット</span><span class="sxs-lookup"><span data-stu-id="1b246-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="1b246-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="1b246-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="1b246-107">Office アドインは Office の複数のバージョンで機能します。</span><span class="sxs-lookup"><span data-stu-id="1b246-107">Office Add-ins run across multiple versions of Office.</span></span> <span data-ttu-id="1b246-108">次の表に、イメージ強制の要件セット、その要件セットをサポートする Office ホストアプリケーション、Office アプリケーションのビルド番号またはバージョン番号を示します。</span><span class="sxs-lookup"><span data-stu-id="1b246-108">The following table lists the Image Coercion requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="1b246-109">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="1b246-109">ImageCoercion 1.1</span></span>

<span data-ttu-id="1b246-110">ImageCoercion 1.1 は、メソッドを使用し`Office.CoercionType.Image`てデータを書き込むときに[`Document.setSelectedDataAsync`](/javascript/api/office/document#setselecteddataasync-data--options--callback-) 、image () への変換を有効にします。</span><span class="sxs-lookup"><span data-stu-id="1b246-110">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="1b246-111">次のホストがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="1b246-111">The following hosts are supported:</span></span>

- <span data-ttu-id="1b246-112">Excel 2013 以降</span><span class="sxs-lookup"><span data-stu-id="1b246-112">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="1b246-113">Excel 2016 以降 (Mac)</span><span class="sxs-lookup"><span data-stu-id="1b246-113">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="1b246-114">Excel on the web</span><span class="sxs-lookup"><span data-stu-id="1b246-114">Excel on the web</span></span>
- <span data-ttu-id="1b246-115">IPad の Excel</span><span class="sxs-lookup"><span data-stu-id="1b246-115">Excel on iPad</span></span>
- <span data-ttu-id="1b246-116">Web 上の OneNote</span><span class="sxs-lookup"><span data-stu-id="1b246-116">OneNote on the web</span></span>
- <span data-ttu-id="1b246-117">PowerPoint 2013 以降</span><span class="sxs-lookup"><span data-stu-id="1b246-117">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="1b246-118">PowerPoint 2016 以降の Mac</span><span class="sxs-lookup"><span data-stu-id="1b246-118">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="1b246-119">PowerPoint on the web</span><span class="sxs-lookup"><span data-stu-id="1b246-119">PowerPoint on the web</span></span>
- <span data-ttu-id="1b246-120">IPad の PowerPoint</span><span class="sxs-lookup"><span data-stu-id="1b246-120">PowerPoint on iPad</span></span>
- <span data-ttu-id="1b246-121">Word 2013 以降 (Windows)</span><span class="sxs-lookup"><span data-stu-id="1b246-121">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="1b246-122">Word 2016 以降の Mac</span><span class="sxs-lookup"><span data-stu-id="1b246-122">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="1b246-123">Web 上の Word</span><span class="sxs-lookup"><span data-stu-id="1b246-123">Word on the web</span></span>
- <span data-ttu-id="1b246-124">iPad の Word</span><span class="sxs-lookup"><span data-stu-id="1b246-124">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="1b246-125">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="1b246-125">ImageCoercion 1.2</span></span>

<span data-ttu-id="1b246-126">ImageCoercion 1.2 は、メソッドを使用し`Office.CoercionType.XmlSvg`てデータを書き込むときに[`Document.setSelectedDataAsync`](/javascript/api/office/document#setselecteddataasync-data--options--callback-) SVG 形式 () への変換を有効にします。</span><span class="sxs-lookup"><span data-stu-id="1b246-126">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="1b246-127">次のホストがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="1b246-127">The following hosts are supported:</span></span>

- <span data-ttu-id="1b246-128">Windows 上の Excel (Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="1b246-128">Excel on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="1b246-129">Excel on Mac (Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="1b246-129">Excel on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="1b246-130">Excel on the web</span><span class="sxs-lookup"><span data-stu-id="1b246-130">Excel on the web</span></span>
- <span data-ttu-id="1b246-131">Windows 上の PowerPoint (Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="1b246-131">PowerPoint on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="1b246-132">PowerPoint on Mac (Office 365 サブスクリプションに接続されている)</span><span class="sxs-lookup"><span data-stu-id="1b246-132">PowerPoint on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="1b246-133">PowerPoint on the web</span><span class="sxs-lookup"><span data-stu-id="1b246-133">PowerPoint on the web</span></span>
- <span data-ttu-id="1b246-134">Windows 上の Word (Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="1b246-134">Word on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="1b246-135">Mac 上の Word (Office 365 サブスクリプションに接続されている)</span><span class="sxs-lookup"><span data-stu-id="1b246-135">Word on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="1b246-136">Web 上の Word</span><span class="sxs-lookup"><span data-stu-id="1b246-136">Word on the web</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="1b246-137">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="1b246-137">Office Common API requirement sets</span></span>

<span data-ttu-id="1b246-138">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="1b246-138">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="1b246-139">関連項目</span><span class="sxs-lookup"><span data-stu-id="1b246-139">See also</span></span>

- [<span data-ttu-id="1b246-140">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="1b246-140">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="1b246-141">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="1b246-141">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="1b246-142">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="1b246-142">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
