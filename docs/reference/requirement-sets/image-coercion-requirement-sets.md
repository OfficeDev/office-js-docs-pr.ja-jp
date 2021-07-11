---
title: 画像強制型変換要件セット
description: 複数のアドインを使用した Image Coercion 要件セットOffice、Excel、Word PowerPointサポート。
ms.date: 02/19/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 29614718378fd51013360a2a922e11f89bca14b8
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350219"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="f451e-103">画像強制型変換要件セット</span><span class="sxs-lookup"><span data-stu-id="f451e-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="f451e-p101">要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="f451e-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="f451e-107">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="f451e-107">ImageCoercion 1.1</span></span>

<span data-ttu-id="f451e-108">ImageCoercion 1.1 では、メソッドを使用してデータを書き込むときにイメージ ( `Office.CoercionType.Image` ) への変換が有効 [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) です。</span><span class="sxs-lookup"><span data-stu-id="f451e-108">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="f451e-109">次のアプリケーションがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="f451e-109">The following applications are supported.</span></span>

- <span data-ttu-id="f451e-110">Excel 2013 以降のWindows</span><span class="sxs-lookup"><span data-stu-id="f451e-110">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="f451e-111">Excel 2016以降の Mac</span><span class="sxs-lookup"><span data-stu-id="f451e-111">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="f451e-112">Excel on iPad</span><span class="sxs-lookup"><span data-stu-id="f451e-112">Excel on iPad</span></span>
- <span data-ttu-id="f451e-113">OneNote on the web</span><span class="sxs-lookup"><span data-stu-id="f451e-113">OneNote on the web</span></span>
- <span data-ttu-id="f451e-114">PowerPoint 2013 以降のWindows</span><span class="sxs-lookup"><span data-stu-id="f451e-114">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="f451e-115">PowerPoint 2016以降の Mac</span><span class="sxs-lookup"><span data-stu-id="f451e-115">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="f451e-116">PowerPoint on the web</span><span class="sxs-lookup"><span data-stu-id="f451e-116">PowerPoint on the web</span></span>
- <span data-ttu-id="f451e-117">PowerPoint on iPad</span><span class="sxs-lookup"><span data-stu-id="f451e-117">PowerPoint on iPad</span></span>
- <span data-ttu-id="f451e-118">Word on Windows (Word 2013 以降)</span><span class="sxs-lookup"><span data-stu-id="f451e-118">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="f451e-119">Word on Mac (Word 2016 以降)</span><span class="sxs-lookup"><span data-stu-id="f451e-119">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="f451e-120">Word on the web</span><span class="sxs-lookup"><span data-stu-id="f451e-120">Word on the web</span></span>
- <span data-ttu-id="f451e-121">Word on iPad</span><span class="sxs-lookup"><span data-stu-id="f451e-121">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="f451e-122">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="f451e-122">ImageCoercion 1.2</span></span>

<span data-ttu-id="f451e-123">ImageCoercion 1.2 では、メソッドを使用してデータを書き込むときに SVG 形式 ( `Office.CoercionType.XmlSvg` ) に変換 [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) できます。</span><span class="sxs-lookup"><span data-stu-id="f451e-123">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="f451e-124">次のアプリケーションがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="f451e-124">The following applications are supported.</span></span>

- <span data-ttu-id="f451e-125">ExcelオンWindows (サブスクリプションに接続Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="f451e-125">Excel on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="f451e-126">Excel Mac (サブスクリプションに接続Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="f451e-126">Excel on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="f451e-127">PowerPoint (WindowsサブスクリプションにMicrosoft 365)</span><span class="sxs-lookup"><span data-stu-id="f451e-127">PowerPoint on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="f451e-128">PowerPoint (サブスクリプションにMicrosoft 365)</span><span class="sxs-lookup"><span data-stu-id="f451e-128">PowerPoint on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="f451e-129">PowerPoint on the web</span><span class="sxs-lookup"><span data-stu-id="f451e-129">PowerPoint on the web</span></span>
- <span data-ttu-id="f451e-130">Word on Windows (サブスクリプションにMicrosoft 365)</span><span class="sxs-lookup"><span data-stu-id="f451e-130">Word on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="f451e-131">Mac 上の Word (サブスクリプションに接続Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="f451e-131">Word on Mac (connected to a Microsoft 365 subscription)</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="f451e-132">Office 共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="f451e-132">Office Common API requirement sets</span></span>

<span data-ttu-id="f451e-133">共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="f451e-133">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="f451e-134">関連項目</span><span class="sxs-lookup"><span data-stu-id="f451e-134">See also</span></span>

- [<span data-ttu-id="f451e-135">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="f451e-135">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="f451e-136">Office アプリケーションと API 要件を指定する</span><span class="sxs-lookup"><span data-stu-id="f451e-136">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="f451e-137">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="f451e-137">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
