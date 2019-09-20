---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、OneNote、Outlook、PowerPoint、Project、Word のサポートされる要件セット。
ms.date: 08/13/2019
localization_priority: Priority
ms.openlocfilehash: a3c580f32ad7cd384309a9b53e55ea488a470a90
ms.sourcegitcommit: f781d7cfd980cd866d6d1d00c5b9d16c8a4b7f9b
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/20/2019
ms.locfileid: "37053327"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="62c0f-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="62c0f-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="62c0f-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="62c0f-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="62c0f-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="62c0f-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="62c0f-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="62c0f-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="62c0f-108">Excel</span><span class="sxs-lookup"><span data-stu-id="62c0f-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="62c0f-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="62c0f-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="62c0f-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="62c0f-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="62c0f-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="62c0f-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="62c0f-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="62c0f-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-113">Office on the web</span><span class="sxs-lookup"><span data-stu-id="62c0f-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="62c0f-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-114">- TaskPane</span></span><br><span data-ttu-id="62c0f-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="62c0f-115">
        - Content</span></span><br><span data-ttu-id="62c0f-116">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="62c0f-116">
        - Custom Functions</span></span><br><span data-ttu-id="62c0f-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="62c0f-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="62c0f-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="62c0f-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="62c0f-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="62c0f-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="62c0f-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="62c0f-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="62c0f-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="62c0f-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="62c0f-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="62c0f-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="62c0f-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-128">
        - BindingEvents</span></span><br><span data-ttu-id="62c0f-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-129">
        - CompressedFile</span></span><br><span data-ttu-id="62c0f-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-130">
        - DocumentEvents</span></span><br><span data-ttu-id="62c0f-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-131">
        - File</span></span><br><span data-ttu-id="62c0f-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-132">
        - MatrixBindings</span></span><br><span data-ttu-id="62c0f-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="62c0f-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-134">
        - Selection</span></span><br><span data-ttu-id="62c0f-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-135">
        - Settings</span></span><br><span data-ttu-id="62c0f-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-136">
        - TableBindings</span></span><br><span data-ttu-id="62c0f-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-137">
        - TableCoercion</span></span><br><span data-ttu-id="62c0f-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-138">
        - TextBindings</span></span><br><span data-ttu-id="62c0f-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-140">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="62c0f-140">Office on Windows</span></span><br><span data-ttu-id="62c0f-141">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="62c0f-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="62c0f-142">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-142">- TaskPane</span></span><br><span data-ttu-id="62c0f-143">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="62c0f-143">
        - Content</span></span><br><span data-ttu-id="62c0f-144">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="62c0f-144">
        - Custom Functions</span></span><br><span data-ttu-id="62c0f-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="62c0f-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="62c0f-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="62c0f-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="62c0f-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="62c0f-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="62c0f-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="62c0f-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="62c0f-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="62c0f-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="62c0f-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="62c0f-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="62c0f-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="62c0f-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="62c0f-158">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-158">
        - BindingEvents</span></span><br><span data-ttu-id="62c0f-159">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-159">
        - CompressedFile</span></span><br><span data-ttu-id="62c0f-160">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-160">
        - DocumentEvents</span></span><br><span data-ttu-id="62c0f-161">
        - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-161">
        - File</span></span><br><span data-ttu-id="62c0f-162">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-162">
        - MatrixBindings</span></span><br><span data-ttu-id="62c0f-163">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-163">
        - MatrixCoercion</span></span><br><span data-ttu-id="62c0f-164">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-164">
        - Selection</span></span><br><span data-ttu-id="62c0f-165">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-165">
        - Settings</span></span><br><span data-ttu-id="62c0f-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-166">
        - TableBindings</span></span><br><span data-ttu-id="62c0f-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-167">
        - TableCoercion</span></span><br><span data-ttu-id="62c0f-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-168">
        - TextBindings</span></span><br><span data-ttu-id="62c0f-169">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-169">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-170">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="62c0f-170">Office 2019 on Windows</span></span><br><span data-ttu-id="62c0f-171">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-171">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="62c0f-172">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-172">- TaskPane</span></span><br><span data-ttu-id="62c0f-173">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="62c0f-173">
        - Content</span></span><br><span data-ttu-id="62c0f-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="62c0f-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="62c0f-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="62c0f-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="62c0f-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="62c0f-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="62c0f-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="62c0f-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="62c0f-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="62c0f-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="62c0f-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="62c0f-185">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-185">- BindingEvents</span></span><br><span data-ttu-id="62c0f-186">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-186">
        - CompressedFile</span></span><br><span data-ttu-id="62c0f-187">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-187">
        - DocumentEvents</span></span><br><span data-ttu-id="62c0f-188">
        - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-188">
        - File</span></span><br><span data-ttu-id="62c0f-189">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-189">
        - MatrixBindings</span></span><br><span data-ttu-id="62c0f-190">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-190">
        - MatrixCoercion</span></span><br><span data-ttu-id="62c0f-191">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-191">
        - Selection</span></span><br><span data-ttu-id="62c0f-192">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-192">
        - Settings</span></span><br><span data-ttu-id="62c0f-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-193">
        - TableBindings</span></span><br><span data-ttu-id="62c0f-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-194">
        - TableCoercion</span></span><br><span data-ttu-id="62c0f-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-195">
        - TextBindings</span></span><br><span data-ttu-id="62c0f-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-196">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-197">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="62c0f-197">Office 2016 on Windows</span></span><br><span data-ttu-id="62c0f-198">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-198">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="62c0f-199">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-199">- TaskPane</span></span><br><span data-ttu-id="62c0f-200">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="62c0f-200">
        - Content</span></span></td>
    <td><span data-ttu-id="62c0f-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="62c0f-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="62c0f-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="62c0f-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="62c0f-204">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-204">- BindingEvents</span></span><br><span data-ttu-id="62c0f-205">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-205">
        - CompressedFile</span></span><br><span data-ttu-id="62c0f-206">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-206">
        - DocumentEvents</span></span><br><span data-ttu-id="62c0f-207">
        - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-207">
        - File</span></span><br><span data-ttu-id="62c0f-208">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-208">
        - MatrixBindings</span></span><br><span data-ttu-id="62c0f-209">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-209">
        - MatrixCoercion</span></span><br><span data-ttu-id="62c0f-210">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-210">
        - Selection</span></span><br><span data-ttu-id="62c0f-211">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-211">
        - Settings</span></span><br><span data-ttu-id="62c0f-212">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-212">
        - TableBindings</span></span><br><span data-ttu-id="62c0f-213">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-213">
        - TableCoercion</span></span><br><span data-ttu-id="62c0f-214">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-214">
        - TextBindings</span></span><br><span data-ttu-id="62c0f-215">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-215">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-216">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="62c0f-216">Office 2013 on Windows</span></span><br><span data-ttu-id="62c0f-217">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-217">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="62c0f-218">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-218">
        - TaskPane</span></span><br><span data-ttu-id="62c0f-219">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="62c0f-219">
        - Content</span></span></td>
    <td>  <span data-ttu-id="62c0f-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="62c0f-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="62c0f-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="62c0f-222">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-222">
        - BindingEvents</span></span><br><span data-ttu-id="62c0f-223">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-223">
        - CompressedFile</span></span><br><span data-ttu-id="62c0f-224">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-224">
        - DocumentEvents</span></span><br><span data-ttu-id="62c0f-225">
        - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-225">
        - File</span></span><br><span data-ttu-id="62c0f-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-226">
        - MatrixBindings</span></span><br><span data-ttu-id="62c0f-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-227">
        - MatrixCoercion</span></span><br><span data-ttu-id="62c0f-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-228">
        - Selection</span></span><br><span data-ttu-id="62c0f-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-229">
        - Settings</span></span><br><span data-ttu-id="62c0f-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-230">
        - TableBindings</span></span><br><span data-ttu-id="62c0f-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-231">
        - TableCoercion</span></span><br><span data-ttu-id="62c0f-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-232">
        - TextBindings</span></span><br><span data-ttu-id="62c0f-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-233">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-234">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="62c0f-234">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="62c0f-235">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="62c0f-235">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="62c0f-236">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-236">- TaskPane</span></span><br><span data-ttu-id="62c0f-237">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="62c0f-237">
        - Content</span></span></td>
    <td><span data-ttu-id="62c0f-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="62c0f-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="62c0f-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="62c0f-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="62c0f-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="62c0f-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="62c0f-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="62c0f-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="62c0f-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="62c0f-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="62c0f-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="62c0f-249">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-249">- BindingEvents</span></span><br><span data-ttu-id="62c0f-250">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-250">
        - DocumentEvents</span></span><br><span data-ttu-id="62c0f-251">
        - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-251">
        - File</span></span><br><span data-ttu-id="62c0f-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-252">
        - MatrixBindings</span></span><br><span data-ttu-id="62c0f-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-253">
        - MatrixCoercion</span></span><br><span data-ttu-id="62c0f-254">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-254">
        - Selection</span></span><br><span data-ttu-id="62c0f-255">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-255">
        - Settings</span></span><br><span data-ttu-id="62c0f-256">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-256">
        - TableBindings</span></span><br><span data-ttu-id="62c0f-257">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-257">
        - TableCoercion</span></span><br><span data-ttu-id="62c0f-258">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-258">
        - TextBindings</span></span><br><span data-ttu-id="62c0f-259">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-259">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-260">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="62c0f-260">Office apps on Mac</span></span><br><span data-ttu-id="62c0f-261">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="62c0f-261">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="62c0f-262">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-262">- TaskPane</span></span><br><span data-ttu-id="62c0f-263">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="62c0f-263">
        - Content</span></span><br><span data-ttu-id="62c0f-264">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="62c0f-264">
        - Custom Functions</span></span><br><span data-ttu-id="62c0f-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="62c0f-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="62c0f-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="62c0f-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="62c0f-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="62c0f-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="62c0f-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="62c0f-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="62c0f-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="62c0f-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="62c0f-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="62c0f-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="62c0f-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="62c0f-278">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-278">- BindingEvents</span></span><br><span data-ttu-id="62c0f-279">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-279">
        - CompressedFile</span></span><br><span data-ttu-id="62c0f-280">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-280">
        - DocumentEvents</span></span><br><span data-ttu-id="62c0f-281">
        - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-281">
        - File</span></span><br><span data-ttu-id="62c0f-282">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-282">
        - MatrixBindings</span></span><br><span data-ttu-id="62c0f-283">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-283">
        - MatrixCoercion</span></span><br><span data-ttu-id="62c0f-284">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-284">
        - PdfFile</span></span><br><span data-ttu-id="62c0f-285">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-285">
        - Selection</span></span><br><span data-ttu-id="62c0f-286">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-286">
        - Settings</span></span><br><span data-ttu-id="62c0f-287">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-287">
        - TableBindings</span></span><br><span data-ttu-id="62c0f-288">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-288">
        - TableCoercion</span></span><br><span data-ttu-id="62c0f-289">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-289">
        - TextBindings</span></span><br><span data-ttu-id="62c0f-290">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-290">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-291">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="62c0f-291">Office 2019 for Mac</span></span><br><span data-ttu-id="62c0f-292">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-292">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="62c0f-293">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-293">- TaskPane</span></span><br><span data-ttu-id="62c0f-294">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="62c0f-294">
        - Content</span></span><br><span data-ttu-id="62c0f-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="62c0f-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="62c0f-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="62c0f-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="62c0f-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="62c0f-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="62c0f-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="62c0f-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="62c0f-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="62c0f-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="62c0f-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="62c0f-306">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-306">- BindingEvents</span></span><br><span data-ttu-id="62c0f-307">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-307">
        - CompressedFile</span></span><br><span data-ttu-id="62c0f-308">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-308">
        - DocumentEvents</span></span><br><span data-ttu-id="62c0f-309">
        - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-309">
        - File</span></span><br><span data-ttu-id="62c0f-310">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-310">
        - MatrixBindings</span></span><br><span data-ttu-id="62c0f-311">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-311">
        - MatrixCoercion</span></span><br><span data-ttu-id="62c0f-312">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-312">
        - PdfFile</span></span><br><span data-ttu-id="62c0f-313">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-313">
        - Selection</span></span><br><span data-ttu-id="62c0f-314">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-314">
        - Settings</span></span><br><span data-ttu-id="62c0f-315">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-315">
        - TableBindings</span></span><br><span data-ttu-id="62c0f-316">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-316">
        - TableCoercion</span></span><br><span data-ttu-id="62c0f-317">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-317">
        - TextBindings</span></span><br><span data-ttu-id="62c0f-318">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-318">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-319">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="62c0f-319">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="62c0f-320">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-320">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="62c0f-321">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-321">- TaskPane</span></span><br><span data-ttu-id="62c0f-322">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="62c0f-322">
        - Content</span></span></td>
    <td><span data-ttu-id="62c0f-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="62c0f-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="62c0f-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="62c0f-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="62c0f-326">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-326">- BindingEvents</span></span><br><span data-ttu-id="62c0f-327">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-327">
        - CompressedFile</span></span><br><span data-ttu-id="62c0f-328">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-328">
        - DocumentEvents</span></span><br><span data-ttu-id="62c0f-329">
        - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-329">
        - File</span></span><br><span data-ttu-id="62c0f-330">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-330">
        - MatrixBindings</span></span><br><span data-ttu-id="62c0f-331">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-331">
        - MatrixCoercion</span></span><br><span data-ttu-id="62c0f-332">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-332">
        - PdfFile</span></span><br><span data-ttu-id="62c0f-333">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-333">
        - Selection</span></span><br><span data-ttu-id="62c0f-334">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-334">
        - Settings</span></span><br><span data-ttu-id="62c0f-335">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-335">
        - TableBindings</span></span><br><span data-ttu-id="62c0f-336">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-336">
        - TableCoercion</span></span><br><span data-ttu-id="62c0f-337">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-337">
        - TextBindings</span></span><br><span data-ttu-id="62c0f-338">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-338">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="62c0f-339">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="62c0f-339">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="62c0f-340">カスタム関数</span><span class="sxs-lookup"><span data-stu-id="62c0f-340">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="62c0f-341">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="62c0f-341">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="62c0f-342">拡張点</span><span class="sxs-lookup"><span data-stu-id="62c0f-342">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="62c0f-343">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="62c0f-343">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="62c0f-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="62c0f-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-345">Office on the web</span><span class="sxs-lookup"><span data-stu-id="62c0f-345">Office on the web</span></span></td>
    <td><span data-ttu-id="62c0f-346">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="62c0f-346">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="62c0f-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-348">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="62c0f-348">Office on Windows</span></span><br><span data-ttu-id="62c0f-349">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="62c0f-349">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="62c0f-350">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="62c0f-350">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="62c0f-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-352">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="62c0f-352">Office for Mac</span></span><br><span data-ttu-id="62c0f-353">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="62c0f-353">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="62c0f-354">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="62c0f-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="62c0f-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="62c0f-356">Outlook</span><span class="sxs-lookup"><span data-stu-id="62c0f-356">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="62c0f-357">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="62c0f-357">Platform</span></span></th>
    <th><span data-ttu-id="62c0f-358">拡張点</span><span class="sxs-lookup"><span data-stu-id="62c0f-358">Extension points</span></span></th>
    <th><span data-ttu-id="62c0f-359">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="62c0f-359">API requirement sets</span></span></th>
    <th><span data-ttu-id="62c0f-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="62c0f-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-361">Office on the web</span><span class="sxs-lookup"><span data-stu-id="62c0f-361">Office on the web</span></span><br><span data-ttu-id="62c0f-362">(モダン)</span><span class="sxs-lookup"><span data-stu-id="62c0f-362">Modern</span></span></td>
    <td> <span data-ttu-id="62c0f-363">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="62c0f-363">- Mail Read</span></span><br><span data-ttu-id="62c0f-364">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="62c0f-364">
      - Mail Compose</span></span><br><span data-ttu-id="62c0f-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="62c0f-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="62c0f-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="62c0f-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="62c0f-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="62c0f-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="62c0f-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="62c0f-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="62c0f-373">使用不可</span><span class="sxs-lookup"><span data-stu-id="62c0f-373">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-374">Office on the web</span><span class="sxs-lookup"><span data-stu-id="62c0f-374">Office on the web</span></span><br><span data-ttu-id="62c0f-375">(クラシック)</span><span class="sxs-lookup"><span data-stu-id="62c0f-375">Classic</span></span></td>
    <td> <span data-ttu-id="62c0f-376">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="62c0f-376">- Mail Read</span></span><br><span data-ttu-id="62c0f-377">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="62c0f-377">
      - Mail Compose</span></span><br><span data-ttu-id="62c0f-378">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-378">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="62c0f-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="62c0f-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="62c0f-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="62c0f-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="62c0f-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="62c0f-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="62c0f-385">使用不可</span><span class="sxs-lookup"><span data-stu-id="62c0f-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-386">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="62c0f-386">Office on Windows</span></span><br><span data-ttu-id="62c0f-387">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="62c0f-387">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="62c0f-388">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="62c0f-388">- Mail Read</span></span><br><span data-ttu-id="62c0f-389">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="62c0f-389">
      - Mail Compose</span></span><br><span data-ttu-id="62c0f-390">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-390">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="62c0f-391">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="62c0f-391">
      - Modules</span></span></td>
    <td> <span data-ttu-id="62c0f-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="62c0f-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="62c0f-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="62c0f-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="62c0f-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="62c0f-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="62c0f-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="62c0f-399">使用不可</span><span class="sxs-lookup"><span data-stu-id="62c0f-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-400">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="62c0f-400">Office 2019 on Windows</span></span><br><span data-ttu-id="62c0f-401">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-401">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="62c0f-402">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="62c0f-402">- Mail Read</span></span><br><span data-ttu-id="62c0f-403">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="62c0f-403">
      - Mail Compose</span></span><br><span data-ttu-id="62c0f-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="62c0f-405">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="62c0f-405">
      - Modules</span></span></td>
    <td> <span data-ttu-id="62c0f-406">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-406">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="62c0f-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="62c0f-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="62c0f-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="62c0f-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="62c0f-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="62c0f-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="62c0f-413">使用不可</span><span class="sxs-lookup"><span data-stu-id="62c0f-413">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-414">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="62c0f-414">Office 2016 on Windows</span></span><br><span data-ttu-id="62c0f-415">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-415">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="62c0f-416">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="62c0f-416">- Mail Read</span></span><br><span data-ttu-id="62c0f-417">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="62c0f-417">
      - Mail Compose</span></span><br><span data-ttu-id="62c0f-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="62c0f-419">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="62c0f-419">
      - Modules</span></span></td>
    <td> <span data-ttu-id="62c0f-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="62c0f-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="62c0f-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="62c0f-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="62c0f-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="62c0f-424">使用不可</span><span class="sxs-lookup"><span data-stu-id="62c0f-424">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-425">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="62c0f-425">Office 2013 on Windows</span></span><br><span data-ttu-id="62c0f-426">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-426">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="62c0f-427">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="62c0f-427">- Mail Read</span></span><br><span data-ttu-id="62c0f-428">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="62c0f-428">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="62c0f-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="62c0f-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="62c0f-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="62c0f-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="62c0f-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="62c0f-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="62c0f-433">使用不可</span><span class="sxs-lookup"><span data-stu-id="62c0f-433">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-434">iOS 上の Office</span><span class="sxs-lookup"><span data-stu-id="62c0f-434">Office apps on iOS</span></span><br><span data-ttu-id="62c0f-435">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="62c0f-435">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="62c0f-436">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="62c0f-436">- Mail Read</span></span><br><span data-ttu-id="62c0f-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="62c0f-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="62c0f-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="62c0f-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="62c0f-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="62c0f-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="62c0f-443">使用不可</span><span class="sxs-lookup"><span data-stu-id="62c0f-443">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-444">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="62c0f-444">Office apps on Mac</span></span><br><span data-ttu-id="62c0f-445">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="62c0f-445">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="62c0f-446">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="62c0f-446">- Mail Read</span></span><br><span data-ttu-id="62c0f-447">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="62c0f-447">
      - Mail Compose</span></span><br><span data-ttu-id="62c0f-448">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-448">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="62c0f-449">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-449">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="62c0f-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="62c0f-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="62c0f-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="62c0f-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="62c0f-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="62c0f-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="62c0f-456">使用不可</span><span class="sxs-lookup"><span data-stu-id="62c0f-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-457">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="62c0f-457">Office 2019 for Mac</span></span><br><span data-ttu-id="62c0f-458">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-458">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="62c0f-459">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="62c0f-459">- Mail Read</span></span><br><span data-ttu-id="62c0f-460">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="62c0f-460">
      - Mail Compose</span></span><br><span data-ttu-id="62c0f-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="62c0f-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="62c0f-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="62c0f-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="62c0f-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="62c0f-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="62c0f-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="62c0f-468">使用不可</span><span class="sxs-lookup"><span data-stu-id="62c0f-468">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-469">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="62c0f-469">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="62c0f-470">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-470">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="62c0f-471">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="62c0f-471">- Mail Read</span></span><br><span data-ttu-id="62c0f-472">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="62c0f-472">
      - Mail Compose</span></span><br><span data-ttu-id="62c0f-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="62c0f-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="62c0f-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="62c0f-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="62c0f-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="62c0f-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="62c0f-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="62c0f-480">使用不可</span><span class="sxs-lookup"><span data-stu-id="62c0f-480">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-481">Android 上の Office</span><span class="sxs-lookup"><span data-stu-id="62c0f-481">Office apps on Android</span></span><br><span data-ttu-id="62c0f-482">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="62c0f-482">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="62c0f-483">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="62c0f-483">- Mail Read</span></span><br><span data-ttu-id="62c0f-484">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-484">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="62c0f-485">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-485">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="62c0f-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="62c0f-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="62c0f-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="62c0f-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="62c0f-490">利用不可</span><span class="sxs-lookup"><span data-stu-id="62c0f-490">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="62c0f-491">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="62c0f-491">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="62c0f-492">Word</span><span class="sxs-lookup"><span data-stu-id="62c0f-492">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="62c0f-493">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="62c0f-493">Platform</span></span></th>
    <th><span data-ttu-id="62c0f-494">拡張点</span><span class="sxs-lookup"><span data-stu-id="62c0f-494">Extension points</span></span></th>
    <th><span data-ttu-id="62c0f-495">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="62c0f-495">API requirement sets</span></span></th>
    <th><span data-ttu-id="62c0f-496"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="62c0f-496"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-497">Office on the web</span><span class="sxs-lookup"><span data-stu-id="62c0f-497">Office on the web</span></span></td>
    <td> <span data-ttu-id="62c0f-498">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-498">- TaskPane</span></span><br><span data-ttu-id="62c0f-499">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-499">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="62c0f-500">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-500">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="62c0f-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="62c0f-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="62c0f-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="62c0f-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="62c0f-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="62c0f-506">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-506">- BindingEvents</span></span><br><span data-ttu-id="62c0f-507">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="62c0f-507">
         - CustomXmlParts</span></span><br><span data-ttu-id="62c0f-508">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-508">
         - DocumentEvents</span></span><br><span data-ttu-id="62c0f-509">
         - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-509">
         - File</span></span><br><span data-ttu-id="62c0f-510">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-510">
         - HtmlCoercion</span></span><br><span data-ttu-id="62c0f-511">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-511">
         - MatrixBindings</span></span><br><span data-ttu-id="62c0f-512">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-512">
         - MatrixCoercion</span></span><br><span data-ttu-id="62c0f-513">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-513">
         - OoxmlCoercion</span></span><br><span data-ttu-id="62c0f-514">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-514">
         - PdfFile</span></span><br><span data-ttu-id="62c0f-515">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-515">
         - Selection</span></span><br><span data-ttu-id="62c0f-516">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-516">
         - Settings</span></span><br><span data-ttu-id="62c0f-517">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-517">
         - TableBindings</span></span><br><span data-ttu-id="62c0f-518">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-518">
         - TableCoercion</span></span><br><span data-ttu-id="62c0f-519">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-519">
         - TextBindings</span></span><br><span data-ttu-id="62c0f-520">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-520">
         - TextCoercion</span></span><br><span data-ttu-id="62c0f-521">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-521">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-522">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="62c0f-522">Office on Windows</span></span><br><span data-ttu-id="62c0f-523">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="62c0f-523">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="62c0f-524">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-524">- TaskPane</span></span><br><span data-ttu-id="62c0f-525">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-525">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="62c0f-526">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-526">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="62c0f-527">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-527">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="62c0f-528">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-528">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="62c0f-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="62c0f-530">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-530">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="62c0f-531">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-531">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="62c0f-532">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-532">- BindingEvents</span></span><br><span data-ttu-id="62c0f-533">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-533">
         - CompressedFile</span></span><br><span data-ttu-id="62c0f-534">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="62c0f-534">
         - CustomXmlParts</span></span><br><span data-ttu-id="62c0f-535">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-535">
         - DocumentEvents</span></span><br><span data-ttu-id="62c0f-536">
         - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-536">
         - File</span></span><br><span data-ttu-id="62c0f-537">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-537">
         - HtmlCoercion</span></span><br><span data-ttu-id="62c0f-538">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-538">
         - MatrixBindings</span></span><br><span data-ttu-id="62c0f-539">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-539">
         - MatrixCoercion</span></span><br><span data-ttu-id="62c0f-540">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-540">
         - OoxmlCoercion</span></span><br><span data-ttu-id="62c0f-541">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-541">
         - PdfFile</span></span><br><span data-ttu-id="62c0f-542">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-542">
         - Selection</span></span><br><span data-ttu-id="62c0f-543">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-543">
         - Settings</span></span><br><span data-ttu-id="62c0f-544">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-544">
         - TableBindings</span></span><br><span data-ttu-id="62c0f-545">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-545">
         - TableCoercion</span></span><br><span data-ttu-id="62c0f-546">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-546">
         - TextBindings</span></span><br><span data-ttu-id="62c0f-547">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-547">
         - TextCoercion</span></span><br><span data-ttu-id="62c0f-548">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-548">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-549">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="62c0f-549">Office 2019 on Windows</span></span><br><span data-ttu-id="62c0f-550">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-550">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="62c0f-551">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-551">- TaskPane</span></span><br><span data-ttu-id="62c0f-552">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-552">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="62c0f-553">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-553">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="62c0f-554">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-554">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="62c0f-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="62c0f-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="62c0f-557">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-557">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="62c0f-558">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-558">- BindingEvents</span></span><br><span data-ttu-id="62c0f-559">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-559">
         - CompressedFile</span></span><br><span data-ttu-id="62c0f-560">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="62c0f-560">
         - CustomXmlParts</span></span><br><span data-ttu-id="62c0f-561">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-561">
         - DocumentEvents</span></span><br><span data-ttu-id="62c0f-562">
         - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-562">
         - File</span></span><br><span data-ttu-id="62c0f-563">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-563">
         - HtmlCoercion</span></span><br><span data-ttu-id="62c0f-564">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-564">
         - MatrixBindings</span></span><br><span data-ttu-id="62c0f-565">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-565">
         - MatrixCoercion</span></span><br><span data-ttu-id="62c0f-566">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-566">
         - OoxmlCoercion</span></span><br><span data-ttu-id="62c0f-567">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-567">
         - PdfFile</span></span><br><span data-ttu-id="62c0f-568">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-568">
         - Selection</span></span><br><span data-ttu-id="62c0f-569">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-569">
         - Settings</span></span><br><span data-ttu-id="62c0f-570">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-570">
         - TableBindings</span></span><br><span data-ttu-id="62c0f-571">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-571">
         - TableCoercion</span></span><br><span data-ttu-id="62c0f-572">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-572">
         - TextBindings</span></span><br><span data-ttu-id="62c0f-573">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-573">
         - TextCoercion</span></span><br><span data-ttu-id="62c0f-574">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-574">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-575">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="62c0f-575">Office 2016 on Windows</span></span><br><span data-ttu-id="62c0f-576">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-576">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="62c0f-577">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-577">- TaskPane</span></span></td>
    <td> <span data-ttu-id="62c0f-578">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-578">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="62c0f-579">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="62c0f-579">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="62c0f-580">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-580">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="62c0f-581">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-581">- BindingEvents</span></span><br><span data-ttu-id="62c0f-582">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-582">
         - CompressedFile</span></span><br><span data-ttu-id="62c0f-583">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="62c0f-583">
         - CustomXmlParts</span></span><br><span data-ttu-id="62c0f-584">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-584">
         - DocumentEvents</span></span><br><span data-ttu-id="62c0f-585">
         - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-585">
         - File</span></span><br><span data-ttu-id="62c0f-586">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-586">
         - HtmlCoercion</span></span><br><span data-ttu-id="62c0f-587">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-587">
         - MatrixBindings</span></span><br><span data-ttu-id="62c0f-588">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-588">
         - MatrixCoercion</span></span><br><span data-ttu-id="62c0f-589">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-589">
         - OoxmlCoercion</span></span><br><span data-ttu-id="62c0f-590">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-590">
         - PdfFile</span></span><br><span data-ttu-id="62c0f-591">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-591">
         - Selection</span></span><br><span data-ttu-id="62c0f-592">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-592">
         - Settings</span></span><br><span data-ttu-id="62c0f-593">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-593">
         - TableBindings</span></span><br><span data-ttu-id="62c0f-594">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-594">
         - TableCoercion</span></span><br><span data-ttu-id="62c0f-595">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-595">
         - TextBindings</span></span><br><span data-ttu-id="62c0f-596">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-596">
         - TextCoercion</span></span><br><span data-ttu-id="62c0f-597">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-597">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-598">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="62c0f-598">Office 2013 on Windows</span></span><br><span data-ttu-id="62c0f-599">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-599">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="62c0f-600">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-600">- TaskPane</span></span></td>
    <td> <span data-ttu-id="62c0f-601">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="62c0f-601">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="62c0f-602">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-602">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="62c0f-603">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-603">- BindingEvents</span></span><br><span data-ttu-id="62c0f-604">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-604">
         - CompressedFile</span></span><br><span data-ttu-id="62c0f-605">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="62c0f-605">
         - CustomXmlParts</span></span><br><span data-ttu-id="62c0f-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-606">
         - DocumentEvents</span></span><br><span data-ttu-id="62c0f-607">
         - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-607">
         - File</span></span><br><span data-ttu-id="62c0f-608">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-608">
         - HtmlCoercion</span></span><br><span data-ttu-id="62c0f-609">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-609">
         - MatrixBindings</span></span><br><span data-ttu-id="62c0f-610">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-610">
         - MatrixCoercion</span></span><br><span data-ttu-id="62c0f-611">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-611">
         - OoxmlCoercion</span></span><br><span data-ttu-id="62c0f-612">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-612">
         - PdfFile</span></span><br><span data-ttu-id="62c0f-613">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-613">
         - Selection</span></span><br><span data-ttu-id="62c0f-614">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-614">
         - Settings</span></span><br><span data-ttu-id="62c0f-615">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-615">
         - TableBindings</span></span><br><span data-ttu-id="62c0f-616">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-616">
         - TableCoercion</span></span><br><span data-ttu-id="62c0f-617">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-617">
         - TextBindings</span></span><br><span data-ttu-id="62c0f-618">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-618">
         - TextCoercion</span></span><br><span data-ttu-id="62c0f-619">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-619">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-620">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="62c0f-620">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="62c0f-621">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="62c0f-621">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="62c0f-622">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-622">- TaskPane</span></span></td>
    <td> <span data-ttu-id="62c0f-623">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-623">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="62c0f-624">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-624">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="62c0f-625">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-625">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="62c0f-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="62c0f-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="62c0f-628">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-628">- BindingEvents</span></span><br><span data-ttu-id="62c0f-629">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-629">
         - CompressedFile</span></span><br><span data-ttu-id="62c0f-630">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="62c0f-630">
         - CustomXmlParts</span></span><br><span data-ttu-id="62c0f-631">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-631">
         - DocumentEvents</span></span><br><span data-ttu-id="62c0f-632">
         - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-632">
         - File</span></span><br><span data-ttu-id="62c0f-633">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-633">
         - HtmlCoercion</span></span><br><span data-ttu-id="62c0f-634">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-634">
         - MatrixBindings</span></span><br><span data-ttu-id="62c0f-635">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-635">
         - MatrixCoercion</span></span><br><span data-ttu-id="62c0f-636">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-636">
         - OoxmlCoercion</span></span><br><span data-ttu-id="62c0f-637">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-637">
         - PdfFile</span></span><br><span data-ttu-id="62c0f-638">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-638">
         - Selection</span></span><br><span data-ttu-id="62c0f-639">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-639">
         - Settings</span></span><br><span data-ttu-id="62c0f-640">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-640">
         - TableBindings</span></span><br><span data-ttu-id="62c0f-641">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-641">
         - TableCoercion</span></span><br><span data-ttu-id="62c0f-642">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-642">
         - TextBindings</span></span><br><span data-ttu-id="62c0f-643">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-643">
         - TextCoercion</span></span><br><span data-ttu-id="62c0f-644">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-644">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-645">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="62c0f-645">Office apps on Mac</span></span><br><span data-ttu-id="62c0f-646">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="62c0f-646">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="62c0f-647">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-647">- TaskPane</span></span><br><span data-ttu-id="62c0f-648">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-648">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="62c0f-649">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-649">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="62c0f-650">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-650">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="62c0f-651">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-651">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="62c0f-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="62c0f-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="62c0f-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="62c0f-655">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-655">- BindingEvents</span></span><br><span data-ttu-id="62c0f-656">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-656">
         - CompressedFile</span></span><br><span data-ttu-id="62c0f-657">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="62c0f-657">
         - CustomXmlParts</span></span><br><span data-ttu-id="62c0f-658">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-658">
         - DocumentEvents</span></span><br><span data-ttu-id="62c0f-659">
         - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-659">
         - File</span></span><br><span data-ttu-id="62c0f-660">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-660">
         - HtmlCoercion</span></span><br><span data-ttu-id="62c0f-661">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-661">
         - MatrixBindings</span></span><br><span data-ttu-id="62c0f-662">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-662">
         - MatrixCoercion</span></span><br><span data-ttu-id="62c0f-663">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-663">
         - OoxmlCoercion</span></span><br><span data-ttu-id="62c0f-664">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-664">
         - PdfFile</span></span><br><span data-ttu-id="62c0f-665">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-665">
         - Selection</span></span><br><span data-ttu-id="62c0f-666">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-666">
         - Settings</span></span><br><span data-ttu-id="62c0f-667">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-667">
         - TableBindings</span></span><br><span data-ttu-id="62c0f-668">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-668">
         - TableCoercion</span></span><br><span data-ttu-id="62c0f-669">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-669">
         - TextBindings</span></span><br><span data-ttu-id="62c0f-670">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-670">
         - TextCoercion</span></span><br><span data-ttu-id="62c0f-671">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-671">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-672">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="62c0f-672">Office 2019 for Mac</span></span><br><span data-ttu-id="62c0f-673">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-673">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="62c0f-674">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-674">- TaskPane</span></span><br><span data-ttu-id="62c0f-675">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-675">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="62c0f-676">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-676">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="62c0f-677">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-677">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="62c0f-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="62c0f-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="62c0f-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="62c0f-681">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-681">- BindingEvents</span></span><br><span data-ttu-id="62c0f-682">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-682">
         - CompressedFile</span></span><br><span data-ttu-id="62c0f-683">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="62c0f-683">
         - CustomXmlParts</span></span><br><span data-ttu-id="62c0f-684">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-684">
         - DocumentEvents</span></span><br><span data-ttu-id="62c0f-685">
         - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-685">
         - File</span></span><br><span data-ttu-id="62c0f-686">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-686">
         - HtmlCoercion</span></span><br><span data-ttu-id="62c0f-687">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-687">
         - MatrixBindings</span></span><br><span data-ttu-id="62c0f-688">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-688">
         - MatrixCoercion</span></span><br><span data-ttu-id="62c0f-689">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-689">
         - OoxmlCoercion</span></span><br><span data-ttu-id="62c0f-690">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-690">
         - PdfFile</span></span><br><span data-ttu-id="62c0f-691">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-691">
         - Selection</span></span><br><span data-ttu-id="62c0f-692">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-692">
         - Settings</span></span><br><span data-ttu-id="62c0f-693">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-693">
         - TableBindings</span></span><br><span data-ttu-id="62c0f-694">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-694">
         - TableCoercion</span></span><br><span data-ttu-id="62c0f-695">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-695">
         - TextBindings</span></span><br><span data-ttu-id="62c0f-696">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-696">
         - TextCoercion</span></span><br><span data-ttu-id="62c0f-697">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-697">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-698">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="62c0f-698">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="62c0f-699">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-699">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="62c0f-700">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-700">- TaskPane</span></span></td>
    <td> <span data-ttu-id="62c0f-701">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-701">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="62c0f-702">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="62c0f-702">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="62c0f-703">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-703">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="62c0f-704">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-704">- BindingEvents</span></span><br><span data-ttu-id="62c0f-705">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-705">
         - CompressedFile</span></span><br><span data-ttu-id="62c0f-706">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="62c0f-706">
         - CustomXmlParts</span></span><br><span data-ttu-id="62c0f-707">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-707">
         - DocumentEvents</span></span><br><span data-ttu-id="62c0f-708">
         - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-708">
         - File</span></span><br><span data-ttu-id="62c0f-709">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-709">
         - HtmlCoercion</span></span><br><span data-ttu-id="62c0f-710">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-710">
         - MatrixBindings</span></span><br><span data-ttu-id="62c0f-711">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-711">
         - MatrixCoercion</span></span><br><span data-ttu-id="62c0f-712">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-712">
         - OoxmlCoercion</span></span><br><span data-ttu-id="62c0f-713">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-713">
         - PdfFile</span></span><br><span data-ttu-id="62c0f-714">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-714">
         - Selection</span></span><br><span data-ttu-id="62c0f-715">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-715">
         - Settings</span></span><br><span data-ttu-id="62c0f-716">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-716">
         - TableBindings</span></span><br><span data-ttu-id="62c0f-717">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-717">
         - TableCoercion</span></span><br><span data-ttu-id="62c0f-718">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="62c0f-718">
         - TextBindings</span></span><br><span data-ttu-id="62c0f-719">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-719">
         - TextCoercion</span></span><br><span data-ttu-id="62c0f-720">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-720">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="62c0f-721">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="62c0f-721">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="62c0f-722">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="62c0f-722">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="62c0f-723">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="62c0f-723">Platform</span></span></th>
    <th><span data-ttu-id="62c0f-724">拡張点</span><span class="sxs-lookup"><span data-stu-id="62c0f-724">Extension points</span></span></th>
    <th><span data-ttu-id="62c0f-725">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="62c0f-725">API requirement sets</span></span></th>
    <th><span data-ttu-id="62c0f-726"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="62c0f-726"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-727">Office on the web</span><span class="sxs-lookup"><span data-stu-id="62c0f-727">Office on the web</span></span></td>
    <td> <span data-ttu-id="62c0f-728">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="62c0f-728">- Content</span></span><br><span data-ttu-id="62c0f-729">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-729">
         - TaskPane</span></span><br><span data-ttu-id="62c0f-730">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-730">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="62c0f-731">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-731">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="62c0f-732">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-732">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="62c0f-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="62c0f-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="62c0f-735">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="62c0f-735">- ActiveView</span></span><br><span data-ttu-id="62c0f-736">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-736">
         - CompressedFile</span></span><br><span data-ttu-id="62c0f-737">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-737">
         - DocumentEvents</span></span><br><span data-ttu-id="62c0f-738">
         - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-738">
         - File</span></span><br><span data-ttu-id="62c0f-739">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-739">
         - PdfFile</span></span><br><span data-ttu-id="62c0f-740">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-740">
         - Selection</span></span><br><span data-ttu-id="62c0f-741">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-741">
         - Settings</span></span><br><span data-ttu-id="62c0f-742">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-742">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-743">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="62c0f-743">Office on Windows</span></span><br><span data-ttu-id="62c0f-744">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="62c0f-744">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="62c0f-745">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="62c0f-745">- Content</span></span><br><span data-ttu-id="62c0f-746">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-746">
         - TaskPane</span></span><br><span data-ttu-id="62c0f-747">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-747">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="62c0f-748">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-748">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="62c0f-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="62c0f-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="62c0f-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="62c0f-752">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="62c0f-752">- ActiveView</span></span><br><span data-ttu-id="62c0f-753">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-753">
         - CompressedFile</span></span><br><span data-ttu-id="62c0f-754">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-754">
         - DocumentEvents</span></span><br><span data-ttu-id="62c0f-755">
         - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-755">
         - File</span></span><br><span data-ttu-id="62c0f-756">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-756">
         - PdfFile</span></span><br><span data-ttu-id="62c0f-757">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-757">
         - Selection</span></span><br><span data-ttu-id="62c0f-758">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-758">
         - Settings</span></span><br><span data-ttu-id="62c0f-759">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-759">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-760">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="62c0f-760">Office 2019 on Windows</span></span><br><span data-ttu-id="62c0f-761">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-761">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="62c0f-762">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="62c0f-762">- Content</span></span><br><span data-ttu-id="62c0f-763">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-763">
         - TaskPane</span></span><br><span data-ttu-id="62c0f-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="62c0f-765">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-765">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="62c0f-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="62c0f-767">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="62c0f-767">- ActiveView</span></span><br><span data-ttu-id="62c0f-768">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-768">
         - CompressedFile</span></span><br><span data-ttu-id="62c0f-769">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-769">
         - DocumentEvents</span></span><br><span data-ttu-id="62c0f-770">
         - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-770">
         - File</span></span><br><span data-ttu-id="62c0f-771">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-771">
         - PdfFile</span></span><br><span data-ttu-id="62c0f-772">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-772">
         - Selection</span></span><br><span data-ttu-id="62c0f-773">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-773">
         - Settings</span></span><br><span data-ttu-id="62c0f-774">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-774">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-775">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="62c0f-775">Office 2016 on Windows</span></span><br><span data-ttu-id="62c0f-776">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-776">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="62c0f-777">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="62c0f-777">- Content</span></span><br><span data-ttu-id="62c0f-778">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-778">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="62c0f-779">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="62c0f-779">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="62c0f-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="62c0f-781">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="62c0f-781">- ActiveView</span></span><br><span data-ttu-id="62c0f-782">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-782">
         - CompressedFile</span></span><br><span data-ttu-id="62c0f-783">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-783">
         - DocumentEvents</span></span><br><span data-ttu-id="62c0f-784">
         - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-784">
         - File</span></span><br><span data-ttu-id="62c0f-785">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-785">
         - PdfFile</span></span><br><span data-ttu-id="62c0f-786">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-786">
         - Selection</span></span><br><span data-ttu-id="62c0f-787">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-787">
         - Settings</span></span><br><span data-ttu-id="62c0f-788">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-788">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-789">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="62c0f-789">Office 2013 on Windows</span></span><br><span data-ttu-id="62c0f-790">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-790">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="62c0f-791">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="62c0f-791">- Content</span></span><br><span data-ttu-id="62c0f-792">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-792">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="62c0f-793">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="62c0f-793">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="62c0f-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="62c0f-795">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="62c0f-795">- ActiveView</span></span><br><span data-ttu-id="62c0f-796">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-796">
         - CompressedFile</span></span><br><span data-ttu-id="62c0f-797">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-797">
         - DocumentEvents</span></span><br><span data-ttu-id="62c0f-798">
         - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-798">
         - File</span></span><br><span data-ttu-id="62c0f-799">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-799">
         - PdfFile</span></span><br><span data-ttu-id="62c0f-800">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-800">
         - Selection</span></span><br><span data-ttu-id="62c0f-801">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-801">
         - Settings</span></span><br><span data-ttu-id="62c0f-802">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-802">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-803">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="62c0f-803">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="62c0f-804">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="62c0f-804">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="62c0f-805">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="62c0f-805">- Content</span></span><br><span data-ttu-id="62c0f-806">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-806">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="62c0f-807">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-807">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="62c0f-808">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-808">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="62c0f-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="62c0f-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="62c0f-810">- ActiveView</span></span><br><span data-ttu-id="62c0f-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-811">
         - CompressedFile</span></span><br><span data-ttu-id="62c0f-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-812">
         - DocumentEvents</span></span><br><span data-ttu-id="62c0f-813">
         - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-813">
         - File</span></span><br><span data-ttu-id="62c0f-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-814">
         - PdfFile</span></span><br><span data-ttu-id="62c0f-815">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-815">
         - Selection</span></span><br><span data-ttu-id="62c0f-816">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-816">
         - Settings</span></span><br><span data-ttu-id="62c0f-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-818">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="62c0f-818">Office apps on Mac</span></span><br><span data-ttu-id="62c0f-819">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="62c0f-819">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="62c0f-820">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="62c0f-820">- Content</span></span><br><span data-ttu-id="62c0f-821">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-821">
         - TaskPane</span></span><br><span data-ttu-id="62c0f-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="62c0f-823">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-823">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="62c0f-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="62c0f-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="62c0f-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="62c0f-827">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="62c0f-827">- ActiveView</span></span><br><span data-ttu-id="62c0f-828">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-828">
         - CompressedFile</span></span><br><span data-ttu-id="62c0f-829">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-829">
         - DocumentEvents</span></span><br><span data-ttu-id="62c0f-830">
         - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-830">
         - File</span></span><br><span data-ttu-id="62c0f-831">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-831">
         - PdfFile</span></span><br><span data-ttu-id="62c0f-832">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-832">
         - Selection</span></span><br><span data-ttu-id="62c0f-833">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-833">
         - Settings</span></span><br><span data-ttu-id="62c0f-834">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-834">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-835">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="62c0f-835">Office 2019 for Mac</span></span><br><span data-ttu-id="62c0f-836">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-836">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="62c0f-837">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="62c0f-837">- Content</span></span><br><span data-ttu-id="62c0f-838">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-838">
         - TaskPane</span></span><br><span data-ttu-id="62c0f-839">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-839">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="62c0f-840">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-840">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="62c0f-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="62c0f-842">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="62c0f-842">- ActiveView</span></span><br><span data-ttu-id="62c0f-843">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-843">
         - CompressedFile</span></span><br><span data-ttu-id="62c0f-844">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-844">
         - DocumentEvents</span></span><br><span data-ttu-id="62c0f-845">
         - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-845">
         - File</span></span><br><span data-ttu-id="62c0f-846">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-846">
         - PdfFile</span></span><br><span data-ttu-id="62c0f-847">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-847">
         - Selection</span></span><br><span data-ttu-id="62c0f-848">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-848">
         - Settings</span></span><br><span data-ttu-id="62c0f-849">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-849">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-850">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="62c0f-850">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="62c0f-851">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-851">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="62c0f-852">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="62c0f-852">- Content</span></span><br><span data-ttu-id="62c0f-853">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-853">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="62c0f-854">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="62c0f-854">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="62c0f-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="62c0f-856">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="62c0f-856">- ActiveView</span></span><br><span data-ttu-id="62c0f-857">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-857">
         - CompressedFile</span></span><br><span data-ttu-id="62c0f-858">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-858">
         - DocumentEvents</span></span><br><span data-ttu-id="62c0f-859">
         - File</span><span class="sxs-lookup"><span data-stu-id="62c0f-859">
         - File</span></span><br><span data-ttu-id="62c0f-860">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="62c0f-860">
         - PdfFile</span></span><br><span data-ttu-id="62c0f-861">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-861">
         - Selection</span></span><br><span data-ttu-id="62c0f-862">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-862">
         - Settings</span></span><br><span data-ttu-id="62c0f-863">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-863">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="62c0f-864">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="62c0f-864">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="62c0f-865">OneNote</span><span class="sxs-lookup"><span data-stu-id="62c0f-865">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="62c0f-866">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="62c0f-866">Platform</span></span></th>
    <th><span data-ttu-id="62c0f-867">拡張点</span><span class="sxs-lookup"><span data-stu-id="62c0f-867">Extension points</span></span></th>
    <th><span data-ttu-id="62c0f-868">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="62c0f-868">API requirement sets</span></span></th>
    <th><span data-ttu-id="62c0f-869"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="62c0f-869"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-870">Office on the web</span><span class="sxs-lookup"><span data-stu-id="62c0f-870">Office on the web</span></span></td>
    <td> <span data-ttu-id="62c0f-871">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="62c0f-871">- Content</span></span><br><span data-ttu-id="62c0f-872">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-872">
         - TaskPane</span></span><br><span data-ttu-id="62c0f-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="62c0f-874">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-874">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="62c0f-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="62c0f-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="62c0f-877">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="62c0f-877">- DocumentEvents</span></span><br><span data-ttu-id="62c0f-878">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-878">
         - HtmlCoercion</span></span><br><span data-ttu-id="62c0f-879">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="62c0f-879">
         - Settings</span></span><br><span data-ttu-id="62c0f-880">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-880">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="62c0f-881">Project</span><span class="sxs-lookup"><span data-stu-id="62c0f-881">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="62c0f-882">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="62c0f-882">Platform</span></span></th>
    <th><span data-ttu-id="62c0f-883">拡張点</span><span class="sxs-lookup"><span data-stu-id="62c0f-883">Extension points</span></span></th>
    <th><span data-ttu-id="62c0f-884">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="62c0f-884">API requirement sets</span></span></th>
    <th><span data-ttu-id="62c0f-885"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="62c0f-885"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-886">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="62c0f-886">Office 2019 on Windows</span></span><br><span data-ttu-id="62c0f-887">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-887">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="62c0f-888">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-888">- TaskPane</span></span></td>
    <td> <span data-ttu-id="62c0f-889">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-889">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="62c0f-890">- Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-890">- Selection</span></span><br><span data-ttu-id="62c0f-891">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-891">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-892">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="62c0f-892">Office 2016 on Windows</span></span><br><span data-ttu-id="62c0f-893">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-893">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="62c0f-894">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-894">- TaskPane</span></span></td>
    <td> <span data-ttu-id="62c0f-895">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-895">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="62c0f-896">- Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-896">- Selection</span></span><br><span data-ttu-id="62c0f-897">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-897">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="62c0f-898">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="62c0f-898">Office 2013 on Windows</span></span><br><span data-ttu-id="62c0f-899">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="62c0f-899">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="62c0f-900">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="62c0f-900">- TaskPane</span></span></td>
    <td> <span data-ttu-id="62c0f-901">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="62c0f-901">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="62c0f-902">- Selection</span><span class="sxs-lookup"><span data-stu-id="62c0f-902">- Selection</span></span><br><span data-ttu-id="62c0f-903">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="62c0f-903">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="62c0f-904">関連項目</span><span class="sxs-lookup"><span data-stu-id="62c0f-904">See also</span></span>

- [<span data-ttu-id="62c0f-905">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="62c0f-905">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="62c0f-906">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="62c0f-906">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="62c0f-907">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="62c0f-907">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="62c0f-908">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="62c0f-908">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="62c0f-909">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="62c0f-909">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="62c0f-910">Office 365 ProPlus の更新履歴</span><span class="sxs-lookup"><span data-stu-id="62c0f-910">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="62c0f-911">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="62c0f-911">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="62c0f-912">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="62c0f-912">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="62c0f-913">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="62c0f-913">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="62c0f-914">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="62c0f-914">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="62c0f-915">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="62c0f-915">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
