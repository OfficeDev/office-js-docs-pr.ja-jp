---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、OneNote、Outlook、PowerPoint、Project、Word のサポートされる要件セット。
ms.date: 07/11/2019
localization_priority: Priority
ms.openlocfilehash: 2bfeb7cc5c6e8846f1d882abf3a0149302e53914
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771836"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="e2b1c-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e2b1c-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="e2b1c-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="e2b1c-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="e2b1c-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="e2b1c-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="e2b1c-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="e2b1c-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="e2b1c-108">Excel</span><span class="sxs-lookup"><span data-stu-id="e2b1c-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="e2b1c-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e2b1c-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="e2b1c-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="e2b1c-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="e2b1c-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="e2b1c-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="e2b1c-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-113">Office on the web</span><span class="sxs-lookup"><span data-stu-id="e2b1c-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="e2b1c-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-114">- TaskPane</span></span><br><span data-ttu-id="e2b1c-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-115">
        - Content</span></span><br><span data-ttu-id="e2b1c-116">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="e2b1c-116">
        - Custom Functions</span></span><br><span data-ttu-id="e2b1c-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="e2b1c-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="e2b1c-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e2b1c-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e2b1c-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e2b1c-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e2b1c-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e2b1c-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e2b1c-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e2b1c-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="e2b1c-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="e2b1c-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="e2b1c-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-130">
        - BindingEvents</span></span><br><span data-ttu-id="e2b1c-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-131">
        - CompressedFile</span></span><br><span data-ttu-id="e2b1c-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-132">
        - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-133">
        - File</span></span><br><span data-ttu-id="e2b1c-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-134">
        - MatrixBindings</span></span><br><span data-ttu-id="e2b1c-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="e2b1c-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-136">
        - Selection</span></span><br><span data-ttu-id="e2b1c-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-137">
        - Settings</span></span><br><span data-ttu-id="e2b1c-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-138">
        - TableBindings</span></span><br><span data-ttu-id="e2b1c-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-139">
        - TableCoercion</span></span><br><span data-ttu-id="e2b1c-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-140">
        - TextBindings</span></span><br><span data-ttu-id="e2b1c-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-142">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="e2b1c-142">Office on Windows</span></span><br><span data-ttu-id="e2b1c-143">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e2b1c-144">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-144">- TaskPane</span></span><br><span data-ttu-id="e2b1c-145">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-145">
        - Content</span></span><br><span data-ttu-id="e2b1c-146">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="e2b1c-146">
        - Custom Functions</span></span><br><span data-ttu-id="e2b1c-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="e2b1c-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="e2b1c-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e2b1c-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e2b1c-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e2b1c-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e2b1c-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e2b1c-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e2b1c-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e2b1c-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="e2b1c-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="e2b1c-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="e2b1c-160">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-160">
        - BindingEvents</span></span><br><span data-ttu-id="e2b1c-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-161">
        - CompressedFile</span></span><br><span data-ttu-id="e2b1c-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-162">
        - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-163">
        - File</span></span><br><span data-ttu-id="e2b1c-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-164">
        - MatrixBindings</span></span><br><span data-ttu-id="e2b1c-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-165">
        - MatrixCoercion</span></span><br><span data-ttu-id="e2b1c-166">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-166">
        - Selection</span></span><br><span data-ttu-id="e2b1c-167">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-167">
        - Settings</span></span><br><span data-ttu-id="e2b1c-168">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-168">
        - TableBindings</span></span><br><span data-ttu-id="e2b1c-169">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-169">
        - TableCoercion</span></span><br><span data-ttu-id="e2b1c-170">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-170">
        - TextBindings</span></span><br><span data-ttu-id="e2b1c-171">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-171">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-172">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="e2b1c-172">Office 2019 on Windows</span></span><br><span data-ttu-id="e2b1c-173">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-173">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="e2b1c-174">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-174">- TaskPane</span></span><br><span data-ttu-id="e2b1c-175">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-175">
        - Content</span></span><br><span data-ttu-id="e2b1c-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e2b1c-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e2b1c-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e2b1c-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e2b1c-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e2b1c-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e2b1c-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e2b1c-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e2b1c-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="e2b1c-187">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-187">- BindingEvents</span></span><br><span data-ttu-id="e2b1c-188">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-188">
        - CompressedFile</span></span><br><span data-ttu-id="e2b1c-189">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-189">
        - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-190">
        - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-190">
        - File</span></span><br><span data-ttu-id="e2b1c-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-191">
        - MatrixBindings</span></span><br><span data-ttu-id="e2b1c-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="e2b1c-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-193">
        - Selection</span></span><br><span data-ttu-id="e2b1c-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-194">
        - Settings</span></span><br><span data-ttu-id="e2b1c-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-195">
        - TableBindings</span></span><br><span data-ttu-id="e2b1c-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-196">
        - TableCoercion</span></span><br><span data-ttu-id="e2b1c-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-197">
        - TextBindings</span></span><br><span data-ttu-id="e2b1c-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-199">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="e2b1c-199">Office 2016 on Windows</span></span><br><span data-ttu-id="e2b1c-200">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-200">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="e2b1c-201">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-201">- TaskPane</span></span><br><span data-ttu-id="e2b1c-202">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-202">
        - Content</span></span></td>
    <td><span data-ttu-id="e2b1c-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="e2b1c-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="e2b1c-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="e2b1c-206">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-206">- BindingEvents</span></span><br><span data-ttu-id="e2b1c-207">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-207">
        - CompressedFile</span></span><br><span data-ttu-id="e2b1c-208">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-208">
        - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-209">
        - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-209">
        - File</span></span><br><span data-ttu-id="e2b1c-210">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-210">
        - MatrixBindings</span></span><br><span data-ttu-id="e2b1c-211">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-211">
        - MatrixCoercion</span></span><br><span data-ttu-id="e2b1c-212">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-212">
        - Selection</span></span><br><span data-ttu-id="e2b1c-213">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-213">
        - Settings</span></span><br><span data-ttu-id="e2b1c-214">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-214">
        - TableBindings</span></span><br><span data-ttu-id="e2b1c-215">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-215">
        - TableCoercion</span></span><br><span data-ttu-id="e2b1c-216">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-216">
        - TextBindings</span></span><br><span data-ttu-id="e2b1c-217">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-217">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-218">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="e2b1c-218">Office 2013 on Windows</span></span><br><span data-ttu-id="e2b1c-219">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-219">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="e2b1c-220">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-220">
        - TaskPane</span></span><br><span data-ttu-id="e2b1c-221">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-221">
        - Content</span></span></td>
    <td>  <span data-ttu-id="e2b1c-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="e2b1c-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="e2b1c-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="e2b1c-224">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-224">
        - BindingEvents</span></span><br><span data-ttu-id="e2b1c-225">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-225">
        - CompressedFile</span></span><br><span data-ttu-id="e2b1c-226">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-226">
        - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-227">
        - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-227">
        - File</span></span><br><span data-ttu-id="e2b1c-228">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-228">
        - MatrixBindings</span></span><br><span data-ttu-id="e2b1c-229">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-229">
        - MatrixCoercion</span></span><br><span data-ttu-id="e2b1c-230">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-230">
        - Selection</span></span><br><span data-ttu-id="e2b1c-231">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-231">
        - Settings</span></span><br><span data-ttu-id="e2b1c-232">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-232">
        - TableBindings</span></span><br><span data-ttu-id="e2b1c-233">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-233">
        - TableCoercion</span></span><br><span data-ttu-id="e2b1c-234">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-234">
        - TextBindings</span></span><br><span data-ttu-id="e2b1c-235">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-235">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-236">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="e2b1c-236">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="e2b1c-237">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-237">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="e2b1c-238">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-238">- TaskPane</span></span><br><span data-ttu-id="e2b1c-239">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-239">
        - Content</span></span><br><span data-ttu-id="e2b1c-240">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="e2b1c-240">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="e2b1c-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e2b1c-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e2b1c-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e2b1c-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e2b1c-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e2b1c-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e2b1c-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e2b1c-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="e2b1c-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="e2b1c-252">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-252">- BindingEvents</span></span><br><span data-ttu-id="e2b1c-253">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-253">
        - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-254">
        - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-254">
        - File</span></span><br><span data-ttu-id="e2b1c-255">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-255">
        - MatrixBindings</span></span><br><span data-ttu-id="e2b1c-256">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-256">
        - MatrixCoercion</span></span><br><span data-ttu-id="e2b1c-257">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-257">
        - Selection</span></span><br><span data-ttu-id="e2b1c-258">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-258">
        - Settings</span></span><br><span data-ttu-id="e2b1c-259">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-259">
        - TableBindings</span></span><br><span data-ttu-id="e2b1c-260">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-260">
        - TableCoercion</span></span><br><span data-ttu-id="e2b1c-261">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-261">
        - TextBindings</span></span><br><span data-ttu-id="e2b1c-262">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-262">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-263">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="e2b1c-263">Office apps on Mac</span></span><br><span data-ttu-id="e2b1c-264">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-264">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="e2b1c-265">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-265">- TaskPane</span></span><br><span data-ttu-id="e2b1c-266">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-266">
        - Content</span></span><br><span data-ttu-id="e2b1c-267">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="e2b1c-267">
        - Custom Functions</span></span><br><span data-ttu-id="e2b1c-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e2b1c-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e2b1c-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e2b1c-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e2b1c-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e2b1c-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e2b1c-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e2b1c-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e2b1c-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="e2b1c-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="e2b1c-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="e2b1c-281">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-281">- BindingEvents</span></span><br><span data-ttu-id="e2b1c-282">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-282">
        - CompressedFile</span></span><br><span data-ttu-id="e2b1c-283">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-283">
        - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-284">
        - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-284">
        - File</span></span><br><span data-ttu-id="e2b1c-285">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-285">
        - MatrixBindings</span></span><br><span data-ttu-id="e2b1c-286">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-286">
        - MatrixCoercion</span></span><br><span data-ttu-id="e2b1c-287">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-287">
        - PdfFile</span></span><br><span data-ttu-id="e2b1c-288">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-288">
        - Selection</span></span><br><span data-ttu-id="e2b1c-289">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-289">
        - Settings</span></span><br><span data-ttu-id="e2b1c-290">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-290">
        - TableBindings</span></span><br><span data-ttu-id="e2b1c-291">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-291">
        - TableCoercion</span></span><br><span data-ttu-id="e2b1c-292">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-292">
        - TextBindings</span></span><br><span data-ttu-id="e2b1c-293">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-293">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-294">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="e2b1c-294">Office 2019 for Mac</span></span><br><span data-ttu-id="e2b1c-295">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-295">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="e2b1c-296">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-296">- TaskPane</span></span><br><span data-ttu-id="e2b1c-297">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-297">
        - Content</span></span><br><span data-ttu-id="e2b1c-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e2b1c-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e2b1c-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e2b1c-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e2b1c-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e2b1c-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e2b1c-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e2b1c-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e2b1c-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="e2b1c-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-309">- BindingEvents</span></span><br><span data-ttu-id="e2b1c-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-310">
        - CompressedFile</span></span><br><span data-ttu-id="e2b1c-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-311">
        - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-312">
        - File</span></span><br><span data-ttu-id="e2b1c-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-313">
        - MatrixBindings</span></span><br><span data-ttu-id="e2b1c-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="e2b1c-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-315">
        - PdfFile</span></span><br><span data-ttu-id="e2b1c-316">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-316">
        - Selection</span></span><br><span data-ttu-id="e2b1c-317">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-317">
        - Settings</span></span><br><span data-ttu-id="e2b1c-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-318">
        - TableBindings</span></span><br><span data-ttu-id="e2b1c-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-319">
        - TableCoercion</span></span><br><span data-ttu-id="e2b1c-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-320">
        - TextBindings</span></span><br><span data-ttu-id="e2b1c-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-321">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-322">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="e2b1c-322">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="e2b1c-323">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-323">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="e2b1c-324">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-324">- TaskPane</span></span><br><span data-ttu-id="e2b1c-325">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-325">
        - Content</span></span></td>
    <td><span data-ttu-id="e2b1c-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="e2b1c-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="e2b1c-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="e2b1c-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-329">- BindingEvents</span></span><br><span data-ttu-id="e2b1c-330">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-330">
        - CompressedFile</span></span><br><span data-ttu-id="e2b1c-331">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-331">
        - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-332">
        - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-332">
        - File</span></span><br><span data-ttu-id="e2b1c-333">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-333">
        - MatrixBindings</span></span><br><span data-ttu-id="e2b1c-334">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-334">
        - MatrixCoercion</span></span><br><span data-ttu-id="e2b1c-335">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-335">
        - PdfFile</span></span><br><span data-ttu-id="e2b1c-336">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-336">
        - Selection</span></span><br><span data-ttu-id="e2b1c-337">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-337">
        - Settings</span></span><br><span data-ttu-id="e2b1c-338">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-338">
        - TableBindings</span></span><br><span data-ttu-id="e2b1c-339">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-339">
        - TableCoercion</span></span><br><span data-ttu-id="e2b1c-340">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-340">
        - TextBindings</span></span><br><span data-ttu-id="e2b1c-341">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-341">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="e2b1c-342">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="e2b1c-342">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="e2b1c-343">カスタム関数</span><span class="sxs-lookup"><span data-stu-id="e2b1c-343">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="e2b1c-344">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e2b1c-344">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="e2b1c-345">拡張点</span><span class="sxs-lookup"><span data-stu-id="e2b1c-345">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="e2b1c-346">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="e2b1c-346">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="e2b1c-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-348">Office on the web</span><span class="sxs-lookup"><span data-stu-id="e2b1c-348">Office on the web</span></span></td>
    <td><span data-ttu-id="e2b1c-349">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="e2b1c-349">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="e2b1c-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-351">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="e2b1c-351">Office on Windows</span></span><br><span data-ttu-id="e2b1c-352">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-352">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="e2b1c-353">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="e2b1c-353">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="e2b1c-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-355">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="e2b1c-355">Office for Mac</span></span><br><span data-ttu-id="e2b1c-356">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-356">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="e2b1c-357">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="e2b1c-357">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="e2b1c-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="e2b1c-359">Outlook</span><span class="sxs-lookup"><span data-stu-id="e2b1c-359">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e2b1c-360">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e2b1c-360">Platform</span></span></th>
    <th><span data-ttu-id="e2b1c-361">拡張点</span><span class="sxs-lookup"><span data-stu-id="e2b1c-361">Extension points</span></span></th>
    <th><span data-ttu-id="e2b1c-362">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="e2b1c-362">API requirement sets</span></span></th>
    <th><span data-ttu-id="e2b1c-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-364">Office on the web</span><span class="sxs-lookup"><span data-stu-id="e2b1c-364">Office on the web</span></span><br><span data-ttu-id="e2b1c-365">(新規)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-365">New</span></span></td>
    <td> <span data-ttu-id="e2b1c-366">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e2b1c-366">- Mail Read</span></span><br><span data-ttu-id="e2b1c-367">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="e2b1c-367">
      - Mail Compose</span></span><br><span data-ttu-id="e2b1c-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e2b1c-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e2b1c-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e2b1c-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e2b1c-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e2b1c-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e2b1c-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e2b1c-376">使用不可</span><span class="sxs-lookup"><span data-stu-id="e2b1c-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-377">Office on the web</span><span class="sxs-lookup"><span data-stu-id="e2b1c-377">Office on the web</span></span><br><span data-ttu-id="e2b1c-378">(クラシック)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-378">Classic</span></span></td>
    <td> <span data-ttu-id="e2b1c-379">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e2b1c-379">- Mail Read</span></span><br><span data-ttu-id="e2b1c-380">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="e2b1c-380">
      - Mail Compose</span></span><br><span data-ttu-id="e2b1c-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e2b1c-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e2b1c-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e2b1c-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e2b1c-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e2b1c-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e2b1c-388">使用不可</span><span class="sxs-lookup"><span data-stu-id="e2b1c-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-389">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="e2b1c-389">Office on Windows</span></span><br><span data-ttu-id="e2b1c-390">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-390">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e2b1c-391">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e2b1c-391">- Mail Read</span></span><br><span data-ttu-id="e2b1c-392">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="e2b1c-392">
      - Mail Compose</span></span><br><span data-ttu-id="e2b1c-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="e2b1c-394">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="e2b1c-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="e2b1c-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e2b1c-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e2b1c-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e2b1c-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e2b1c-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e2b1c-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e2b1c-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e2b1c-402">使用不可</span><span class="sxs-lookup"><span data-stu-id="e2b1c-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-403">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="e2b1c-403">Office 2019 on Windows</span></span><br><span data-ttu-id="e2b1c-404">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e2b1c-405">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e2b1c-405">- Mail Read</span></span><br><span data-ttu-id="e2b1c-406">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="e2b1c-406">
      - Mail Compose</span></span><br><span data-ttu-id="e2b1c-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="e2b1c-408">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="e2b1c-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="e2b1c-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e2b1c-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e2b1c-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e2b1c-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e2b1c-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e2b1c-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e2b1c-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e2b1c-416">使用不可</span><span class="sxs-lookup"><span data-stu-id="e2b1c-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-417">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="e2b1c-417">Office 2016 on Windows</span></span><br><span data-ttu-id="e2b1c-418">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e2b1c-419">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e2b1c-419">- Mail Read</span></span><br><span data-ttu-id="e2b1c-420">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="e2b1c-420">
      - Mail Compose</span></span><br><span data-ttu-id="e2b1c-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="e2b1c-422">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="e2b1c-422">
      - Modules</span></span></td>
    <td> <span data-ttu-id="e2b1c-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e2b1c-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e2b1c-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e2b1c-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="e2b1c-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="e2b1c-427">使用不可</span><span class="sxs-lookup"><span data-stu-id="e2b1c-427">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-428">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="e2b1c-428">Office 2013 on Windows</span></span><br><span data-ttu-id="e2b1c-429">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-429">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e2b1c-430">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e2b1c-430">- Mail Read</span></span><br><span data-ttu-id="e2b1c-431">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="e2b1c-431">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="e2b1c-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e2b1c-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e2b1c-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="e2b1c-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="e2b1c-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="e2b1c-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="e2b1c-436">使用不可</span><span class="sxs-lookup"><span data-stu-id="e2b1c-436">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-437">iOS 上の Office</span><span class="sxs-lookup"><span data-stu-id="e2b1c-437">Office apps on iOS</span></span><br><span data-ttu-id="e2b1c-438">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-438">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e2b1c-439">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e2b1c-439">- Mail Read</span></span><br><span data-ttu-id="e2b1c-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e2b1c-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e2b1c-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e2b1c-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e2b1c-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="e2b1c-446">使用不可</span><span class="sxs-lookup"><span data-stu-id="e2b1c-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-447">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="e2b1c-447">Office apps on Mac</span></span><br><span data-ttu-id="e2b1c-448">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-448">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e2b1c-449">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e2b1c-449">- Mail Read</span></span><br><span data-ttu-id="e2b1c-450">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="e2b1c-450">
      - Mail Compose</span></span><br><span data-ttu-id="e2b1c-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e2b1c-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e2b1c-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e2b1c-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e2b1c-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e2b1c-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e2b1c-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e2b1c-459">使用不可</span><span class="sxs-lookup"><span data-stu-id="e2b1c-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-460">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="e2b1c-460">Office 2019 for Mac</span></span><br><span data-ttu-id="e2b1c-461">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e2b1c-462">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e2b1c-462">- Mail Read</span></span><br><span data-ttu-id="e2b1c-463">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="e2b1c-463">
      - Mail Compose</span></span><br><span data-ttu-id="e2b1c-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e2b1c-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e2b1c-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e2b1c-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e2b1c-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e2b1c-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e2b1c-471">使用不可</span><span class="sxs-lookup"><span data-stu-id="e2b1c-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-472">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="e2b1c-472">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="e2b1c-473">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e2b1c-474">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e2b1c-474">- Mail Read</span></span><br><span data-ttu-id="e2b1c-475">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="e2b1c-475">
      - Mail Compose</span></span><br><span data-ttu-id="e2b1c-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e2b1c-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e2b1c-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e2b1c-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e2b1c-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e2b1c-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e2b1c-483">使用不可</span><span class="sxs-lookup"><span data-stu-id="e2b1c-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-484">Android 上の Office</span><span class="sxs-lookup"><span data-stu-id="e2b1c-484">Office apps on Android</span></span><br><span data-ttu-id="e2b1c-485">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e2b1c-486">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e2b1c-486">- Mail Read</span></span><br><span data-ttu-id="e2b1c-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e2b1c-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e2b1c-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e2b1c-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e2b1c-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="e2b1c-493">利用不可</span><span class="sxs-lookup"><span data-stu-id="e2b1c-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="e2b1c-494">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="e2b1c-494">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="e2b1c-495">Word</span><span class="sxs-lookup"><span data-stu-id="e2b1c-495">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e2b1c-496">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e2b1c-496">Platform</span></span></th>
    <th><span data-ttu-id="e2b1c-497">拡張点</span><span class="sxs-lookup"><span data-stu-id="e2b1c-497">Extension points</span></span></th>
    <th><span data-ttu-id="e2b1c-498">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="e2b1c-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="e2b1c-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-500">Office on the web</span><span class="sxs-lookup"><span data-stu-id="e2b1c-500">Office on the web</span></span></td>
    <td> <span data-ttu-id="e2b1c-501">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-501">- TaskPane</span></span><br><span data-ttu-id="e2b1c-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e2b1c-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e2b1c-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="e2b1c-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-509">- BindingEvents</span></span><br><span data-ttu-id="e2b1c-510">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e2b1c-510">
         - CustomXmlParts</span></span><br><span data-ttu-id="e2b1c-511">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-511">
         - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-512">
         - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-512">
         - File</span></span><br><span data-ttu-id="e2b1c-513">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-513">
         - HtmlCoercion</span></span><br><span data-ttu-id="e2b1c-514">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-514">
         - MatrixBindings</span></span><br><span data-ttu-id="e2b1c-515">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-515">
         - MatrixCoercion</span></span><br><span data-ttu-id="e2b1c-516">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-516">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e2b1c-517">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-517">
         - PdfFile</span></span><br><span data-ttu-id="e2b1c-518">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-518">
         - Selection</span></span><br><span data-ttu-id="e2b1c-519">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-519">
         - Settings</span></span><br><span data-ttu-id="e2b1c-520">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-520">
         - TableBindings</span></span><br><span data-ttu-id="e2b1c-521">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-521">
         - TableCoercion</span></span><br><span data-ttu-id="e2b1c-522">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-522">
         - TextBindings</span></span><br><span data-ttu-id="e2b1c-523">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-523">
         - TextCoercion</span></span><br><span data-ttu-id="e2b1c-524">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-524">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-525">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="e2b1c-525">Office on Windows</span></span><br><span data-ttu-id="e2b1c-526">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-526">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e2b1c-527">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-527">- TaskPane</span></span><br><span data-ttu-id="e2b1c-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e2b1c-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e2b1c-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="e2b1c-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-535">- BindingEvents</span></span><br><span data-ttu-id="e2b1c-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-536">
         - CompressedFile</span></span><br><span data-ttu-id="e2b1c-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e2b1c-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="e2b1c-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-538">
         - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-539">
         - File</span></span><br><span data-ttu-id="e2b1c-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="e2b1c-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-541">
         - MatrixBindings</span></span><br><span data-ttu-id="e2b1c-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="e2b1c-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e2b1c-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-544">
         - PdfFile</span></span><br><span data-ttu-id="e2b1c-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-545">
         - Selection</span></span><br><span data-ttu-id="e2b1c-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-546">
         - Settings</span></span><br><span data-ttu-id="e2b1c-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-547">
         - TableBindings</span></span><br><span data-ttu-id="e2b1c-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-548">
         - TableCoercion</span></span><br><span data-ttu-id="e2b1c-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-549">
         - TextBindings</span></span><br><span data-ttu-id="e2b1c-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-550">
         - TextCoercion</span></span><br><span data-ttu-id="e2b1c-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-551">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-552">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="e2b1c-552">Office 2019 on Windows</span></span><br><span data-ttu-id="e2b1c-553">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-553">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e2b1c-554">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-554">- TaskPane</span></span><br><span data-ttu-id="e2b1c-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e2b1c-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e2b1c-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-561">- BindingEvents</span></span><br><span data-ttu-id="e2b1c-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-562">
         - CompressedFile</span></span><br><span data-ttu-id="e2b1c-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e2b1c-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="e2b1c-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-564">
         - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-565">
         - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-565">
         - File</span></span><br><span data-ttu-id="e2b1c-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="e2b1c-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-567">
         - MatrixBindings</span></span><br><span data-ttu-id="e2b1c-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="e2b1c-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e2b1c-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-570">
         - PdfFile</span></span><br><span data-ttu-id="e2b1c-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-571">
         - Selection</span></span><br><span data-ttu-id="e2b1c-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-572">
         - Settings</span></span><br><span data-ttu-id="e2b1c-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-573">
         - TableBindings</span></span><br><span data-ttu-id="e2b1c-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-574">
         - TableCoercion</span></span><br><span data-ttu-id="e2b1c-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-575">
         - TextBindings</span></span><br><span data-ttu-id="e2b1c-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-576">
         - TextCoercion</span></span><br><span data-ttu-id="e2b1c-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-578">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="e2b1c-578">Office 2016 on Windows</span></span><br><span data-ttu-id="e2b1c-579">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e2b1c-580">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-580">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e2b1c-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="e2b1c-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="e2b1c-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-584">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-584">- BindingEvents</span></span><br><span data-ttu-id="e2b1c-585">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-585">
         - CompressedFile</span></span><br><span data-ttu-id="e2b1c-586">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e2b1c-586">
         - CustomXmlParts</span></span><br><span data-ttu-id="e2b1c-587">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-587">
         - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-588">
         - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-588">
         - File</span></span><br><span data-ttu-id="e2b1c-589">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-589">
         - HtmlCoercion</span></span><br><span data-ttu-id="e2b1c-590">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-590">
         - MatrixBindings</span></span><br><span data-ttu-id="e2b1c-591">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-591">
         - MatrixCoercion</span></span><br><span data-ttu-id="e2b1c-592">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-592">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e2b1c-593">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-593">
         - PdfFile</span></span><br><span data-ttu-id="e2b1c-594">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-594">
         - Selection</span></span><br><span data-ttu-id="e2b1c-595">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-595">
         - Settings</span></span><br><span data-ttu-id="e2b1c-596">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-596">
         - TableBindings</span></span><br><span data-ttu-id="e2b1c-597">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-597">
         - TableCoercion</span></span><br><span data-ttu-id="e2b1c-598">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-598">
         - TextBindings</span></span><br><span data-ttu-id="e2b1c-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-599">
         - TextCoercion</span></span><br><span data-ttu-id="e2b1c-600">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-600">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-601">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="e2b1c-601">Office 2013 on Windows</span></span><br><span data-ttu-id="e2b1c-602">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-602">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e2b1c-603">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-603">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e2b1c-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="e2b1c-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="e2b1c-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-606">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-606">- BindingEvents</span></span><br><span data-ttu-id="e2b1c-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-607">
         - CompressedFile</span></span><br><span data-ttu-id="e2b1c-608">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e2b1c-608">
         - CustomXmlParts</span></span><br><span data-ttu-id="e2b1c-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-609">
         - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-610">
         - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-610">
         - File</span></span><br><span data-ttu-id="e2b1c-611">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-611">
         - HtmlCoercion</span></span><br><span data-ttu-id="e2b1c-612">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-612">
         - MatrixBindings</span></span><br><span data-ttu-id="e2b1c-613">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-613">
         - MatrixCoercion</span></span><br><span data-ttu-id="e2b1c-614">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-614">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e2b1c-615">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-615">
         - PdfFile</span></span><br><span data-ttu-id="e2b1c-616">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-616">
         - Selection</span></span><br><span data-ttu-id="e2b1c-617">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-617">
         - Settings</span></span><br><span data-ttu-id="e2b1c-618">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-618">
         - TableBindings</span></span><br><span data-ttu-id="e2b1c-619">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-619">
         - TableCoercion</span></span><br><span data-ttu-id="e2b1c-620">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-620">
         - TextBindings</span></span><br><span data-ttu-id="e2b1c-621">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-621">
         - TextCoercion</span></span><br><span data-ttu-id="e2b1c-622">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-622">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-623">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="e2b1c-623">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="e2b1c-624">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-624">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e2b1c-625">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-625">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e2b1c-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e2b1c-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e2b1c-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="e2b1c-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-631">- BindingEvents</span></span><br><span data-ttu-id="e2b1c-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-632">
         - CompressedFile</span></span><br><span data-ttu-id="e2b1c-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e2b1c-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="e2b1c-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-634">
         - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-635">
         - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-635">
         - File</span></span><br><span data-ttu-id="e2b1c-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="e2b1c-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-637">
         - MatrixBindings</span></span><br><span data-ttu-id="e2b1c-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="e2b1c-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e2b1c-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-640">
         - PdfFile</span></span><br><span data-ttu-id="e2b1c-641">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-641">
         - Selection</span></span><br><span data-ttu-id="e2b1c-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-642">
         - Settings</span></span><br><span data-ttu-id="e2b1c-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-643">
         - TableBindings</span></span><br><span data-ttu-id="e2b1c-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-644">
         - TableCoercion</span></span><br><span data-ttu-id="e2b1c-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-645">
         - TextBindings</span></span><br><span data-ttu-id="e2b1c-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-646">
         - TextCoercion</span></span><br><span data-ttu-id="e2b1c-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-647">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-648">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="e2b1c-648">Office apps on Mac</span></span><br><span data-ttu-id="e2b1c-649">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e2b1c-650">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-650">- TaskPane</span></span><br><span data-ttu-id="e2b1c-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e2b1c-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e2b1c-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="e2b1c-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="e2b1c-658">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-658">- BindingEvents</span></span><br><span data-ttu-id="e2b1c-659">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-659">
         - CompressedFile</span></span><br><span data-ttu-id="e2b1c-660">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e2b1c-660">
         - CustomXmlParts</span></span><br><span data-ttu-id="e2b1c-661">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-661">
         - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-662">
         - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-662">
         - File</span></span><br><span data-ttu-id="e2b1c-663">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-663">
         - HtmlCoercion</span></span><br><span data-ttu-id="e2b1c-664">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-664">
         - MatrixBindings</span></span><br><span data-ttu-id="e2b1c-665">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-665">
         - MatrixCoercion</span></span><br><span data-ttu-id="e2b1c-666">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-666">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e2b1c-667">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-667">
         - PdfFile</span></span><br><span data-ttu-id="e2b1c-668">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-668">
         - Selection</span></span><br><span data-ttu-id="e2b1c-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-669">
         - Settings</span></span><br><span data-ttu-id="e2b1c-670">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-670">
         - TableBindings</span></span><br><span data-ttu-id="e2b1c-671">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-671">
         - TableCoercion</span></span><br><span data-ttu-id="e2b1c-672">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-672">
         - TextBindings</span></span><br><span data-ttu-id="e2b1c-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-673">
         - TextCoercion</span></span><br><span data-ttu-id="e2b1c-674">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-674">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-675">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="e2b1c-675">Office 2019 for Mac</span></span><br><span data-ttu-id="e2b1c-676">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-676">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e2b1c-677">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-677">- TaskPane</span></span><br><span data-ttu-id="e2b1c-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e2b1c-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e2b1c-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="e2b1c-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-684">- BindingEvents</span></span><br><span data-ttu-id="e2b1c-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-685">
         - CompressedFile</span></span><br><span data-ttu-id="e2b1c-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e2b1c-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="e2b1c-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-687">
         - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-688">
         - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-688">
         - File</span></span><br><span data-ttu-id="e2b1c-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="e2b1c-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-690">
         - MatrixBindings</span></span><br><span data-ttu-id="e2b1c-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="e2b1c-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e2b1c-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-693">
         - PdfFile</span></span><br><span data-ttu-id="e2b1c-694">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-694">
         - Selection</span></span><br><span data-ttu-id="e2b1c-695">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-695">
         - Settings</span></span><br><span data-ttu-id="e2b1c-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-696">
         - TableBindings</span></span><br><span data-ttu-id="e2b1c-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-697">
         - TableCoercion</span></span><br><span data-ttu-id="e2b1c-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-698">
         - TextBindings</span></span><br><span data-ttu-id="e2b1c-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-699">
         - TextCoercion</span></span><br><span data-ttu-id="e2b1c-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-701">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="e2b1c-701">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="e2b1c-702">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e2b1c-703">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-703">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e2b1c-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="e2b1c-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="e2b1c-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-707">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-707">- BindingEvents</span></span><br><span data-ttu-id="e2b1c-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-708">
         - CompressedFile</span></span><br><span data-ttu-id="e2b1c-709">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e2b1c-709">
         - CustomXmlParts</span></span><br><span data-ttu-id="e2b1c-710">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-710">
         - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-711">
         - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-711">
         - File</span></span><br><span data-ttu-id="e2b1c-712">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-712">
         - HtmlCoercion</span></span><br><span data-ttu-id="e2b1c-713">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-713">
         - MatrixBindings</span></span><br><span data-ttu-id="e2b1c-714">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-714">
         - MatrixCoercion</span></span><br><span data-ttu-id="e2b1c-715">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-715">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e2b1c-716">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-716">
         - PdfFile</span></span><br><span data-ttu-id="e2b1c-717">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-717">
         - Selection</span></span><br><span data-ttu-id="e2b1c-718">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-718">
         - Settings</span></span><br><span data-ttu-id="e2b1c-719">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-719">
         - TableBindings</span></span><br><span data-ttu-id="e2b1c-720">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-720">
         - TableCoercion</span></span><br><span data-ttu-id="e2b1c-721">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-721">
         - TextBindings</span></span><br><span data-ttu-id="e2b1c-722">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-722">
         - TextCoercion</span></span><br><span data-ttu-id="e2b1c-723">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-723">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="e2b1c-724">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="e2b1c-724">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="e2b1c-725">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="e2b1c-725">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e2b1c-726">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e2b1c-726">Platform</span></span></th>
    <th><span data-ttu-id="e2b1c-727">拡張点</span><span class="sxs-lookup"><span data-stu-id="e2b1c-727">Extension points</span></span></th>
    <th><span data-ttu-id="e2b1c-728">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="e2b1c-728">API requirement sets</span></span></th>
    <th><span data-ttu-id="e2b1c-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-730">Office on the web</span><span class="sxs-lookup"><span data-stu-id="e2b1c-730">Office on the web</span></span></td>
    <td> <span data-ttu-id="e2b1c-731">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-731">- Content</span></span><br><span data-ttu-id="e2b1c-732">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-732">
         - TaskPane</span></span><br><span data-ttu-id="e2b1c-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="e2b1c-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-737">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e2b1c-737">- ActiveView</span></span><br><span data-ttu-id="e2b1c-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-738">
         - CompressedFile</span></span><br><span data-ttu-id="e2b1c-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-739">
         - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-740">
         - File</span></span><br><span data-ttu-id="e2b1c-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-741">
         - PdfFile</span></span><br><span data-ttu-id="e2b1c-742">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-742">
         - Selection</span></span><br><span data-ttu-id="e2b1c-743">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-743">
         - Settings</span></span><br><span data-ttu-id="e2b1c-744">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-744">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-745">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="e2b1c-745">Office on Windows</span></span><br><span data-ttu-id="e2b1c-746">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-746">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e2b1c-747">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-747">- Content</span></span><br><span data-ttu-id="e2b1c-748">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-748">
         - TaskPane</span></span><br><span data-ttu-id="e2b1c-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="e2b1c-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-753">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e2b1c-753">- ActiveView</span></span><br><span data-ttu-id="e2b1c-754">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-754">
         - CompressedFile</span></span><br><span data-ttu-id="e2b1c-755">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-755">
         - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-756">
         - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-756">
         - File</span></span><br><span data-ttu-id="e2b1c-757">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-757">
         - PdfFile</span></span><br><span data-ttu-id="e2b1c-758">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-758">
         - Selection</span></span><br><span data-ttu-id="e2b1c-759">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-759">
         - Settings</span></span><br><span data-ttu-id="e2b1c-760">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-760">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-761">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="e2b1c-761">Office 2019 on Windows</span></span><br><span data-ttu-id="e2b1c-762">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-762">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e2b1c-763">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-763">- Content</span></span><br><span data-ttu-id="e2b1c-764">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-764">
         - TaskPane</span></span><br><span data-ttu-id="e2b1c-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e2b1c-768">- ActiveView</span></span><br><span data-ttu-id="e2b1c-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-769">
         - CompressedFile</span></span><br><span data-ttu-id="e2b1c-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-770">
         - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-771">
         - File</span></span><br><span data-ttu-id="e2b1c-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-772">
         - PdfFile</span></span><br><span data-ttu-id="e2b1c-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-773">
         - Selection</span></span><br><span data-ttu-id="e2b1c-774">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-774">
         - Settings</span></span><br><span data-ttu-id="e2b1c-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-776">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="e2b1c-776">Office 2016 on Windows</span></span><br><span data-ttu-id="e2b1c-777">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-777">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e2b1c-778">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-778">- Content</span></span><br><span data-ttu-id="e2b1c-779">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-779">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="e2b1c-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="e2b1c-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="e2b1c-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-782">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e2b1c-782">- ActiveView</span></span><br><span data-ttu-id="e2b1c-783">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-783">
         - CompressedFile</span></span><br><span data-ttu-id="e2b1c-784">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-784">
         - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-785">
         - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-785">
         - File</span></span><br><span data-ttu-id="e2b1c-786">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-786">
         - PdfFile</span></span><br><span data-ttu-id="e2b1c-787">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-787">
         - Selection</span></span><br><span data-ttu-id="e2b1c-788">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-788">
         - Settings</span></span><br><span data-ttu-id="e2b1c-789">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-789">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-790">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="e2b1c-790">Office 2013 on Windows</span></span><br><span data-ttu-id="e2b1c-791">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-791">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e2b1c-792">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-792">- Content</span></span><br><span data-ttu-id="e2b1c-793">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-793">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="e2b1c-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="e2b1c-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="e2b1c-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e2b1c-796">- ActiveView</span></span><br><span data-ttu-id="e2b1c-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-797">
         - CompressedFile</span></span><br><span data-ttu-id="e2b1c-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-798">
         - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-799">
         - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-799">
         - File</span></span><br><span data-ttu-id="e2b1c-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-800">
         - PdfFile</span></span><br><span data-ttu-id="e2b1c-801">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-801">
         - Selection</span></span><br><span data-ttu-id="e2b1c-802">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-802">
         - Settings</span></span><br><span data-ttu-id="e2b1c-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-804">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="e2b1c-804">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="e2b1c-805">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-805">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e2b1c-806">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-806">- Content</span></span><br><span data-ttu-id="e2b1c-807">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="e2b1c-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e2b1c-810">- ActiveView</span></span><br><span data-ttu-id="e2b1c-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-811">
         - CompressedFile</span></span><br><span data-ttu-id="e2b1c-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-812">
         - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-813">
         - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-813">
         - File</span></span><br><span data-ttu-id="e2b1c-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-814">
         - PdfFile</span></span><br><span data-ttu-id="e2b1c-815">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-815">
         - Selection</span></span><br><span data-ttu-id="e2b1c-816">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-816">
         - Settings</span></span><br><span data-ttu-id="e2b1c-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-818">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="e2b1c-818">Office apps on Mac</span></span><br><span data-ttu-id="e2b1c-819">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-819">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="e2b1c-820">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-820">- Content</span></span><br><span data-ttu-id="e2b1c-821">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-821">
         - TaskPane</span></span><br><span data-ttu-id="e2b1c-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="e2b1c-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-826">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e2b1c-826">- ActiveView</span></span><br><span data-ttu-id="e2b1c-827">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-827">
         - CompressedFile</span></span><br><span data-ttu-id="e2b1c-828">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-828">
         - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-829">
         - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-829">
         - File</span></span><br><span data-ttu-id="e2b1c-830">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-830">
         - PdfFile</span></span><br><span data-ttu-id="e2b1c-831">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-831">
         - Selection</span></span><br><span data-ttu-id="e2b1c-832">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-832">
         - Settings</span></span><br><span data-ttu-id="e2b1c-833">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-833">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-834">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="e2b1c-834">Office 2019 for Mac</span></span><br><span data-ttu-id="e2b1c-835">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-835">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e2b1c-836">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-836">- Content</span></span><br><span data-ttu-id="e2b1c-837">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-837">
         - TaskPane</span></span><br><span data-ttu-id="e2b1c-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-841">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e2b1c-841">- ActiveView</span></span><br><span data-ttu-id="e2b1c-842">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-842">
         - CompressedFile</span></span><br><span data-ttu-id="e2b1c-843">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-843">
         - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-844">
         - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-844">
         - File</span></span><br><span data-ttu-id="e2b1c-845">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-845">
         - PdfFile</span></span><br><span data-ttu-id="e2b1c-846">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-846">
         - Selection</span></span><br><span data-ttu-id="e2b1c-847">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-847">
         - Settings</span></span><br><span data-ttu-id="e2b1c-848">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-848">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-849">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="e2b1c-849">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="e2b1c-850">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-850">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e2b1c-851">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-851">- Content</span></span><br><span data-ttu-id="e2b1c-852">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-852">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="e2b1c-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="e2b1c-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="e2b1c-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-855">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e2b1c-855">- ActiveView</span></span><br><span data-ttu-id="e2b1c-856">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-856">
         - CompressedFile</span></span><br><span data-ttu-id="e2b1c-857">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-857">
         - DocumentEvents</span></span><br><span data-ttu-id="e2b1c-858">
         - File</span><span class="sxs-lookup"><span data-stu-id="e2b1c-858">
         - File</span></span><br><span data-ttu-id="e2b1c-859">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e2b1c-859">
         - PdfFile</span></span><br><span data-ttu-id="e2b1c-860">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-860">
         - Selection</span></span><br><span data-ttu-id="e2b1c-861">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-861">
         - Settings</span></span><br><span data-ttu-id="e2b1c-862">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-862">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="e2b1c-863">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="e2b1c-863">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="e2b1c-864">OneNote</span><span class="sxs-lookup"><span data-stu-id="e2b1c-864">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e2b1c-865">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e2b1c-865">Platform</span></span></th>
    <th><span data-ttu-id="e2b1c-866">拡張点</span><span class="sxs-lookup"><span data-stu-id="e2b1c-866">Extension points</span></span></th>
    <th><span data-ttu-id="e2b1c-867">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="e2b1c-867">API requirement sets</span></span></th>
    <th><span data-ttu-id="e2b1c-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-869">Office on the web</span><span class="sxs-lookup"><span data-stu-id="e2b1c-869">Office on the web</span></span></td>
    <td> <span data-ttu-id="e2b1c-870">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-870">- Content</span></span><br><span data-ttu-id="e2b1c-871">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-871">
         - TaskPane</span></span><br><span data-ttu-id="e2b1c-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="e2b1c-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-876">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e2b1c-876">- DocumentEvents</span></span><br><span data-ttu-id="e2b1c-877">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-877">
         - HtmlCoercion</span></span><br><span data-ttu-id="e2b1c-878">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e2b1c-878">
         - Settings</span></span><br><span data-ttu-id="e2b1c-879">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-879">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="e2b1c-880">Project</span><span class="sxs-lookup"><span data-stu-id="e2b1c-880">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e2b1c-881">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e2b1c-881">Platform</span></span></th>
    <th><span data-ttu-id="e2b1c-882">拡張点</span><span class="sxs-lookup"><span data-stu-id="e2b1c-882">Extension points</span></span></th>
    <th><span data-ttu-id="e2b1c-883">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="e2b1c-883">API requirement sets</span></span></th>
    <th><span data-ttu-id="e2b1c-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-885">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="e2b1c-885">Office 2019 on Windows</span></span><br><span data-ttu-id="e2b1c-886">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-886">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e2b1c-887">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-887">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e2b1c-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-889">- Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-889">- Selection</span></span><br><span data-ttu-id="e2b1c-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-890">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-891">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="e2b1c-891">Office 2016 on Windows</span></span><br><span data-ttu-id="e2b1c-892">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-892">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e2b1c-893">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-893">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e2b1c-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-895">- Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-895">- Selection</span></span><br><span data-ttu-id="e2b1c-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-896">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e2b1c-897">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="e2b1c-897">Office 2013 on Windows</span></span><br><span data-ttu-id="e2b1c-898">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="e2b1c-898">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="e2b1c-899">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e2b1c-899">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e2b1c-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e2b1c-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e2b1c-901">- Selection</span><span class="sxs-lookup"><span data-stu-id="e2b1c-901">- Selection</span></span><br><span data-ttu-id="e2b1c-902">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e2b1c-902">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="e2b1c-903">関連項目</span><span class="sxs-lookup"><span data-stu-id="e2b1c-903">See also</span></span>

- [<span data-ttu-id="e2b1c-904">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="e2b1c-904">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="e2b1c-905">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="e2b1c-905">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="e2b1c-906">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="e2b1c-906">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="e2b1c-907">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="e2b1c-907">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="e2b1c-908">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="e2b1c-908">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="e2b1c-909">Office 365 ProPlus の更新履歴</span><span class="sxs-lookup"><span data-stu-id="e2b1c-909">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="e2b1c-910">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="e2b1c-910">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="e2b1c-911">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="e2b1c-911">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="e2b1c-912">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="e2b1c-912">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="e2b1c-913">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="e2b1c-913">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="e2b1c-914">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="e2b1c-914">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
