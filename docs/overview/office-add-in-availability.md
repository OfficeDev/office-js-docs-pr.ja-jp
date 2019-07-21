---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、OneNote、Outlook、PowerPoint、Project、Word のサポートされる要件セット。
ms.date: 07/18/2019
localization_priority: Priority
ms.openlocfilehash: 510f2419d5d364a536f8c96f2057505161f03993
ms.sourcegitcommit: 6d9b4820a62a914c50cef13af8b80ce626034c26
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/19/2019
ms.locfileid: "35804647"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="87898-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="87898-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="87898-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="87898-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="87898-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="87898-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="87898-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="87898-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="87898-108">Excel</span><span class="sxs-lookup"><span data-stu-id="87898-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="87898-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="87898-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="87898-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="87898-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="87898-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="87898-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="87898-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="87898-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-113">Office on the web</span><span class="sxs-lookup"><span data-stu-id="87898-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="87898-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-114">- TaskPane</span></span><br><span data-ttu-id="87898-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="87898-115">
        - Content</span></span><br><span data-ttu-id="87898-116">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="87898-116">
        - Custom Functions</span></span><br><span data-ttu-id="87898-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="87898-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="87898-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="87898-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="87898-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="87898-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="87898-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="87898-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="87898-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="87898-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="87898-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="87898-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="87898-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="87898-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="87898-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="87898-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="87898-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="87898-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="87898-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="87898-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="87898-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="87898-130">
        - BindingEvents</span></span><br><span data-ttu-id="87898-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-131">
        - CompressedFile</span></span><br><span data-ttu-id="87898-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-132">
        - DocumentEvents</span></span><br><span data-ttu-id="87898-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="87898-133">
        - File</span></span><br><span data-ttu-id="87898-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="87898-134">
        - MatrixBindings</span></span><br><span data-ttu-id="87898-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="87898-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-136">
        - Selection</span></span><br><span data-ttu-id="87898-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-137">
        - Settings</span></span><br><span data-ttu-id="87898-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="87898-138">
        - TableBindings</span></span><br><span data-ttu-id="87898-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-139">
        - TableCoercion</span></span><br><span data-ttu-id="87898-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="87898-140">
        - TextBindings</span></span><br><span data-ttu-id="87898-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-142">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="87898-142">Office on Windows</span></span><br><span data-ttu-id="87898-143">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="87898-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="87898-144">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-144">- TaskPane</span></span><br><span data-ttu-id="87898-145">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="87898-145">
        - Content</span></span><br><span data-ttu-id="87898-146">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="87898-146">
        - Custom Functions</span></span><br><span data-ttu-id="87898-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="87898-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="87898-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="87898-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="87898-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="87898-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="87898-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="87898-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="87898-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="87898-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="87898-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="87898-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="87898-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="87898-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="87898-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="87898-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="87898-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="87898-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="87898-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="87898-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="87898-160">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="87898-160">
        - BindingEvents</span></span><br><span data-ttu-id="87898-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-161">
        - CompressedFile</span></span><br><span data-ttu-id="87898-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-162">
        - DocumentEvents</span></span><br><span data-ttu-id="87898-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="87898-163">
        - File</span></span><br><span data-ttu-id="87898-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="87898-164">
        - MatrixBindings</span></span><br><span data-ttu-id="87898-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-165">
        - MatrixCoercion</span></span><br><span data-ttu-id="87898-166">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-166">
        - Selection</span></span><br><span data-ttu-id="87898-167">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-167">
        - Settings</span></span><br><span data-ttu-id="87898-168">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="87898-168">
        - TableBindings</span></span><br><span data-ttu-id="87898-169">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-169">
        - TableCoercion</span></span><br><span data-ttu-id="87898-170">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="87898-170">
        - TextBindings</span></span><br><span data-ttu-id="87898-171">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-171">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-172">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="87898-172">Office 2019 on Windows</span></span><br><span data-ttu-id="87898-173">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-173">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="87898-174">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-174">- TaskPane</span></span><br><span data-ttu-id="87898-175">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="87898-175">
        - Content</span></span><br><span data-ttu-id="87898-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="87898-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="87898-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="87898-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="87898-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="87898-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="87898-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="87898-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="87898-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="87898-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="87898-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="87898-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="87898-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="87898-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="87898-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="87898-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="87898-187">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="87898-187">- BindingEvents</span></span><br><span data-ttu-id="87898-188">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-188">
        - CompressedFile</span></span><br><span data-ttu-id="87898-189">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-189">
        - DocumentEvents</span></span><br><span data-ttu-id="87898-190">
        - File</span><span class="sxs-lookup"><span data-stu-id="87898-190">
        - File</span></span><br><span data-ttu-id="87898-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="87898-191">
        - MatrixBindings</span></span><br><span data-ttu-id="87898-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="87898-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-193">
        - Selection</span></span><br><span data-ttu-id="87898-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-194">
        - Settings</span></span><br><span data-ttu-id="87898-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="87898-195">
        - TableBindings</span></span><br><span data-ttu-id="87898-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-196">
        - TableCoercion</span></span><br><span data-ttu-id="87898-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="87898-197">
        - TextBindings</span></span><br><span data-ttu-id="87898-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-199">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="87898-199">Office 2016 on Windows</span></span><br><span data-ttu-id="87898-200">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-200">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="87898-201">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-201">- TaskPane</span></span><br><span data-ttu-id="87898-202">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="87898-202">
        - Content</span></span></td>
    <td><span data-ttu-id="87898-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="87898-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="87898-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="87898-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="87898-206">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="87898-206">- BindingEvents</span></span><br><span data-ttu-id="87898-207">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-207">
        - CompressedFile</span></span><br><span data-ttu-id="87898-208">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-208">
        - DocumentEvents</span></span><br><span data-ttu-id="87898-209">
        - File</span><span class="sxs-lookup"><span data-stu-id="87898-209">
        - File</span></span><br><span data-ttu-id="87898-210">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="87898-210">
        - MatrixBindings</span></span><br><span data-ttu-id="87898-211">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-211">
        - MatrixCoercion</span></span><br><span data-ttu-id="87898-212">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-212">
        - Selection</span></span><br><span data-ttu-id="87898-213">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-213">
        - Settings</span></span><br><span data-ttu-id="87898-214">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="87898-214">
        - TableBindings</span></span><br><span data-ttu-id="87898-215">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-215">
        - TableCoercion</span></span><br><span data-ttu-id="87898-216">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="87898-216">
        - TextBindings</span></span><br><span data-ttu-id="87898-217">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-217">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-218">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="87898-218">Office 2013 on Windows</span></span><br><span data-ttu-id="87898-219">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-219">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="87898-220">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-220">
        - TaskPane</span></span><br><span data-ttu-id="87898-221">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="87898-221">
        - Content</span></span></td>
    <td>  <span data-ttu-id="87898-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="87898-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="87898-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="87898-224">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="87898-224">
        - BindingEvents</span></span><br><span data-ttu-id="87898-225">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-225">
        - CompressedFile</span></span><br><span data-ttu-id="87898-226">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-226">
        - DocumentEvents</span></span><br><span data-ttu-id="87898-227">
        - File</span><span class="sxs-lookup"><span data-stu-id="87898-227">
        - File</span></span><br><span data-ttu-id="87898-228">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="87898-228">
        - MatrixBindings</span></span><br><span data-ttu-id="87898-229">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-229">
        - MatrixCoercion</span></span><br><span data-ttu-id="87898-230">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-230">
        - Selection</span></span><br><span data-ttu-id="87898-231">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-231">
        - Settings</span></span><br><span data-ttu-id="87898-232">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="87898-232">
        - TableBindings</span></span><br><span data-ttu-id="87898-233">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-233">
        - TableCoercion</span></span><br><span data-ttu-id="87898-234">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="87898-234">
        - TextBindings</span></span><br><span data-ttu-id="87898-235">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-235">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-236">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="87898-236">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="87898-237">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="87898-237">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="87898-238">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-238">- TaskPane</span></span><br><span data-ttu-id="87898-239">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="87898-239">
        - Content</span></span><br><span data-ttu-id="87898-240">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="87898-240">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="87898-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="87898-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="87898-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="87898-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="87898-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="87898-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="87898-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="87898-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="87898-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="87898-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="87898-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="87898-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="87898-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="87898-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="87898-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="87898-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="87898-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="87898-252">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="87898-252">- BindingEvents</span></span><br><span data-ttu-id="87898-253">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-253">
        - DocumentEvents</span></span><br><span data-ttu-id="87898-254">
        - File</span><span class="sxs-lookup"><span data-stu-id="87898-254">
        - File</span></span><br><span data-ttu-id="87898-255">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="87898-255">
        - MatrixBindings</span></span><br><span data-ttu-id="87898-256">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-256">
        - MatrixCoercion</span></span><br><span data-ttu-id="87898-257">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-257">
        - Selection</span></span><br><span data-ttu-id="87898-258">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-258">
        - Settings</span></span><br><span data-ttu-id="87898-259">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="87898-259">
        - TableBindings</span></span><br><span data-ttu-id="87898-260">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-260">
        - TableCoercion</span></span><br><span data-ttu-id="87898-261">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="87898-261">
        - TextBindings</span></span><br><span data-ttu-id="87898-262">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-262">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-263">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="87898-263">Office apps on Mac</span></span><br><span data-ttu-id="87898-264">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="87898-264">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="87898-265">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-265">- TaskPane</span></span><br><span data-ttu-id="87898-266">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="87898-266">
        - Content</span></span><br><span data-ttu-id="87898-267">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="87898-267">
        - Custom Functions</span></span><br><span data-ttu-id="87898-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="87898-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="87898-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="87898-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="87898-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="87898-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="87898-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="87898-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="87898-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="87898-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="87898-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="87898-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="87898-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="87898-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="87898-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="87898-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="87898-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="87898-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="87898-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="87898-281">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="87898-281">- BindingEvents</span></span><br><span data-ttu-id="87898-282">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-282">
        - CompressedFile</span></span><br><span data-ttu-id="87898-283">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-283">
        - DocumentEvents</span></span><br><span data-ttu-id="87898-284">
        - File</span><span class="sxs-lookup"><span data-stu-id="87898-284">
        - File</span></span><br><span data-ttu-id="87898-285">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="87898-285">
        - MatrixBindings</span></span><br><span data-ttu-id="87898-286">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-286">
        - MatrixCoercion</span></span><br><span data-ttu-id="87898-287">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-287">
        - PdfFile</span></span><br><span data-ttu-id="87898-288">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-288">
        - Selection</span></span><br><span data-ttu-id="87898-289">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-289">
        - Settings</span></span><br><span data-ttu-id="87898-290">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="87898-290">
        - TableBindings</span></span><br><span data-ttu-id="87898-291">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-291">
        - TableCoercion</span></span><br><span data-ttu-id="87898-292">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="87898-292">
        - TextBindings</span></span><br><span data-ttu-id="87898-293">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-293">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-294">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="87898-294">Office 2019 for Mac</span></span><br><span data-ttu-id="87898-295">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-295">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="87898-296">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-296">- TaskPane</span></span><br><span data-ttu-id="87898-297">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="87898-297">
        - Content</span></span><br><span data-ttu-id="87898-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="87898-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="87898-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="87898-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="87898-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="87898-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="87898-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="87898-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="87898-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="87898-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="87898-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="87898-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="87898-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="87898-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="87898-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="87898-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="87898-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="87898-309">- BindingEvents</span></span><br><span data-ttu-id="87898-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-310">
        - CompressedFile</span></span><br><span data-ttu-id="87898-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-311">
        - DocumentEvents</span></span><br><span data-ttu-id="87898-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="87898-312">
        - File</span></span><br><span data-ttu-id="87898-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="87898-313">
        - MatrixBindings</span></span><br><span data-ttu-id="87898-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="87898-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-315">
        - PdfFile</span></span><br><span data-ttu-id="87898-316">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-316">
        - Selection</span></span><br><span data-ttu-id="87898-317">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-317">
        - Settings</span></span><br><span data-ttu-id="87898-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="87898-318">
        - TableBindings</span></span><br><span data-ttu-id="87898-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-319">
        - TableCoercion</span></span><br><span data-ttu-id="87898-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="87898-320">
        - TextBindings</span></span><br><span data-ttu-id="87898-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-321">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-322">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="87898-322">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="87898-323">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-323">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="87898-324">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-324">- TaskPane</span></span><br><span data-ttu-id="87898-325">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="87898-325">
        - Content</span></span></td>
    <td><span data-ttu-id="87898-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="87898-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="87898-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="87898-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="87898-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="87898-329">- BindingEvents</span></span><br><span data-ttu-id="87898-330">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-330">
        - CompressedFile</span></span><br><span data-ttu-id="87898-331">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-331">
        - DocumentEvents</span></span><br><span data-ttu-id="87898-332">
        - File</span><span class="sxs-lookup"><span data-stu-id="87898-332">
        - File</span></span><br><span data-ttu-id="87898-333">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="87898-333">
        - MatrixBindings</span></span><br><span data-ttu-id="87898-334">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-334">
        - MatrixCoercion</span></span><br><span data-ttu-id="87898-335">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-335">
        - PdfFile</span></span><br><span data-ttu-id="87898-336">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-336">
        - Selection</span></span><br><span data-ttu-id="87898-337">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-337">
        - Settings</span></span><br><span data-ttu-id="87898-338">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="87898-338">
        - TableBindings</span></span><br><span data-ttu-id="87898-339">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-339">
        - TableCoercion</span></span><br><span data-ttu-id="87898-340">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="87898-340">
        - TextBindings</span></span><br><span data-ttu-id="87898-341">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-341">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="87898-342">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="87898-342">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="87898-343">カスタム関数</span><span class="sxs-lookup"><span data-stu-id="87898-343">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="87898-344">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="87898-344">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="87898-345">拡張点</span><span class="sxs-lookup"><span data-stu-id="87898-345">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="87898-346">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="87898-346">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="87898-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="87898-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-348">Office on the web</span><span class="sxs-lookup"><span data-stu-id="87898-348">Office on the web</span></span></td>
    <td><span data-ttu-id="87898-349">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="87898-349">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="87898-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-351">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="87898-351">Office on Windows</span></span><br><span data-ttu-id="87898-352">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="87898-352">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="87898-353">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="87898-353">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="87898-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-355">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="87898-355">Office for Mac</span></span><br><span data-ttu-id="87898-356">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="87898-356">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="87898-357">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="87898-357">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="87898-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="87898-359">Outlook</span><span class="sxs-lookup"><span data-stu-id="87898-359">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="87898-360">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="87898-360">Platform</span></span></th>
    <th><span data-ttu-id="87898-361">拡張点</span><span class="sxs-lookup"><span data-stu-id="87898-361">Extension points</span></span></th>
    <th><span data-ttu-id="87898-362">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="87898-362">API requirement sets</span></span></th>
    <th><span data-ttu-id="87898-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="87898-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-364">Office on the web</span><span class="sxs-lookup"><span data-stu-id="87898-364">Office on the web</span></span><br><span data-ttu-id="87898-365">(モダン)</span><span class="sxs-lookup"><span data-stu-id="87898-365">Modern</span></span></td>
    <td> <span data-ttu-id="87898-366">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="87898-366">- Mail Read</span></span><br><span data-ttu-id="87898-367">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="87898-367">
      - Mail Compose</span></span><br><span data-ttu-id="87898-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="87898-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="87898-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="87898-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="87898-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="87898-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="87898-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="87898-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="87898-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="87898-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="87898-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="87898-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="87898-376">使用不可</span><span class="sxs-lookup"><span data-stu-id="87898-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-377">Office on the web</span><span class="sxs-lookup"><span data-stu-id="87898-377">Office on the web</span></span><br><span data-ttu-id="87898-378">(クラシック)</span><span class="sxs-lookup"><span data-stu-id="87898-378">Classic</span></span></td>
    <td> <span data-ttu-id="87898-379">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="87898-379">- Mail Read</span></span><br><span data-ttu-id="87898-380">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="87898-380">
      - Mail Compose</span></span><br><span data-ttu-id="87898-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="87898-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="87898-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="87898-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="87898-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="87898-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="87898-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="87898-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="87898-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="87898-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="87898-388">使用不可</span><span class="sxs-lookup"><span data-stu-id="87898-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-389">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="87898-389">Office on Windows</span></span><br><span data-ttu-id="87898-390">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="87898-390">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="87898-391">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="87898-391">- Mail Read</span></span><br><span data-ttu-id="87898-392">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="87898-392">
      - Mail Compose</span></span><br><span data-ttu-id="87898-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="87898-394">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="87898-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="87898-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="87898-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="87898-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="87898-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="87898-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="87898-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="87898-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="87898-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="87898-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="87898-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="87898-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="87898-402">使用不可</span><span class="sxs-lookup"><span data-stu-id="87898-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-403">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="87898-403">Office 2019 on Windows</span></span><br><span data-ttu-id="87898-404">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="87898-405">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="87898-405">- Mail Read</span></span><br><span data-ttu-id="87898-406">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="87898-406">
      - Mail Compose</span></span><br><span data-ttu-id="87898-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="87898-408">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="87898-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="87898-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="87898-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="87898-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="87898-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="87898-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="87898-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="87898-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="87898-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="87898-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="87898-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="87898-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="87898-416">使用不可</span><span class="sxs-lookup"><span data-stu-id="87898-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-417">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="87898-417">Office 2016 on Windows</span></span><br><span data-ttu-id="87898-418">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="87898-419">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="87898-419">- Mail Read</span></span><br><span data-ttu-id="87898-420">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="87898-420">
      - Mail Compose</span></span><br><span data-ttu-id="87898-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="87898-422">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="87898-422">
      - Modules</span></span></td>
    <td> <span data-ttu-id="87898-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="87898-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="87898-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="87898-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="87898-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="87898-427">使用不可</span><span class="sxs-lookup"><span data-stu-id="87898-427">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-428">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="87898-428">Office 2013 on Windows</span></span><br><span data-ttu-id="87898-429">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-429">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="87898-430">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="87898-430">- Mail Read</span></span><br><span data-ttu-id="87898-431">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="87898-431">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="87898-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="87898-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="87898-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="87898-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="87898-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="87898-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="87898-436">使用不可</span><span class="sxs-lookup"><span data-stu-id="87898-436">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-437">iOS 上の Office</span><span class="sxs-lookup"><span data-stu-id="87898-437">Office apps on iOS</span></span><br><span data-ttu-id="87898-438">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="87898-438">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="87898-439">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="87898-439">- Mail Read</span></span><br><span data-ttu-id="87898-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="87898-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="87898-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="87898-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="87898-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="87898-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="87898-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="87898-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="87898-446">使用不可</span><span class="sxs-lookup"><span data-stu-id="87898-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-447">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="87898-447">Office apps on Mac</span></span><br><span data-ttu-id="87898-448">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="87898-448">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="87898-449">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="87898-449">- Mail Read</span></span><br><span data-ttu-id="87898-450">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="87898-450">
      - Mail Compose</span></span><br><span data-ttu-id="87898-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="87898-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="87898-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="87898-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="87898-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="87898-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="87898-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="87898-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="87898-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="87898-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="87898-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="87898-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="87898-459">使用不可</span><span class="sxs-lookup"><span data-stu-id="87898-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-460">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="87898-460">Office 2019 for Mac</span></span><br><span data-ttu-id="87898-461">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="87898-462">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="87898-462">- Mail Read</span></span><br><span data-ttu-id="87898-463">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="87898-463">
      - Mail Compose</span></span><br><span data-ttu-id="87898-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="87898-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="87898-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="87898-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="87898-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="87898-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="87898-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="87898-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="87898-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="87898-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="87898-471">使用不可</span><span class="sxs-lookup"><span data-stu-id="87898-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-472">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="87898-472">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="87898-473">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="87898-474">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="87898-474">- Mail Read</span></span><br><span data-ttu-id="87898-475">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="87898-475">
      - Mail Compose</span></span><br><span data-ttu-id="87898-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="87898-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="87898-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="87898-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="87898-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="87898-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="87898-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="87898-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="87898-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="87898-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="87898-483">使用不可</span><span class="sxs-lookup"><span data-stu-id="87898-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-484">Android 上の Office</span><span class="sxs-lookup"><span data-stu-id="87898-484">Office apps on Android</span></span><br><span data-ttu-id="87898-485">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="87898-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="87898-486">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="87898-486">- Mail Read</span></span><br><span data-ttu-id="87898-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="87898-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="87898-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="87898-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="87898-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="87898-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="87898-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="87898-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="87898-493">利用不可</span><span class="sxs-lookup"><span data-stu-id="87898-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="87898-494">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="87898-494">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="87898-495">Word</span><span class="sxs-lookup"><span data-stu-id="87898-495">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="87898-496">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="87898-496">Platform</span></span></th>
    <th><span data-ttu-id="87898-497">拡張点</span><span class="sxs-lookup"><span data-stu-id="87898-497">Extension points</span></span></th>
    <th><span data-ttu-id="87898-498">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="87898-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="87898-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="87898-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-500">Office on the web</span><span class="sxs-lookup"><span data-stu-id="87898-500">Office on the web</span></span></td>
    <td> <span data-ttu-id="87898-501">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-501">- TaskPane</span></span><br><span data-ttu-id="87898-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="87898-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="87898-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="87898-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="87898-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="87898-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="87898-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="87898-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="87898-509">- BindingEvents</span></span><br><span data-ttu-id="87898-510">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="87898-510">
         - CustomXmlParts</span></span><br><span data-ttu-id="87898-511">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-511">
         - DocumentEvents</span></span><br><span data-ttu-id="87898-512">
         - File</span><span class="sxs-lookup"><span data-stu-id="87898-512">
         - File</span></span><br><span data-ttu-id="87898-513">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-513">
         - HtmlCoercion</span></span><br><span data-ttu-id="87898-514">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="87898-514">
         - MatrixBindings</span></span><br><span data-ttu-id="87898-515">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-515">
         - MatrixCoercion</span></span><br><span data-ttu-id="87898-516">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-516">
         - OoxmlCoercion</span></span><br><span data-ttu-id="87898-517">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-517">
         - PdfFile</span></span><br><span data-ttu-id="87898-518">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-518">
         - Selection</span></span><br><span data-ttu-id="87898-519">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-519">
         - Settings</span></span><br><span data-ttu-id="87898-520">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="87898-520">
         - TableBindings</span></span><br><span data-ttu-id="87898-521">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-521">
         - TableCoercion</span></span><br><span data-ttu-id="87898-522">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="87898-522">
         - TextBindings</span></span><br><span data-ttu-id="87898-523">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-523">
         - TextCoercion</span></span><br><span data-ttu-id="87898-524">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="87898-524">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-525">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="87898-525">Office on Windows</span></span><br><span data-ttu-id="87898-526">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="87898-526">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="87898-527">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-527">- TaskPane</span></span><br><span data-ttu-id="87898-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="87898-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="87898-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="87898-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="87898-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="87898-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="87898-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="87898-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="87898-535">- BindingEvents</span></span><br><span data-ttu-id="87898-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-536">
         - CompressedFile</span></span><br><span data-ttu-id="87898-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="87898-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="87898-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-538">
         - DocumentEvents</span></span><br><span data-ttu-id="87898-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="87898-539">
         - File</span></span><br><span data-ttu-id="87898-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="87898-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="87898-541">
         - MatrixBindings</span></span><br><span data-ttu-id="87898-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="87898-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="87898-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-544">
         - PdfFile</span></span><br><span data-ttu-id="87898-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-545">
         - Selection</span></span><br><span data-ttu-id="87898-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-546">
         - Settings</span></span><br><span data-ttu-id="87898-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="87898-547">
         - TableBindings</span></span><br><span data-ttu-id="87898-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-548">
         - TableCoercion</span></span><br><span data-ttu-id="87898-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="87898-549">
         - TextBindings</span></span><br><span data-ttu-id="87898-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-550">
         - TextCoercion</span></span><br><span data-ttu-id="87898-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="87898-551">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-552">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="87898-552">Office 2019 on Windows</span></span><br><span data-ttu-id="87898-553">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-553">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="87898-554">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-554">- TaskPane</span></span><br><span data-ttu-id="87898-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="87898-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="87898-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="87898-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="87898-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="87898-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="87898-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="87898-561">- BindingEvents</span></span><br><span data-ttu-id="87898-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-562">
         - CompressedFile</span></span><br><span data-ttu-id="87898-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="87898-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="87898-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-564">
         - DocumentEvents</span></span><br><span data-ttu-id="87898-565">
         - File</span><span class="sxs-lookup"><span data-stu-id="87898-565">
         - File</span></span><br><span data-ttu-id="87898-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="87898-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="87898-567">
         - MatrixBindings</span></span><br><span data-ttu-id="87898-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="87898-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="87898-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-570">
         - PdfFile</span></span><br><span data-ttu-id="87898-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-571">
         - Selection</span></span><br><span data-ttu-id="87898-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-572">
         - Settings</span></span><br><span data-ttu-id="87898-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="87898-573">
         - TableBindings</span></span><br><span data-ttu-id="87898-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-574">
         - TableCoercion</span></span><br><span data-ttu-id="87898-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="87898-575">
         - TextBindings</span></span><br><span data-ttu-id="87898-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-576">
         - TextCoercion</span></span><br><span data-ttu-id="87898-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="87898-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-578">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="87898-578">Office 2016 on Windows</span></span><br><span data-ttu-id="87898-579">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="87898-580">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-580">- TaskPane</span></span></td>
    <td> <span data-ttu-id="87898-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="87898-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="87898-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="87898-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="87898-584">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="87898-584">- BindingEvents</span></span><br><span data-ttu-id="87898-585">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-585">
         - CompressedFile</span></span><br><span data-ttu-id="87898-586">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="87898-586">
         - CustomXmlParts</span></span><br><span data-ttu-id="87898-587">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-587">
         - DocumentEvents</span></span><br><span data-ttu-id="87898-588">
         - File</span><span class="sxs-lookup"><span data-stu-id="87898-588">
         - File</span></span><br><span data-ttu-id="87898-589">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-589">
         - HtmlCoercion</span></span><br><span data-ttu-id="87898-590">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="87898-590">
         - MatrixBindings</span></span><br><span data-ttu-id="87898-591">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-591">
         - MatrixCoercion</span></span><br><span data-ttu-id="87898-592">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-592">
         - OoxmlCoercion</span></span><br><span data-ttu-id="87898-593">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-593">
         - PdfFile</span></span><br><span data-ttu-id="87898-594">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-594">
         - Selection</span></span><br><span data-ttu-id="87898-595">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-595">
         - Settings</span></span><br><span data-ttu-id="87898-596">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="87898-596">
         - TableBindings</span></span><br><span data-ttu-id="87898-597">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-597">
         - TableCoercion</span></span><br><span data-ttu-id="87898-598">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="87898-598">
         - TextBindings</span></span><br><span data-ttu-id="87898-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-599">
         - TextCoercion</span></span><br><span data-ttu-id="87898-600">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="87898-600">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-601">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="87898-601">Office 2013 on Windows</span></span><br><span data-ttu-id="87898-602">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-602">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="87898-603">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-603">- TaskPane</span></span></td>
    <td> <span data-ttu-id="87898-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="87898-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="87898-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="87898-606">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="87898-606">- BindingEvents</span></span><br><span data-ttu-id="87898-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-607">
         - CompressedFile</span></span><br><span data-ttu-id="87898-608">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="87898-608">
         - CustomXmlParts</span></span><br><span data-ttu-id="87898-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-609">
         - DocumentEvents</span></span><br><span data-ttu-id="87898-610">
         - File</span><span class="sxs-lookup"><span data-stu-id="87898-610">
         - File</span></span><br><span data-ttu-id="87898-611">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-611">
         - HtmlCoercion</span></span><br><span data-ttu-id="87898-612">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="87898-612">
         - MatrixBindings</span></span><br><span data-ttu-id="87898-613">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-613">
         - MatrixCoercion</span></span><br><span data-ttu-id="87898-614">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-614">
         - OoxmlCoercion</span></span><br><span data-ttu-id="87898-615">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-615">
         - PdfFile</span></span><br><span data-ttu-id="87898-616">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-616">
         - Selection</span></span><br><span data-ttu-id="87898-617">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-617">
         - Settings</span></span><br><span data-ttu-id="87898-618">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="87898-618">
         - TableBindings</span></span><br><span data-ttu-id="87898-619">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-619">
         - TableCoercion</span></span><br><span data-ttu-id="87898-620">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="87898-620">
         - TextBindings</span></span><br><span data-ttu-id="87898-621">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-621">
         - TextCoercion</span></span><br><span data-ttu-id="87898-622">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="87898-622">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-623">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="87898-623">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="87898-624">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="87898-624">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="87898-625">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-625">- TaskPane</span></span></td>
    <td> <span data-ttu-id="87898-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="87898-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="87898-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="87898-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="87898-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="87898-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="87898-631">- BindingEvents</span></span><br><span data-ttu-id="87898-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-632">
         - CompressedFile</span></span><br><span data-ttu-id="87898-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="87898-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="87898-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-634">
         - DocumentEvents</span></span><br><span data-ttu-id="87898-635">
         - File</span><span class="sxs-lookup"><span data-stu-id="87898-635">
         - File</span></span><br><span data-ttu-id="87898-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="87898-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="87898-637">
         - MatrixBindings</span></span><br><span data-ttu-id="87898-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="87898-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="87898-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-640">
         - PdfFile</span></span><br><span data-ttu-id="87898-641">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-641">
         - Selection</span></span><br><span data-ttu-id="87898-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-642">
         - Settings</span></span><br><span data-ttu-id="87898-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="87898-643">
         - TableBindings</span></span><br><span data-ttu-id="87898-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-644">
         - TableCoercion</span></span><br><span data-ttu-id="87898-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="87898-645">
         - TextBindings</span></span><br><span data-ttu-id="87898-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-646">
         - TextCoercion</span></span><br><span data-ttu-id="87898-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="87898-647">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-648">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="87898-648">Office apps on Mac</span></span><br><span data-ttu-id="87898-649">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="87898-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="87898-650">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-650">- TaskPane</span></span><br><span data-ttu-id="87898-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="87898-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="87898-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="87898-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="87898-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="87898-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="87898-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="87898-658">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="87898-658">- BindingEvents</span></span><br><span data-ttu-id="87898-659">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-659">
         - CompressedFile</span></span><br><span data-ttu-id="87898-660">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="87898-660">
         - CustomXmlParts</span></span><br><span data-ttu-id="87898-661">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-661">
         - DocumentEvents</span></span><br><span data-ttu-id="87898-662">
         - File</span><span class="sxs-lookup"><span data-stu-id="87898-662">
         - File</span></span><br><span data-ttu-id="87898-663">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-663">
         - HtmlCoercion</span></span><br><span data-ttu-id="87898-664">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="87898-664">
         - MatrixBindings</span></span><br><span data-ttu-id="87898-665">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-665">
         - MatrixCoercion</span></span><br><span data-ttu-id="87898-666">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-666">
         - OoxmlCoercion</span></span><br><span data-ttu-id="87898-667">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-667">
         - PdfFile</span></span><br><span data-ttu-id="87898-668">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-668">
         - Selection</span></span><br><span data-ttu-id="87898-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-669">
         - Settings</span></span><br><span data-ttu-id="87898-670">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="87898-670">
         - TableBindings</span></span><br><span data-ttu-id="87898-671">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-671">
         - TableCoercion</span></span><br><span data-ttu-id="87898-672">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="87898-672">
         - TextBindings</span></span><br><span data-ttu-id="87898-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-673">
         - TextCoercion</span></span><br><span data-ttu-id="87898-674">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="87898-674">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-675">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="87898-675">Office 2019 for Mac</span></span><br><span data-ttu-id="87898-676">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-676">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="87898-677">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-677">- TaskPane</span></span><br><span data-ttu-id="87898-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="87898-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="87898-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="87898-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="87898-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="87898-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="87898-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="87898-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="87898-684">- BindingEvents</span></span><br><span data-ttu-id="87898-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-685">
         - CompressedFile</span></span><br><span data-ttu-id="87898-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="87898-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="87898-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-687">
         - DocumentEvents</span></span><br><span data-ttu-id="87898-688">
         - File</span><span class="sxs-lookup"><span data-stu-id="87898-688">
         - File</span></span><br><span data-ttu-id="87898-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="87898-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="87898-690">
         - MatrixBindings</span></span><br><span data-ttu-id="87898-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="87898-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="87898-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-693">
         - PdfFile</span></span><br><span data-ttu-id="87898-694">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-694">
         - Selection</span></span><br><span data-ttu-id="87898-695">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-695">
         - Settings</span></span><br><span data-ttu-id="87898-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="87898-696">
         - TableBindings</span></span><br><span data-ttu-id="87898-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-697">
         - TableCoercion</span></span><br><span data-ttu-id="87898-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="87898-698">
         - TextBindings</span></span><br><span data-ttu-id="87898-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-699">
         - TextCoercion</span></span><br><span data-ttu-id="87898-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="87898-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-701">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="87898-701">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="87898-702">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="87898-703">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-703">- TaskPane</span></span></td>
    <td> <span data-ttu-id="87898-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="87898-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="87898-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="87898-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="87898-707">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="87898-707">- BindingEvents</span></span><br><span data-ttu-id="87898-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-708">
         - CompressedFile</span></span><br><span data-ttu-id="87898-709">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="87898-709">
         - CustomXmlParts</span></span><br><span data-ttu-id="87898-710">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-710">
         - DocumentEvents</span></span><br><span data-ttu-id="87898-711">
         - File</span><span class="sxs-lookup"><span data-stu-id="87898-711">
         - File</span></span><br><span data-ttu-id="87898-712">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-712">
         - HtmlCoercion</span></span><br><span data-ttu-id="87898-713">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="87898-713">
         - MatrixBindings</span></span><br><span data-ttu-id="87898-714">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-714">
         - MatrixCoercion</span></span><br><span data-ttu-id="87898-715">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-715">
         - OoxmlCoercion</span></span><br><span data-ttu-id="87898-716">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-716">
         - PdfFile</span></span><br><span data-ttu-id="87898-717">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-717">
         - Selection</span></span><br><span data-ttu-id="87898-718">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-718">
         - Settings</span></span><br><span data-ttu-id="87898-719">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="87898-719">
         - TableBindings</span></span><br><span data-ttu-id="87898-720">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-720">
         - TableCoercion</span></span><br><span data-ttu-id="87898-721">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="87898-721">
         - TextBindings</span></span><br><span data-ttu-id="87898-722">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-722">
         - TextCoercion</span></span><br><span data-ttu-id="87898-723">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="87898-723">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="87898-724">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="87898-724">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="87898-725">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="87898-725">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="87898-726">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="87898-726">Platform</span></span></th>
    <th><span data-ttu-id="87898-727">拡張点</span><span class="sxs-lookup"><span data-stu-id="87898-727">Extension points</span></span></th>
    <th><span data-ttu-id="87898-728">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="87898-728">API requirement sets</span></span></th>
    <th><span data-ttu-id="87898-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="87898-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-730">Office on the web</span><span class="sxs-lookup"><span data-stu-id="87898-730">Office on the web</span></span></td>
    <td> <span data-ttu-id="87898-731">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="87898-731">- Content</span></span><br><span data-ttu-id="87898-732">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-732">
         - TaskPane</span></span><br><span data-ttu-id="87898-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="87898-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="87898-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="87898-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="87898-737">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="87898-737">- ActiveView</span></span><br><span data-ttu-id="87898-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-738">
         - CompressedFile</span></span><br><span data-ttu-id="87898-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-739">
         - DocumentEvents</span></span><br><span data-ttu-id="87898-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="87898-740">
         - File</span></span><br><span data-ttu-id="87898-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-741">
         - PdfFile</span></span><br><span data-ttu-id="87898-742">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-742">
         - Selection</span></span><br><span data-ttu-id="87898-743">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-743">
         - Settings</span></span><br><span data-ttu-id="87898-744">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-744">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-745">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="87898-745">Office on Windows</span></span><br><span data-ttu-id="87898-746">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="87898-746">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="87898-747">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="87898-747">- Content</span></span><br><span data-ttu-id="87898-748">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-748">
         - TaskPane</span></span><br><span data-ttu-id="87898-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="87898-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="87898-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="87898-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="87898-753">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="87898-753">- ActiveView</span></span><br><span data-ttu-id="87898-754">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-754">
         - CompressedFile</span></span><br><span data-ttu-id="87898-755">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-755">
         - DocumentEvents</span></span><br><span data-ttu-id="87898-756">
         - File</span><span class="sxs-lookup"><span data-stu-id="87898-756">
         - File</span></span><br><span data-ttu-id="87898-757">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-757">
         - PdfFile</span></span><br><span data-ttu-id="87898-758">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-758">
         - Selection</span></span><br><span data-ttu-id="87898-759">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-759">
         - Settings</span></span><br><span data-ttu-id="87898-760">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-760">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-761">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="87898-761">Office 2019 on Windows</span></span><br><span data-ttu-id="87898-762">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-762">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="87898-763">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="87898-763">- Content</span></span><br><span data-ttu-id="87898-764">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-764">
         - TaskPane</span></span><br><span data-ttu-id="87898-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="87898-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="87898-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="87898-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="87898-768">- ActiveView</span></span><br><span data-ttu-id="87898-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-769">
         - CompressedFile</span></span><br><span data-ttu-id="87898-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-770">
         - DocumentEvents</span></span><br><span data-ttu-id="87898-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="87898-771">
         - File</span></span><br><span data-ttu-id="87898-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-772">
         - PdfFile</span></span><br><span data-ttu-id="87898-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-773">
         - Selection</span></span><br><span data-ttu-id="87898-774">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-774">
         - Settings</span></span><br><span data-ttu-id="87898-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-776">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="87898-776">Office 2016 on Windows</span></span><br><span data-ttu-id="87898-777">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-777">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="87898-778">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="87898-778">- Content</span></span><br><span data-ttu-id="87898-779">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-779">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="87898-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="87898-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="87898-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="87898-782">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="87898-782">- ActiveView</span></span><br><span data-ttu-id="87898-783">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-783">
         - CompressedFile</span></span><br><span data-ttu-id="87898-784">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-784">
         - DocumentEvents</span></span><br><span data-ttu-id="87898-785">
         - File</span><span class="sxs-lookup"><span data-stu-id="87898-785">
         - File</span></span><br><span data-ttu-id="87898-786">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-786">
         - PdfFile</span></span><br><span data-ttu-id="87898-787">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-787">
         - Selection</span></span><br><span data-ttu-id="87898-788">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-788">
         - Settings</span></span><br><span data-ttu-id="87898-789">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-789">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-790">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="87898-790">Office 2013 on Windows</span></span><br><span data-ttu-id="87898-791">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-791">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="87898-792">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="87898-792">- Content</span></span><br><span data-ttu-id="87898-793">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-793">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="87898-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="87898-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="87898-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="87898-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="87898-796">- ActiveView</span></span><br><span data-ttu-id="87898-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-797">
         - CompressedFile</span></span><br><span data-ttu-id="87898-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-798">
         - DocumentEvents</span></span><br><span data-ttu-id="87898-799">
         - File</span><span class="sxs-lookup"><span data-stu-id="87898-799">
         - File</span></span><br><span data-ttu-id="87898-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-800">
         - PdfFile</span></span><br><span data-ttu-id="87898-801">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-801">
         - Selection</span></span><br><span data-ttu-id="87898-802">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-802">
         - Settings</span></span><br><span data-ttu-id="87898-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-804">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="87898-804">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="87898-805">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="87898-805">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="87898-806">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="87898-806">- Content</span></span><br><span data-ttu-id="87898-807">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="87898-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="87898-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="87898-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="87898-810">- ActiveView</span></span><br><span data-ttu-id="87898-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-811">
         - CompressedFile</span></span><br><span data-ttu-id="87898-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-812">
         - DocumentEvents</span></span><br><span data-ttu-id="87898-813">
         - File</span><span class="sxs-lookup"><span data-stu-id="87898-813">
         - File</span></span><br><span data-ttu-id="87898-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-814">
         - PdfFile</span></span><br><span data-ttu-id="87898-815">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-815">
         - Selection</span></span><br><span data-ttu-id="87898-816">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-816">
         - Settings</span></span><br><span data-ttu-id="87898-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-818">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="87898-818">Office apps on Mac</span></span><br><span data-ttu-id="87898-819">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="87898-819">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="87898-820">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="87898-820">- Content</span></span><br><span data-ttu-id="87898-821">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-821">
         - TaskPane</span></span><br><span data-ttu-id="87898-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="87898-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="87898-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="87898-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="87898-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="87898-826">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="87898-826">- ActiveView</span></span><br><span data-ttu-id="87898-827">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-827">
         - CompressedFile</span></span><br><span data-ttu-id="87898-828">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-828">
         - DocumentEvents</span></span><br><span data-ttu-id="87898-829">
         - File</span><span class="sxs-lookup"><span data-stu-id="87898-829">
         - File</span></span><br><span data-ttu-id="87898-830">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-830">
         - PdfFile</span></span><br><span data-ttu-id="87898-831">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-831">
         - Selection</span></span><br><span data-ttu-id="87898-832">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-832">
         - Settings</span></span><br><span data-ttu-id="87898-833">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-833">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-834">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="87898-834">Office 2019 for Mac</span></span><br><span data-ttu-id="87898-835">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-835">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="87898-836">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="87898-836">- Content</span></span><br><span data-ttu-id="87898-837">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-837">
         - TaskPane</span></span><br><span data-ttu-id="87898-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="87898-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="87898-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="87898-841">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="87898-841">- ActiveView</span></span><br><span data-ttu-id="87898-842">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-842">
         - CompressedFile</span></span><br><span data-ttu-id="87898-843">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-843">
         - DocumentEvents</span></span><br><span data-ttu-id="87898-844">
         - File</span><span class="sxs-lookup"><span data-stu-id="87898-844">
         - File</span></span><br><span data-ttu-id="87898-845">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-845">
         - PdfFile</span></span><br><span data-ttu-id="87898-846">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-846">
         - Selection</span></span><br><span data-ttu-id="87898-847">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-847">
         - Settings</span></span><br><span data-ttu-id="87898-848">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-848">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-849">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="87898-849">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="87898-850">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-850">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="87898-851">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="87898-851">- Content</span></span><br><span data-ttu-id="87898-852">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-852">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="87898-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="87898-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="87898-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="87898-855">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="87898-855">- ActiveView</span></span><br><span data-ttu-id="87898-856">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="87898-856">
         - CompressedFile</span></span><br><span data-ttu-id="87898-857">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-857">
         - DocumentEvents</span></span><br><span data-ttu-id="87898-858">
         - File</span><span class="sxs-lookup"><span data-stu-id="87898-858">
         - File</span></span><br><span data-ttu-id="87898-859">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="87898-859">
         - PdfFile</span></span><br><span data-ttu-id="87898-860">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="87898-860">
         - Selection</span></span><br><span data-ttu-id="87898-861">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-861">
         - Settings</span></span><br><span data-ttu-id="87898-862">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-862">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="87898-863">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="87898-863">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="87898-864">OneNote</span><span class="sxs-lookup"><span data-stu-id="87898-864">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="87898-865">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="87898-865">Platform</span></span></th>
    <th><span data-ttu-id="87898-866">拡張点</span><span class="sxs-lookup"><span data-stu-id="87898-866">Extension points</span></span></th>
    <th><span data-ttu-id="87898-867">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="87898-867">API requirement sets</span></span></th>
    <th><span data-ttu-id="87898-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="87898-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-869">Office on the web</span><span class="sxs-lookup"><span data-stu-id="87898-869">Office on the web</span></span></td>
    <td> <span data-ttu-id="87898-870">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="87898-870">- Content</span></span><br><span data-ttu-id="87898-871">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-871">
         - TaskPane</span></span><br><span data-ttu-id="87898-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="87898-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="87898-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="87898-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="87898-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="87898-876">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="87898-876">- DocumentEvents</span></span><br><span data-ttu-id="87898-877">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-877">
         - HtmlCoercion</span></span><br><span data-ttu-id="87898-878">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="87898-878">
         - Settings</span></span><br><span data-ttu-id="87898-879">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-879">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="87898-880">Project</span><span class="sxs-lookup"><span data-stu-id="87898-880">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="87898-881">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="87898-881">Platform</span></span></th>
    <th><span data-ttu-id="87898-882">拡張点</span><span class="sxs-lookup"><span data-stu-id="87898-882">Extension points</span></span></th>
    <th><span data-ttu-id="87898-883">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="87898-883">API requirement sets</span></span></th>
    <th><span data-ttu-id="87898-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="87898-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-885">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="87898-885">Office 2019 on Windows</span></span><br><span data-ttu-id="87898-886">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-886">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="87898-887">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-887">- TaskPane</span></span></td>
    <td> <span data-ttu-id="87898-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="87898-889">- Selection</span><span class="sxs-lookup"><span data-stu-id="87898-889">- Selection</span></span><br><span data-ttu-id="87898-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-890">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-891">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="87898-891">Office 2016 on Windows</span></span><br><span data-ttu-id="87898-892">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-892">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="87898-893">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-893">- TaskPane</span></span></td>
    <td> <span data-ttu-id="87898-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="87898-895">- Selection</span><span class="sxs-lookup"><span data-stu-id="87898-895">- Selection</span></span><br><span data-ttu-id="87898-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-896">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="87898-897">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="87898-897">Office 2013 on Windows</span></span><br><span data-ttu-id="87898-898">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="87898-898">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="87898-899">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="87898-899">- TaskPane</span></span></td>
    <td> <span data-ttu-id="87898-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="87898-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="87898-901">- Selection</span><span class="sxs-lookup"><span data-stu-id="87898-901">- Selection</span></span><br><span data-ttu-id="87898-902">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="87898-902">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="87898-903">関連項目</span><span class="sxs-lookup"><span data-stu-id="87898-903">See also</span></span>

- [<span data-ttu-id="87898-904">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="87898-904">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="87898-905">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="87898-905">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="87898-906">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="87898-906">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="87898-907">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="87898-907">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="87898-908">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="87898-908">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="87898-909">Office 365 ProPlus の更新履歴</span><span class="sxs-lookup"><span data-stu-id="87898-909">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="87898-910">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="87898-910">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="87898-911">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="87898-911">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="87898-912">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="87898-912">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="87898-913">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="87898-913">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="87898-914">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="87898-914">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
