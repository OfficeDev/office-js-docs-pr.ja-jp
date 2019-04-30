---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、Word、Outlook、PowerPoint、OneNote、および Project のサポートされる要件セット。
ms.date: 04/03/2019
localization_priority: Priority
ms.openlocfilehash: a9ecd44edf9221a403eb42756cd1e9f5e676ad01
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448148"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="dee93-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dee93-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="dee93-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="dee93-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="dee93-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="dee93-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="dee93-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="dee93-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="dee93-108">Excel</span><span class="sxs-lookup"><span data-stu-id="dee93-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="dee93-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dee93-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="dee93-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="dee93-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="dee93-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="dee93-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="dee93-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="dee93-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="dee93-113">Office Online</span></span></td>
    <td> <span data-ttu-id="dee93-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-114">- TaskPane</span></span><br><span data-ttu-id="dee93-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dee93-115">
        - Content</span></span><br><span data-ttu-id="dee93-116">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="dee93-116">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="dee93-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dee93-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dee93-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dee93-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dee93-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dee93-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dee93-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dee93-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dee93-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dee93-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dee93-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dee93-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dee93-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dee93-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="dee93-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-126">
        - BindingEvents</span></span><br><span data-ttu-id="dee93-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-127">
        - CompressedFile</span></span><br><span data-ttu-id="dee93-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-128">
        - DocumentEvents</span></span><br><span data-ttu-id="dee93-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="dee93-129">
        - File</span></span><br><span data-ttu-id="dee93-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-130">
        - MatrixBindings</span></span><br><span data-ttu-id="dee93-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="dee93-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-132">
        - Selection</span></span><br><span data-ttu-id="dee93-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-133">
        - Settings</span></span><br><span data-ttu-id="dee93-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-134">
        - TableBindings</span></span><br><span data-ttu-id="dee93-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-135">
        - TableCoercion</span></span><br><span data-ttu-id="dee93-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-136">
        - TextBindings</span></span><br><span data-ttu-id="dee93-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-138">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="dee93-138">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="dee93-139">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-139">- TaskPane</span></span><br><span data-ttu-id="dee93-140">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dee93-140">
        - Content</span></span><br><span data-ttu-id="dee93-141">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="dee93-141">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="dee93-142">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-142">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dee93-143">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-143">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dee93-144">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-144">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dee93-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dee93-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dee93-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dee93-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dee93-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dee93-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dee93-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dee93-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dee93-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dee93-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dee93-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="dee93-151">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-151">
        - BindingEvents</span></span><br><span data-ttu-id="dee93-152">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-152">
        - CompressedFile</span></span><br><span data-ttu-id="dee93-153">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-153">
        - DocumentEvents</span></span><br><span data-ttu-id="dee93-154">
        - File</span><span class="sxs-lookup"><span data-stu-id="dee93-154">
        - File</span></span><br><span data-ttu-id="dee93-155">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-155">
        - MatrixBindings</span></span><br><span data-ttu-id="dee93-156">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-156">
        - MatrixCoercion</span></span><br><span data-ttu-id="dee93-157">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-157">
        - Selection</span></span><br><span data-ttu-id="dee93-158">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-158">
        - Settings</span></span><br><span data-ttu-id="dee93-159">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-159">
        - TableBindings</span></span><br><span data-ttu-id="dee93-160">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-160">
        - TableCoercion</span></span><br><span data-ttu-id="dee93-161">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-161">
        - TextBindings</span></span><br><span data-ttu-id="dee93-162">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-162">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-163">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="dee93-163">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="dee93-164">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-164">- TaskPane</span></span><br><span data-ttu-id="dee93-165">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dee93-165">
        - Content</span></span><br><span data-ttu-id="dee93-166">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-166">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="dee93-167">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-167">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dee93-168">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-168">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dee93-169">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-169">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dee93-170">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dee93-170">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dee93-171">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dee93-171">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dee93-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dee93-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dee93-173">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dee93-173">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dee93-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dee93-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dee93-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="dee93-176">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-176">- BindingEvents</span></span><br><span data-ttu-id="dee93-177">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-177">
        - CompressedFile</span></span><br><span data-ttu-id="dee93-178">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-178">
        - DocumentEvents</span></span><br><span data-ttu-id="dee93-179">
        - File</span><span class="sxs-lookup"><span data-stu-id="dee93-179">
        - File</span></span><br><span data-ttu-id="dee93-180">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-180">
        - ImageCoercion</span></span><br><span data-ttu-id="dee93-181">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-181">
        - MatrixBindings</span></span><br><span data-ttu-id="dee93-182">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-182">
        - MatrixCoercion</span></span><br><span data-ttu-id="dee93-183">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-183">
        - Selection</span></span><br><span data-ttu-id="dee93-184">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-184">
        - Settings</span></span><br><span data-ttu-id="dee93-185">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-185">
        - TableBindings</span></span><br><span data-ttu-id="dee93-186">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-186">
        - TableCoercion</span></span><br><span data-ttu-id="dee93-187">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-187">
        - TextBindings</span></span><br><span data-ttu-id="dee93-188">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-188">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-189">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="dee93-189">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="dee93-190">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-190">- TaskPane</span></span><br><span data-ttu-id="dee93-191">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dee93-191">
        - Content</span></span></td>
    <td><span data-ttu-id="dee93-192">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-192">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dee93-193">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="dee93-193">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="dee93-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-194">- BindingEvents</span></span><br><span data-ttu-id="dee93-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-195">
        - CompressedFile</span></span><br><span data-ttu-id="dee93-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-196">
        - DocumentEvents</span></span><br><span data-ttu-id="dee93-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="dee93-197">
        - File</span></span><br><span data-ttu-id="dee93-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-198">
        - ImageCoercion</span></span><br><span data-ttu-id="dee93-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-199">
        - MatrixBindings</span></span><br><span data-ttu-id="dee93-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="dee93-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-201">
        - Selection</span></span><br><span data-ttu-id="dee93-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-202">
        - Settings</span></span><br><span data-ttu-id="dee93-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-203">
        - TableBindings</span></span><br><span data-ttu-id="dee93-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-204">
        - TableCoercion</span></span><br><span data-ttu-id="dee93-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-205">
        - TextBindings</span></span><br><span data-ttu-id="dee93-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-207">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="dee93-207">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="dee93-208">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-208">
        - TaskPane</span></span><br><span data-ttu-id="dee93-209">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dee93-209">
        - Content</span></span></td>
    <td>  <span data-ttu-id="dee93-210">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="dee93-210">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="dee93-211">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-211">
        - BindingEvents</span></span><br><span data-ttu-id="dee93-212">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-212">
        - CompressedFile</span></span><br><span data-ttu-id="dee93-213">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-213">
        - DocumentEvents</span></span><br><span data-ttu-id="dee93-214">
        - File</span><span class="sxs-lookup"><span data-stu-id="dee93-214">
        - File</span></span><br><span data-ttu-id="dee93-215">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-215">
        - ImageCoercion</span></span><br><span data-ttu-id="dee93-216">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-216">
        - MatrixBindings</span></span><br><span data-ttu-id="dee93-217">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-217">
        - MatrixCoercion</span></span><br><span data-ttu-id="dee93-218">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-218">
        - Selection</span></span><br><span data-ttu-id="dee93-219">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-219">
        - Settings</span></span><br><span data-ttu-id="dee93-220">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-220">
        - TableBindings</span></span><br><span data-ttu-id="dee93-221">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-221">
        - TableCoercion</span></span><br><span data-ttu-id="dee93-222">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-222">
        - TextBindings</span></span><br><span data-ttu-id="dee93-223">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-223">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-224">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="dee93-224">Office 365 for iPad</span></span></td>
    <td><span data-ttu-id="dee93-225">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-225">- TaskPane</span></span><br><span data-ttu-id="dee93-226">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dee93-226">
        - Content</span></span></td>
    <td><span data-ttu-id="dee93-227">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-227">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dee93-228">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-228">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dee93-229">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-229">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dee93-230">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dee93-230">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dee93-231">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dee93-231">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dee93-232">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dee93-232">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dee93-233">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dee93-233">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dee93-234">
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dee93-234">
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dee93-235">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-235">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="dee93-236">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-236">- BindingEvents</span></span><br><span data-ttu-id="dee93-237">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-237">
        - CompressedFile</span></span><br><span data-ttu-id="dee93-238">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-238">
        - DocumentEvents</span></span><br><span data-ttu-id="dee93-239">
        - File</span><span class="sxs-lookup"><span data-stu-id="dee93-239">
        - File</span></span><br><span data-ttu-id="dee93-240">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-240">
        - ImageCoercion</span></span><br><span data-ttu-id="dee93-241">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-241">
        - MatrixBindings</span></span><br><span data-ttu-id="dee93-242">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-242">
        - MatrixCoercion</span></span><br><span data-ttu-id="dee93-243">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-243">
        - Selection</span></span><br><span data-ttu-id="dee93-244">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-244">
        - Settings</span></span><br><span data-ttu-id="dee93-245">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-245">
        - TableBindings</span></span><br><span data-ttu-id="dee93-246">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-246">
        - TableCoercion</span></span><br><span data-ttu-id="dee93-247">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-247">
        - TextBindings</span></span><br><span data-ttu-id="dee93-248">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-248">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-249">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="dee93-249">Office 365 for Mac</span></span></td>
    <td><span data-ttu-id="dee93-250">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-250">- TaskPane</span></span><br><span data-ttu-id="dee93-251">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dee93-251">
        - Content</span></span><br><span data-ttu-id="dee93-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="dee93-253">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-253">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dee93-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dee93-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dee93-256">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dee93-256">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dee93-257">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dee93-257">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dee93-258">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dee93-258">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dee93-259">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dee93-259">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dee93-260">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dee93-260">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dee93-261">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-261">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="dee93-262">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-262">- BindingEvents</span></span><br><span data-ttu-id="dee93-263">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-263">
        - CompressedFile</span></span><br><span data-ttu-id="dee93-264">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-264">
        - DocumentEvents</span></span><br><span data-ttu-id="dee93-265">
        - File</span><span class="sxs-lookup"><span data-stu-id="dee93-265">
        - File</span></span><br><span data-ttu-id="dee93-266">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-266">
        - ImageCoercion</span></span><br><span data-ttu-id="dee93-267">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-267">
        - MatrixBindings</span></span><br><span data-ttu-id="dee93-268">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-268">
        - MatrixCoercion</span></span><br><span data-ttu-id="dee93-269">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-269">
        - PdfFile</span></span><br><span data-ttu-id="dee93-270">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-270">
        - Selection</span></span><br><span data-ttu-id="dee93-271">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-271">
        - Settings</span></span><br><span data-ttu-id="dee93-272">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-272">
        - TableBindings</span></span><br><span data-ttu-id="dee93-273">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-273">
        - TableCoercion</span></span><br><span data-ttu-id="dee93-274">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-274">
        - TextBindings</span></span><br><span data-ttu-id="dee93-275">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-275">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-276">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="dee93-276">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="dee93-277">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-277">- TaskPane</span></span><br><span data-ttu-id="dee93-278">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dee93-278">
        - Content</span></span><br><span data-ttu-id="dee93-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="dee93-280">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-280">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dee93-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dee93-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dee93-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dee93-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dee93-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dee93-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dee93-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dee93-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dee93-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dee93-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dee93-287">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dee93-287">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dee93-288">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-288">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="dee93-289">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-289">- BindingEvents</span></span><br><span data-ttu-id="dee93-290">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-290">
        - CompressedFile</span></span><br><span data-ttu-id="dee93-291">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-291">
        - DocumentEvents</span></span><br><span data-ttu-id="dee93-292">
        - File</span><span class="sxs-lookup"><span data-stu-id="dee93-292">
        - File</span></span><br><span data-ttu-id="dee93-293">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-293">
        - ImageCoercion</span></span><br><span data-ttu-id="dee93-294">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-294">
        - MatrixBindings</span></span><br><span data-ttu-id="dee93-295">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-295">
        - MatrixCoercion</span></span><br><span data-ttu-id="dee93-296">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-296">
        - PdfFile</span></span><br><span data-ttu-id="dee93-297">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-297">
        - Selection</span></span><br><span data-ttu-id="dee93-298">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-298">
        - Settings</span></span><br><span data-ttu-id="dee93-299">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-299">
        - TableBindings</span></span><br><span data-ttu-id="dee93-300">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-300">
        - TableCoercion</span></span><br><span data-ttu-id="dee93-301">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-301">
        - TextBindings</span></span><br><span data-ttu-id="dee93-302">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-302">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-303">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="dee93-303">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="dee93-304">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-304">- TaskPane</span></span><br><span data-ttu-id="dee93-305">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dee93-305">
        - Content</span></span></td>
    <td><span data-ttu-id="dee93-306">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-306">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dee93-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="dee93-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="dee93-308">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-308">- BindingEvents</span></span><br><span data-ttu-id="dee93-309">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-309">
        - CompressedFile</span></span><br><span data-ttu-id="dee93-310">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-310">
        - DocumentEvents</span></span><br><span data-ttu-id="dee93-311">
        - File</span><span class="sxs-lookup"><span data-stu-id="dee93-311">
        - File</span></span><br><span data-ttu-id="dee93-312">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-312">
        - ImageCoercion</span></span><br><span data-ttu-id="dee93-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-313">
        - MatrixBindings</span></span><br><span data-ttu-id="dee93-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="dee93-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-315">
        - PdfFile</span></span><br><span data-ttu-id="dee93-316">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-316">
        - Selection</span></span><br><span data-ttu-id="dee93-317">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-317">
        - Settings</span></span><br><span data-ttu-id="dee93-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-318">
        - TableBindings</span></span><br><span data-ttu-id="dee93-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-319">
        - TableCoercion</span></span><br><span data-ttu-id="dee93-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-320">
        - TextBindings</span></span><br><span data-ttu-id="dee93-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-321">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="dee93-322">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="dee93-322">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="dee93-323">Outlook</span><span class="sxs-lookup"><span data-stu-id="dee93-323">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="dee93-324">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dee93-324">Platform</span></span></th>
    <th><span data-ttu-id="dee93-325">拡張点</span><span class="sxs-lookup"><span data-stu-id="dee93-325">Extension points</span></span></th>
    <th><span data-ttu-id="dee93-326">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="dee93-326">API requirement sets</span></span></th>
    <th><span data-ttu-id="dee93-327"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="dee93-327"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-328">Office Online</span><span class="sxs-lookup"><span data-stu-id="dee93-328">Office Online</span></span></td>
    <td> <span data-ttu-id="dee93-329">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dee93-329">- Mail Read</span></span><br><span data-ttu-id="dee93-330">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dee93-330">
      - Mail Compose</span></span><br><span data-ttu-id="dee93-331">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-331">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dee93-332">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-332">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dee93-333">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-333">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dee93-334">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-334">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dee93-335">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dee93-335">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dee93-336">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dee93-336">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dee93-337">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dee93-337">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="dee93-338">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dee93-338">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="dee93-339">利用不可</span><span class="sxs-lookup"><span data-stu-id="dee93-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-340">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="dee93-340">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="dee93-341">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dee93-341">- Mail Read</span></span><br><span data-ttu-id="dee93-342">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dee93-342">
      - Mail Compose</span></span><br><span data-ttu-id="dee93-343">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-343">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="dee93-344">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="dee93-344">
      - Modules</span></span></td>
    <td> <span data-ttu-id="dee93-345">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-345">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dee93-346">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-346">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dee93-347">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-347">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dee93-348">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dee93-348">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dee93-349">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dee93-349">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dee93-350">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dee93-350">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="dee93-351">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dee93-351">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="dee93-352">利用不可</span><span class="sxs-lookup"><span data-stu-id="dee93-352">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-353">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="dee93-353">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="dee93-354">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dee93-354">- Mail Read</span></span><br><span data-ttu-id="dee93-355">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dee93-355">
      - Mail Compose</span></span><br><span data-ttu-id="dee93-356">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-356">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="dee93-357">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="dee93-357">
      - Modules</span></span></td>
    <td> <span data-ttu-id="dee93-358">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-358">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dee93-359">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-359">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dee93-360">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-360">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dee93-361">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dee93-361">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dee93-362">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dee93-362">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dee93-363">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dee93-363">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="dee93-364">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dee93-364">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="dee93-365">利用不可</span><span class="sxs-lookup"><span data-stu-id="dee93-365">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-366">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="dee93-366">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="dee93-367">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dee93-367">- Mail Read</span></span><br><span data-ttu-id="dee93-368">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dee93-368">
      - Mail Compose</span></span><br><span data-ttu-id="dee93-369">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-369">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="dee93-370">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="dee93-370">
      - Modules</span></span></td>
    <td> <span data-ttu-id="dee93-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dee93-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dee93-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dee93-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="dee93-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="dee93-375">利用不可</span><span class="sxs-lookup"><span data-stu-id="dee93-375">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-376">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="dee93-376">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="dee93-377">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dee93-377">- Mail Read</span></span><br><span data-ttu-id="dee93-378">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dee93-378">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="dee93-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dee93-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dee93-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="dee93-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="dee93-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="dee93-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="dee93-383">利用不可</span><span class="sxs-lookup"><span data-stu-id="dee93-383">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-384">Office 365 for iOS</span><span class="sxs-lookup"><span data-stu-id="dee93-384">Office 365 for iOS</span></span></td>
    <td> <span data-ttu-id="dee93-385">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dee93-385">- Mail Read</span></span><br><span data-ttu-id="dee93-386">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-386">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dee93-387">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-387">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dee93-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dee93-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dee93-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dee93-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dee93-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dee93-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="dee93-392">利用不可</span><span class="sxs-lookup"><span data-stu-id="dee93-392">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-393">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="dee93-393">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="dee93-394">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dee93-394">- Mail Read</span></span><br><span data-ttu-id="dee93-395">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dee93-395">
      - Mail Compose</span></span><br><span data-ttu-id="dee93-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dee93-397">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-397">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dee93-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dee93-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dee93-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dee93-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dee93-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dee93-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dee93-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dee93-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="dee93-403">利用不可</span><span class="sxs-lookup"><span data-stu-id="dee93-403">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-404">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="dee93-404">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="dee93-405">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dee93-405">- Mail Read</span></span><br><span data-ttu-id="dee93-406">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dee93-406">
      - Mail Compose</span></span><br><span data-ttu-id="dee93-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dee93-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dee93-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dee93-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dee93-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dee93-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dee93-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dee93-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dee93-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dee93-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="dee93-414">利用不可</span><span class="sxs-lookup"><span data-stu-id="dee93-414">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-415">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="dee93-415">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="dee93-416">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dee93-416">- Mail Read</span></span><br><span data-ttu-id="dee93-417">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dee93-417">
      - Mail Compose</span></span><br><span data-ttu-id="dee93-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dee93-419">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-419">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dee93-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dee93-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dee93-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dee93-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dee93-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dee93-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dee93-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dee93-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="dee93-425">利用不可</span><span class="sxs-lookup"><span data-stu-id="dee93-425">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-426">Office 365 for Android</span><span class="sxs-lookup"><span data-stu-id="dee93-426">Office 365 for Android</span></span></td>
    <td> <span data-ttu-id="dee93-427">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dee93-427">- Mail Read</span></span><br><span data-ttu-id="dee93-428">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-428">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dee93-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dee93-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dee93-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dee93-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dee93-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dee93-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dee93-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="dee93-434">利用不可</span><span class="sxs-lookup"><span data-stu-id="dee93-434">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="dee93-435">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="dee93-435">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="dee93-436">Word</span><span class="sxs-lookup"><span data-stu-id="dee93-436">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="dee93-437">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dee93-437">Platform</span></span></th>
    <th><span data-ttu-id="dee93-438">拡張点</span><span class="sxs-lookup"><span data-stu-id="dee93-438">Extension points</span></span></th>
    <th><span data-ttu-id="dee93-439">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="dee93-439">API requirement sets</span></span></th>
    <th><span data-ttu-id="dee93-440"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="dee93-440"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-441">Office Online</span><span class="sxs-lookup"><span data-stu-id="dee93-441">Office Online</span></span></td>
    <td> <span data-ttu-id="dee93-442">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-442">- TaskPane</span></span><br><span data-ttu-id="dee93-443">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-443">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dee93-444">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-444">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="dee93-445">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-445">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="dee93-446">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-446">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="dee93-447">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-447">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dee93-448">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-448">- BindingEvents</span></span><br><span data-ttu-id="dee93-449">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dee93-449">
         - CustomXmlParts</span></span><br><span data-ttu-id="dee93-450">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-450">
         - DocumentEvents</span></span><br><span data-ttu-id="dee93-451">
         - File</span><span class="sxs-lookup"><span data-stu-id="dee93-451">
         - File</span></span><br><span data-ttu-id="dee93-452">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-452">
         - HtmlCoercion</span></span><br><span data-ttu-id="dee93-453">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-453">
         - ImageCoercion</span></span><br><span data-ttu-id="dee93-454">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-454">
         - MatrixBindings</span></span><br><span data-ttu-id="dee93-455">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-455">
         - MatrixCoercion</span></span><br><span data-ttu-id="dee93-456">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-456">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dee93-457">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-457">
         - PdfFile</span></span><br><span data-ttu-id="dee93-458">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-458">
         - Selection</span></span><br><span data-ttu-id="dee93-459">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-459">
         - Settings</span></span><br><span data-ttu-id="dee93-460">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-460">
         - TableBindings</span></span><br><span data-ttu-id="dee93-461">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-461">
         - TableCoercion</span></span><br><span data-ttu-id="dee93-462">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-462">
         - TextBindings</span></span><br><span data-ttu-id="dee93-463">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-463">
         - TextCoercion</span></span><br><span data-ttu-id="dee93-464">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dee93-464">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-465">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="dee93-465">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="dee93-466">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-466">- TaskPane</span></span><br><span data-ttu-id="dee93-467">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-467">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dee93-468">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-468">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="dee93-469">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-469">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="dee93-470">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-470">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="dee93-471">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-471">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dee93-472">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-472">- BindingEvents</span></span><br><span data-ttu-id="dee93-473">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-473">
         - CompressedFile</span></span><br><span data-ttu-id="dee93-474">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dee93-474">
         - CustomXmlParts</span></span><br><span data-ttu-id="dee93-475">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-475">
         - DocumentEvents</span></span><br><span data-ttu-id="dee93-476">
         - File</span><span class="sxs-lookup"><span data-stu-id="dee93-476">
         - File</span></span><br><span data-ttu-id="dee93-477">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-477">
         - HtmlCoercion</span></span><br><span data-ttu-id="dee93-478">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-478">
         - ImageCoercion</span></span><br><span data-ttu-id="dee93-479">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-479">
         - MatrixBindings</span></span><br><span data-ttu-id="dee93-480">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-480">
         - MatrixCoercion</span></span><br><span data-ttu-id="dee93-481">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-481">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dee93-482">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-482">
         - PdfFile</span></span><br><span data-ttu-id="dee93-483">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-483">
         - Selection</span></span><br><span data-ttu-id="dee93-484">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-484">
         - Settings</span></span><br><span data-ttu-id="dee93-485">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-485">
         - TableBindings</span></span><br><span data-ttu-id="dee93-486">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-486">
         - TableCoercion</span></span><br><span data-ttu-id="dee93-487">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-487">
         - TextBindings</span></span><br><span data-ttu-id="dee93-488">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-488">
         - TextCoercion</span></span><br><span data-ttu-id="dee93-489">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dee93-489">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-490">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="dee93-490">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="dee93-491">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-491">- TaskPane</span></span><br><span data-ttu-id="dee93-492">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-492">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dee93-493">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-493">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="dee93-494">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-494">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="dee93-495">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-495">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="dee93-496">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-496">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dee93-497">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-497">- BindingEvents</span></span><br><span data-ttu-id="dee93-498">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-498">
         - CompressedFile</span></span><br><span data-ttu-id="dee93-499">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dee93-499">
         - CustomXmlParts</span></span><br><span data-ttu-id="dee93-500">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-500">
         - DocumentEvents</span></span><br><span data-ttu-id="dee93-501">
         - File</span><span class="sxs-lookup"><span data-stu-id="dee93-501">
         - File</span></span><br><span data-ttu-id="dee93-502">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-502">
         - HtmlCoercion</span></span><br><span data-ttu-id="dee93-503">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-503">
         - ImageCoercion</span></span><br><span data-ttu-id="dee93-504">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-504">
         - MatrixBindings</span></span><br><span data-ttu-id="dee93-505">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-505">
         - MatrixCoercion</span></span><br><span data-ttu-id="dee93-506">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-506">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dee93-507">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-507">
         - PdfFile</span></span><br><span data-ttu-id="dee93-508">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-508">
         - Selection</span></span><br><span data-ttu-id="dee93-509">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-509">
         - Settings</span></span><br><span data-ttu-id="dee93-510">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-510">
         - TableBindings</span></span><br><span data-ttu-id="dee93-511">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-511">
         - TableCoercion</span></span><br><span data-ttu-id="dee93-512">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-512">
         - TextBindings</span></span><br><span data-ttu-id="dee93-513">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-513">
         - TextCoercion</span></span><br><span data-ttu-id="dee93-514">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dee93-514">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-515">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="dee93-515">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="dee93-516">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-516">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dee93-517">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-517">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="dee93-518">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="dee93-518">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="dee93-519">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-519">- BindingEvents</span></span><br><span data-ttu-id="dee93-520">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-520">
         - CompressedFile</span></span><br><span data-ttu-id="dee93-521">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dee93-521">
         - CustomXmlParts</span></span><br><span data-ttu-id="dee93-522">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-522">
         - DocumentEvents</span></span><br><span data-ttu-id="dee93-523">
         - File</span><span class="sxs-lookup"><span data-stu-id="dee93-523">
         - File</span></span><br><span data-ttu-id="dee93-524">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-524">
         - HtmlCoercion</span></span><br><span data-ttu-id="dee93-525">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-525">
         - ImageCoercion</span></span><br><span data-ttu-id="dee93-526">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-526">
         - MatrixBindings</span></span><br><span data-ttu-id="dee93-527">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-527">
         - MatrixCoercion</span></span><br><span data-ttu-id="dee93-528">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-528">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dee93-529">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-529">
         - PdfFile</span></span><br><span data-ttu-id="dee93-530">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-530">
         - Selection</span></span><br><span data-ttu-id="dee93-531">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-531">
         - Settings</span></span><br><span data-ttu-id="dee93-532">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-532">
         - TableBindings</span></span><br><span data-ttu-id="dee93-533">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-533">
         - TableCoercion</span></span><br><span data-ttu-id="dee93-534">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-534">
         - TextBindings</span></span><br><span data-ttu-id="dee93-535">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-535">
         - TextCoercion</span></span><br><span data-ttu-id="dee93-536">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dee93-536">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-537">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="dee93-537">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="dee93-538">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-538">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dee93-539">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="dee93-539">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="dee93-540">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-540">- BindingEvents</span></span><br><span data-ttu-id="dee93-541">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-541">
         - CompressedFile</span></span><br><span data-ttu-id="dee93-542">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dee93-542">
         - CustomXmlParts</span></span><br><span data-ttu-id="dee93-543">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-543">
         - DocumentEvents</span></span><br><span data-ttu-id="dee93-544">
         - File</span><span class="sxs-lookup"><span data-stu-id="dee93-544">
         - File</span></span><br><span data-ttu-id="dee93-545">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-545">
         - HtmlCoercion</span></span><br><span data-ttu-id="dee93-546">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-546">
         - ImageCoercion</span></span><br><span data-ttu-id="dee93-547">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-547">
         - MatrixBindings</span></span><br><span data-ttu-id="dee93-548">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-548">
         - MatrixCoercion</span></span><br><span data-ttu-id="dee93-549">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-549">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dee93-550">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-550">
         - PdfFile</span></span><br><span data-ttu-id="dee93-551">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-551">
         - Selection</span></span><br><span data-ttu-id="dee93-552">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-552">
         - Settings</span></span><br><span data-ttu-id="dee93-553">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-553">
         - TableBindings</span></span><br><span data-ttu-id="dee93-554">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-554">
         - TableCoercion</span></span><br><span data-ttu-id="dee93-555">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-555">
         - TextBindings</span></span><br><span data-ttu-id="dee93-556">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-556">
         - TextCoercion</span></span><br><span data-ttu-id="dee93-557">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dee93-557">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-558">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="dee93-558">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="dee93-559">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-559">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dee93-560">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-560">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="dee93-561">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-561">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="dee93-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="dee93-563">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="dee93-563">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="dee93-564">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-564">- BindingEvents</span></span><br><span data-ttu-id="dee93-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-565">
         - CompressedFile</span></span><br><span data-ttu-id="dee93-566">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dee93-566">
         - CustomXmlParts</span></span><br><span data-ttu-id="dee93-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-567">
         - DocumentEvents</span></span><br><span data-ttu-id="dee93-568">
         - File</span><span class="sxs-lookup"><span data-stu-id="dee93-568">
         - File</span></span><br><span data-ttu-id="dee93-569">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-569">
         - HtmlCoercion</span></span><br><span data-ttu-id="dee93-570">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-570">
         - ImageCoercion</span></span><br><span data-ttu-id="dee93-571">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-571">
         - MatrixBindings</span></span><br><span data-ttu-id="dee93-572">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-572">
         - MatrixCoercion</span></span><br><span data-ttu-id="dee93-573">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-573">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dee93-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-574">
         - PdfFile</span></span><br><span data-ttu-id="dee93-575">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-575">
         - Selection</span></span><br><span data-ttu-id="dee93-576">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-576">
         - Settings</span></span><br><span data-ttu-id="dee93-577">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-577">
         - TableBindings</span></span><br><span data-ttu-id="dee93-578">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-578">
         - TableCoercion</span></span><br><span data-ttu-id="dee93-579">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-579">
         - TextBindings</span></span><br><span data-ttu-id="dee93-580">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-580">
         - TextCoercion</span></span><br><span data-ttu-id="dee93-581">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dee93-581">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-582">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="dee93-582">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="dee93-583">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-583">- TaskPane</span></span><br><span data-ttu-id="dee93-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dee93-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="dee93-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="dee93-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="dee93-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="dee93-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="dee93-589">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-589">- BindingEvents</span></span><br><span data-ttu-id="dee93-590">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-590">
         - CompressedFile</span></span><br><span data-ttu-id="dee93-591">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dee93-591">
         - CustomXmlParts</span></span><br><span data-ttu-id="dee93-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-592">
         - DocumentEvents</span></span><br><span data-ttu-id="dee93-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="dee93-593">
         - File</span></span><br><span data-ttu-id="dee93-594">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-594">
         - HtmlCoercion</span></span><br><span data-ttu-id="dee93-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-595">
         - ImageCoercion</span></span><br><span data-ttu-id="dee93-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-596">
         - MatrixBindings</span></span><br><span data-ttu-id="dee93-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="dee93-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dee93-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-599">
         - PdfFile</span></span><br><span data-ttu-id="dee93-600">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-600">
         - Selection</span></span><br><span data-ttu-id="dee93-601">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-601">
         - Settings</span></span><br><span data-ttu-id="dee93-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-602">
         - TableBindings</span></span><br><span data-ttu-id="dee93-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-603">
         - TableCoercion</span></span><br><span data-ttu-id="dee93-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-604">
         - TextBindings</span></span><br><span data-ttu-id="dee93-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-605">
         - TextCoercion</span></span><br><span data-ttu-id="dee93-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dee93-606">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-607">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="dee93-607">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="dee93-608">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-608">- TaskPane</span></span><br><span data-ttu-id="dee93-609">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-609">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dee93-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="dee93-611">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dee93-611">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="dee93-612">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dee93-612">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="dee93-613">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="dee93-613">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="dee93-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-614">- BindingEvents</span></span><br><span data-ttu-id="dee93-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-615">
         - CompressedFile</span></span><br><span data-ttu-id="dee93-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dee93-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="dee93-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-617">
         - DocumentEvents</span></span><br><span data-ttu-id="dee93-618">
         - File</span><span class="sxs-lookup"><span data-stu-id="dee93-618">
         - File</span></span><br><span data-ttu-id="dee93-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="dee93-620">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-620">
         - ImageCoercion</span></span><br><span data-ttu-id="dee93-621">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-621">
         - MatrixBindings</span></span><br><span data-ttu-id="dee93-622">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-622">
         - MatrixCoercion</span></span><br><span data-ttu-id="dee93-623">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-623">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dee93-624">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-624">
         - PdfFile</span></span><br><span data-ttu-id="dee93-625">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-625">
         - Selection</span></span><br><span data-ttu-id="dee93-626">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-626">
         - Settings</span></span><br><span data-ttu-id="dee93-627">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-627">
         - TableBindings</span></span><br><span data-ttu-id="dee93-628">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-628">
         - TableCoercion</span></span><br><span data-ttu-id="dee93-629">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-629">
         - TextBindings</span></span><br><span data-ttu-id="dee93-630">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-630">
         - TextCoercion</span></span><br><span data-ttu-id="dee93-631">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dee93-631">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-632">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="dee93-632">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="dee93-633">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-633">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dee93-634">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-634">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="dee93-635">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="dee93-635">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="dee93-636">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-636">- BindingEvents</span></span><br><span data-ttu-id="dee93-637">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-637">
         - CompressedFile</span></span><br><span data-ttu-id="dee93-638">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dee93-638">
         - CustomXmlParts</span></span><br><span data-ttu-id="dee93-639">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-639">
         - DocumentEvents</span></span><br><span data-ttu-id="dee93-640">
         - File</span><span class="sxs-lookup"><span data-stu-id="dee93-640">
         - File</span></span><br><span data-ttu-id="dee93-641">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-641">
         - HtmlCoercion</span></span><br><span data-ttu-id="dee93-642">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-642">
         - ImageCoercion</span></span><br><span data-ttu-id="dee93-643">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-643">
         - MatrixBindings</span></span><br><span data-ttu-id="dee93-644">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-644">
         - MatrixCoercion</span></span><br><span data-ttu-id="dee93-645">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-645">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dee93-646">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-646">
         - PdfFile</span></span><br><span data-ttu-id="dee93-647">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-647">
         - Selection</span></span><br><span data-ttu-id="dee93-648">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-648">
         - Settings</span></span><br><span data-ttu-id="dee93-649">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-649">
         - TableBindings</span></span><br><span data-ttu-id="dee93-650">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-650">
         - TableCoercion</span></span><br><span data-ttu-id="dee93-651">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dee93-651">
         - TextBindings</span></span><br><span data-ttu-id="dee93-652">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-652">
         - TextCoercion</span></span><br><span data-ttu-id="dee93-653">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dee93-653">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="dee93-654">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="dee93-654">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="dee93-655">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="dee93-655">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="dee93-656">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dee93-656">Platform</span></span></th>
    <th><span data-ttu-id="dee93-657">拡張点</span><span class="sxs-lookup"><span data-stu-id="dee93-657">Extension points</span></span></th>
    <th><span data-ttu-id="dee93-658">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="dee93-658">API requirement sets</span></span></th>
    <th><span data-ttu-id="dee93-659"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="dee93-659"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="dee93-660">Office Online</span></span></td>
    <td> <span data-ttu-id="dee93-661">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dee93-661">- Content</span></span><br><span data-ttu-id="dee93-662">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-662">
         - TaskPane</span></span><br><span data-ttu-id="dee93-663">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-663">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dee93-664">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-664">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dee93-665">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dee93-665">- ActiveView</span></span><br><span data-ttu-id="dee93-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-666">
         - CompressedFile</span></span><br><span data-ttu-id="dee93-667">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-667">
         - DocumentEvents</span></span><br><span data-ttu-id="dee93-668">
         - File</span><span class="sxs-lookup"><span data-stu-id="dee93-668">
         - File</span></span><br><span data-ttu-id="dee93-669">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-669">
         - ImageCoercion</span></span><br><span data-ttu-id="dee93-670">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-670">
         - PdfFile</span></span><br><span data-ttu-id="dee93-671">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-671">
         - Selection</span></span><br><span data-ttu-id="dee93-672">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-672">
         - Settings</span></span><br><span data-ttu-id="dee93-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-673">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-674">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="dee93-674">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="dee93-675">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dee93-675">- Content</span></span><br><span data-ttu-id="dee93-676">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-676">
         - TaskPane</span></span><br><span data-ttu-id="dee93-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dee93-678">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-678">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dee93-679">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dee93-679">- ActiveView</span></span><br><span data-ttu-id="dee93-680">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-680">
         - CompressedFile</span></span><br><span data-ttu-id="dee93-681">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-681">
         - DocumentEvents</span></span><br><span data-ttu-id="dee93-682">
         - File</span><span class="sxs-lookup"><span data-stu-id="dee93-682">
         - File</span></span><br><span data-ttu-id="dee93-683">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-683">
         - ImageCoercion</span></span><br><span data-ttu-id="dee93-684">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-684">
         - PdfFile</span></span><br><span data-ttu-id="dee93-685">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-685">
         - Selection</span></span><br><span data-ttu-id="dee93-686">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-686">
         - Settings</span></span><br><span data-ttu-id="dee93-687">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-687">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-688">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="dee93-688">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="dee93-689">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dee93-689">- Content</span></span><br><span data-ttu-id="dee93-690">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-690">
         - TaskPane</span></span><br><span data-ttu-id="dee93-691">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-691">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dee93-692">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-692">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dee93-693">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dee93-693">- ActiveView</span></span><br><span data-ttu-id="dee93-694">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-694">
         - CompressedFile</span></span><br><span data-ttu-id="dee93-695">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-695">
         - DocumentEvents</span></span><br><span data-ttu-id="dee93-696">
         - File</span><span class="sxs-lookup"><span data-stu-id="dee93-696">
         - File</span></span><br><span data-ttu-id="dee93-697">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-697">
         - ImageCoercion</span></span><br><span data-ttu-id="dee93-698">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-698">
         - PdfFile</span></span><br><span data-ttu-id="dee93-699">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-699">
         - Selection</span></span><br><span data-ttu-id="dee93-700">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-700">
         - Settings</span></span><br><span data-ttu-id="dee93-701">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-701">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-702">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="dee93-702">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="dee93-703">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dee93-703">- Content</span></span><br><span data-ttu-id="dee93-704">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-704">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="dee93-705">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="dee93-705">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="dee93-706">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dee93-706">- ActiveView</span></span><br><span data-ttu-id="dee93-707">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-707">
         - CompressedFile</span></span><br><span data-ttu-id="dee93-708">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-708">
         - DocumentEvents</span></span><br><span data-ttu-id="dee93-709">
         - File</span><span class="sxs-lookup"><span data-stu-id="dee93-709">
         - File</span></span><br><span data-ttu-id="dee93-710">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-710">
         - ImageCoercion</span></span><br><span data-ttu-id="dee93-711">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-711">
         - PdfFile</span></span><br><span data-ttu-id="dee93-712">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-712">
         - Selection</span></span><br><span data-ttu-id="dee93-713">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-713">
         - Settings</span></span><br><span data-ttu-id="dee93-714">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-714">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-715">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="dee93-715">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="dee93-716">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dee93-716">- Content</span></span><br><span data-ttu-id="dee93-717">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-717">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="dee93-718">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="dee93-718">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="dee93-719">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dee93-719">- ActiveView</span></span><br><span data-ttu-id="dee93-720">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-720">
         - CompressedFile</span></span><br><span data-ttu-id="dee93-721">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-721">
         - DocumentEvents</span></span><br><span data-ttu-id="dee93-722">
         - File</span><span class="sxs-lookup"><span data-stu-id="dee93-722">
         - File</span></span><br><span data-ttu-id="dee93-723">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-723">
         - ImageCoercion</span></span><br><span data-ttu-id="dee93-724">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-724">
         - PdfFile</span></span><br><span data-ttu-id="dee93-725">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-725">
         - Selection</span></span><br><span data-ttu-id="dee93-726">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-726">
         - Settings</span></span><br><span data-ttu-id="dee93-727">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-727">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-728">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="dee93-728">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="dee93-729">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dee93-729">- Content</span></span><br><span data-ttu-id="dee93-730">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-730">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="dee93-731">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-731">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="dee93-732">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dee93-732">- ActiveView</span></span><br><span data-ttu-id="dee93-733">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-733">
         - CompressedFile</span></span><br><span data-ttu-id="dee93-734">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-734">
         - DocumentEvents</span></span><br><span data-ttu-id="dee93-735">
         - File</span><span class="sxs-lookup"><span data-stu-id="dee93-735">
         - File</span></span><br><span data-ttu-id="dee93-736">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-736">
         - PdfFile</span></span><br><span data-ttu-id="dee93-737">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-737">
         - Selection</span></span><br><span data-ttu-id="dee93-738">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-738">
         - Settings</span></span><br><span data-ttu-id="dee93-739">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-739">
         - TextCoercion</span></span><br><span data-ttu-id="dee93-740">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-740">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-741">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="dee93-741">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="dee93-742">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dee93-742">- Content</span></span><br><span data-ttu-id="dee93-743">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-743">
         - TaskPane</span></span><br><span data-ttu-id="dee93-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dee93-745">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-745">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dee93-746">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dee93-746">- ActiveView</span></span><br><span data-ttu-id="dee93-747">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-747">
         - CompressedFile</span></span><br><span data-ttu-id="dee93-748">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-748">
         - DocumentEvents</span></span><br><span data-ttu-id="dee93-749">
         - File</span><span class="sxs-lookup"><span data-stu-id="dee93-749">
         - File</span></span><br><span data-ttu-id="dee93-750">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-750">
         - ImageCoercion</span></span><br><span data-ttu-id="dee93-751">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-751">
         - PdfFile</span></span><br><span data-ttu-id="dee93-752">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-752">
         - Selection</span></span><br><span data-ttu-id="dee93-753">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-753">
         - Settings</span></span><br><span data-ttu-id="dee93-754">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-754">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-755">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="dee93-755">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="dee93-756">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dee93-756">- Content</span></span><br><span data-ttu-id="dee93-757">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-757">
         - TaskPane</span></span><br><span data-ttu-id="dee93-758">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-758">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dee93-759">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-759">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dee93-760">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dee93-760">- ActiveView</span></span><br><span data-ttu-id="dee93-761">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-761">
         - CompressedFile</span></span><br><span data-ttu-id="dee93-762">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-762">
         - DocumentEvents</span></span><br><span data-ttu-id="dee93-763">
         - File</span><span class="sxs-lookup"><span data-stu-id="dee93-763">
         - File</span></span><br><span data-ttu-id="dee93-764">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-764">
         - ImageCoercion</span></span><br><span data-ttu-id="dee93-765">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-765">
         - PdfFile</span></span><br><span data-ttu-id="dee93-766">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-766">
         - Selection</span></span><br><span data-ttu-id="dee93-767">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-767">
         - Settings</span></span><br><span data-ttu-id="dee93-768">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-768">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-769">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="dee93-769">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="dee93-770">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dee93-770">- Content</span></span><br><span data-ttu-id="dee93-771">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-771">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="dee93-772">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="dee93-772">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="dee93-773">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dee93-773">- ActiveView</span></span><br><span data-ttu-id="dee93-774">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dee93-774">
         - CompressedFile</span></span><br><span data-ttu-id="dee93-775">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-775">
         - DocumentEvents</span></span><br><span data-ttu-id="dee93-776">
         - File</span><span class="sxs-lookup"><span data-stu-id="dee93-776">
         - File</span></span><br><span data-ttu-id="dee93-777">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-777">
         - ImageCoercion</span></span><br><span data-ttu-id="dee93-778">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dee93-778">
         - PdfFile</span></span><br><span data-ttu-id="dee93-779">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-779">
         - Selection</span></span><br><span data-ttu-id="dee93-780">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-780">
         - Settings</span></span><br><span data-ttu-id="dee93-781">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-781">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="dee93-782">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="dee93-782">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="dee93-783">OneNote</span><span class="sxs-lookup"><span data-stu-id="dee93-783">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="dee93-784">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dee93-784">Platform</span></span></th>
    <th><span data-ttu-id="dee93-785">拡張点</span><span class="sxs-lookup"><span data-stu-id="dee93-785">Extension points</span></span></th>
    <th><span data-ttu-id="dee93-786">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="dee93-786">API requirement sets</span></span></th>
    <th><span data-ttu-id="dee93-787"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="dee93-787"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-788">Office Online</span><span class="sxs-lookup"><span data-stu-id="dee93-788">Office Online</span></span></td>
    <td> <span data-ttu-id="dee93-789">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dee93-789">- Content</span></span><br><span data-ttu-id="dee93-790">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-790">
         - TaskPane</span></span><br><span data-ttu-id="dee93-791">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dee93-791">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dee93-792">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-792">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="dee93-793">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-793">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dee93-794">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dee93-794">- DocumentEvents</span></span><br><span data-ttu-id="dee93-795">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-795">
         - HtmlCoercion</span></span><br><span data-ttu-id="dee93-796">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-796">
         - ImageCoercion</span></span><br><span data-ttu-id="dee93-797">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dee93-797">
         - Settings</span></span><br><span data-ttu-id="dee93-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-798">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="dee93-799">Project</span><span class="sxs-lookup"><span data-stu-id="dee93-799">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="dee93-800">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dee93-800">Platform</span></span></th>
    <th><span data-ttu-id="dee93-801">拡張点</span><span class="sxs-lookup"><span data-stu-id="dee93-801">Extension points</span></span></th>
    <th><span data-ttu-id="dee93-802">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="dee93-802">API requirement sets</span></span></th>
    <th><span data-ttu-id="dee93-803"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="dee93-803"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-804">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="dee93-804">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="dee93-805">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-805">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dee93-806">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-806">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dee93-807">- Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-807">- Selection</span></span><br><span data-ttu-id="dee93-808">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-808">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-809">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="dee93-809">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="dee93-810">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-810">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dee93-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dee93-812">- Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-812">- Selection</span></span><br><span data-ttu-id="dee93-813">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-813">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dee93-814">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="dee93-814">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="dee93-815">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dee93-815">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dee93-816">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dee93-816">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dee93-817">- Selection</span><span class="sxs-lookup"><span data-stu-id="dee93-817">- Selection</span></span><br><span data-ttu-id="dee93-818">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dee93-818">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="dee93-819">関連項目</span><span class="sxs-lookup"><span data-stu-id="dee93-819">See also</span></span>

- [<span data-ttu-id="dee93-820">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="dee93-820">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="dee93-821">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="dee93-821">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="dee93-822">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="dee93-822">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="dee93-823">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="dee93-823">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="dee93-824">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="dee93-824">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="dee93-825">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="dee93-825">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="dee93-826">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="dee93-826">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="dee93-827">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="dee93-827">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)