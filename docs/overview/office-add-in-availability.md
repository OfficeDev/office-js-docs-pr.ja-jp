---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、Word、Outlook、PowerPoint、OneNote、および Project のサポートされる要件セット。
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 28a6d0e4c86d05855ed9d24461dbeb77454d2b48
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872131"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="71af2-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="71af2-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="71af2-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="71af2-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="71af2-p102">MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。このバージョンには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="71af2-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>
>
> <span data-ttu-id="71af2-108">パッケージ版 Office 2019 のビルド番号は 16.0.10827.20150 です。</span><span class="sxs-lookup"><span data-stu-id="71af2-108">The build number for a one-time purchase of Office 2019 is 16.0.10827.20150.</span></span>

## <a name="excel"></a><span data-ttu-id="71af2-109">Excel</span><span class="sxs-lookup"><span data-stu-id="71af2-109">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="71af2-110">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="71af2-110">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="71af2-111">拡張点</span><span class="sxs-lookup"><span data-stu-id="71af2-111">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="71af2-112">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="71af2-112">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="71af2-113"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="71af2-113"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-114">Office Online</span><span class="sxs-lookup"><span data-stu-id="71af2-114">Office Online</span></span></td>
    <td> <span data-ttu-id="71af2-115">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-115">- TaskPane</span></span><br><span data-ttu-id="71af2-116">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="71af2-116">
        - Content</span></span><br><span data-ttu-id="71af2-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="71af2-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="71af2-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="71af2-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="71af2-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="71af2-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="71af2-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="71af2-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="71af2-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="71af2-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="71af2-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="71af2-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="71af2-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="71af2-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="71af2-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="71af2-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="71af2-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-127">
        - BindingEvents</span></span><br><span data-ttu-id="71af2-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-128">
        - CompressedFile</span></span><br><span data-ttu-id="71af2-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-129">
        - DocumentEvents</span></span><br><span data-ttu-id="71af2-130">
        - File</span><span class="sxs-lookup"><span data-stu-id="71af2-130">
        - File</span></span><br><span data-ttu-id="71af2-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-131">
        - MatrixBindings</span></span><br><span data-ttu-id="71af2-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-132">
        - MatrixCoercion</span></span><br><span data-ttu-id="71af2-133">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-133">
        - Selection</span></span><br><span data-ttu-id="71af2-134">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-134">
        - Settings</span></span><br><span data-ttu-id="71af2-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-135">
        - TableBindings</span></span><br><span data-ttu-id="71af2-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-136">
        - TableCoercion</span></span><br><span data-ttu-id="71af2-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-137">
        - TextBindings</span></span><br><span data-ttu-id="71af2-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-138">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-139">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="71af2-139">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="71af2-140">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-140">- TaskPane</span></span><br><span data-ttu-id="71af2-141">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="71af2-141">
        - Content</span></span><br><span data-ttu-id="71af2-142">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="71af2-142">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="71af2-143">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-143">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="71af2-144">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-144">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="71af2-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="71af2-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="71af2-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="71af2-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="71af2-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="71af2-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="71af2-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="71af2-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="71af2-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="71af2-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="71af2-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="71af2-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="71af2-152">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-152">
        - BindingEvents</span></span><br><span data-ttu-id="71af2-153">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-153">
        - CompressedFile</span></span><br><span data-ttu-id="71af2-154">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-154">
        - DocumentEvents</span></span><br><span data-ttu-id="71af2-155">
        - File</span><span class="sxs-lookup"><span data-stu-id="71af2-155">
        - File</span></span><br><span data-ttu-id="71af2-156">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-156">
        - MatrixBindings</span></span><br><span data-ttu-id="71af2-157">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-157">
        - MatrixCoercion</span></span><br><span data-ttu-id="71af2-158">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-158">
        - Selection</span></span><br><span data-ttu-id="71af2-159">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-159">
        - Settings</span></span><br><span data-ttu-id="71af2-160">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-160">
        - TableBindings</span></span><br><span data-ttu-id="71af2-161">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-161">
        - TableCoercion</span></span><br><span data-ttu-id="71af2-162">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-162">
        - TextBindings</span></span><br><span data-ttu-id="71af2-163">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-163">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-164">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="71af2-164">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="71af2-165">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-165">- TaskPane</span></span><br><span data-ttu-id="71af2-166">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="71af2-166">
        - Content</span></span><br><span data-ttu-id="71af2-167">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-167">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="71af2-168">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-168">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="71af2-169">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-169">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="71af2-170">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-170">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="71af2-171">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="71af2-171">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="71af2-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="71af2-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="71af2-173">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="71af2-173">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="71af2-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="71af2-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="71af2-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="71af2-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="71af2-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="71af2-177">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-177">- BindingEvents</span></span><br><span data-ttu-id="71af2-178">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-178">
        - CompressedFile</span></span><br><span data-ttu-id="71af2-179">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-179">
        - DocumentEvents</span></span><br><span data-ttu-id="71af2-180">
        - File</span><span class="sxs-lookup"><span data-stu-id="71af2-180">
        - File</span></span><br><span data-ttu-id="71af2-181">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-181">
        - ImageCoercion</span></span><br><span data-ttu-id="71af2-182">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-182">
        - MatrixBindings</span></span><br><span data-ttu-id="71af2-183">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-183">
        - MatrixCoercion</span></span><br><span data-ttu-id="71af2-184">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-184">
        - Selection</span></span><br><span data-ttu-id="71af2-185">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-185">
        - Settings</span></span><br><span data-ttu-id="71af2-186">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-186">
        - TableBindings</span></span><br><span data-ttu-id="71af2-187">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-187">
        - TableCoercion</span></span><br><span data-ttu-id="71af2-188">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-188">
        - TextBindings</span></span><br><span data-ttu-id="71af2-189">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-189">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-190">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="71af2-190">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="71af2-191">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-191">- TaskPane</span></span><br><span data-ttu-id="71af2-192">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="71af2-192">
        - Content</span></span></td>
    <td><span data-ttu-id="71af2-193">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-193">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="71af2-194">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="71af2-194">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="71af2-195">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-195">- BindingEvents</span></span><br><span data-ttu-id="71af2-196">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-196">
        - CompressedFile</span></span><br><span data-ttu-id="71af2-197">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-197">
        - DocumentEvents</span></span><br><span data-ttu-id="71af2-198">
        - File</span><span class="sxs-lookup"><span data-stu-id="71af2-198">
        - File</span></span><br><span data-ttu-id="71af2-199">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-199">
        - ImageCoercion</span></span><br><span data-ttu-id="71af2-200">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-200">
        - MatrixBindings</span></span><br><span data-ttu-id="71af2-201">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-201">
        - MatrixCoercion</span></span><br><span data-ttu-id="71af2-202">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-202">
        - Selection</span></span><br><span data-ttu-id="71af2-203">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-203">
        - Settings</span></span><br><span data-ttu-id="71af2-204">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-204">
        - TableBindings</span></span><br><span data-ttu-id="71af2-205">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-205">
        - TableCoercion</span></span><br><span data-ttu-id="71af2-206">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-206">
        - TextBindings</span></span><br><span data-ttu-id="71af2-207">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-207">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-208">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="71af2-208">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="71af2-209">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-209">
        - TaskPane</span></span><br><span data-ttu-id="71af2-210">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="71af2-210">
        - Content</span></span></td>
    <td>  <span data-ttu-id="71af2-211">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="71af2-211">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="71af2-212">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-212">
        - BindingEvents</span></span><br><span data-ttu-id="71af2-213">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-213">
        - CompressedFile</span></span><br><span data-ttu-id="71af2-214">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-214">
        - DocumentEvents</span></span><br><span data-ttu-id="71af2-215">
        - File</span><span class="sxs-lookup"><span data-stu-id="71af2-215">
        - File</span></span><br><span data-ttu-id="71af2-216">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-216">
        - ImageCoercion</span></span><br><span data-ttu-id="71af2-217">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-217">
        - MatrixBindings</span></span><br><span data-ttu-id="71af2-218">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-218">
        - MatrixCoercion</span></span><br><span data-ttu-id="71af2-219">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-219">
        - Selection</span></span><br><span data-ttu-id="71af2-220">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-220">
        - Settings</span></span><br><span data-ttu-id="71af2-221">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-221">
        - TableBindings</span></span><br><span data-ttu-id="71af2-222">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-222">
        - TableCoercion</span></span><br><span data-ttu-id="71af2-223">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-223">
        - TextBindings</span></span><br><span data-ttu-id="71af2-224">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-224">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-225">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="71af2-225">Office 365 for iPad</span></span></td>
    <td><span data-ttu-id="71af2-226">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-226">- TaskPane</span></span><br><span data-ttu-id="71af2-227">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="71af2-227">
        - Content</span></span></td>
    <td><span data-ttu-id="71af2-228">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-228">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="71af2-229">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-229">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="71af2-230">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-230">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="71af2-231">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="71af2-231">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="71af2-232">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="71af2-232">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="71af2-233">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="71af2-233">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="71af2-234">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="71af2-234">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="71af2-235">
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="71af2-235">
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="71af2-236">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-236">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="71af2-237">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-237">- BindingEvents</span></span><br><span data-ttu-id="71af2-238">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-238">
        - CompressedFile</span></span><br><span data-ttu-id="71af2-239">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-239">
        - DocumentEvents</span></span><br><span data-ttu-id="71af2-240">
        - File</span><span class="sxs-lookup"><span data-stu-id="71af2-240">
        - File</span></span><br><span data-ttu-id="71af2-241">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-241">
        - ImageCoercion</span></span><br><span data-ttu-id="71af2-242">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-242">
        - MatrixBindings</span></span><br><span data-ttu-id="71af2-243">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-243">
        - MatrixCoercion</span></span><br><span data-ttu-id="71af2-244">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-244">
        - Selection</span></span><br><span data-ttu-id="71af2-245">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-245">
        - Settings</span></span><br><span data-ttu-id="71af2-246">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-246">
        - TableBindings</span></span><br><span data-ttu-id="71af2-247">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-247">
        - TableCoercion</span></span><br><span data-ttu-id="71af2-248">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-248">
        - TextBindings</span></span><br><span data-ttu-id="71af2-249">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-249">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-250">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="71af2-250">Office 365 for Mac</span></span></td>
    <td><span data-ttu-id="71af2-251">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-251">- TaskPane</span></span><br><span data-ttu-id="71af2-252">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="71af2-252">
        - Content</span></span><br><span data-ttu-id="71af2-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="71af2-254">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-254">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="71af2-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="71af2-256">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-256">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="71af2-257">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="71af2-257">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="71af2-258">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="71af2-258">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="71af2-259">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="71af2-259">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="71af2-260">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="71af2-260">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="71af2-261">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="71af2-261">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="71af2-262">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-262">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="71af2-263">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-263">- BindingEvents</span></span><br><span data-ttu-id="71af2-264">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-264">
        - CompressedFile</span></span><br><span data-ttu-id="71af2-265">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-265">
        - DocumentEvents</span></span><br><span data-ttu-id="71af2-266">
        - File</span><span class="sxs-lookup"><span data-stu-id="71af2-266">
        - File</span></span><br><span data-ttu-id="71af2-267">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-267">
        - ImageCoercion</span></span><br><span data-ttu-id="71af2-268">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-268">
        - MatrixBindings</span></span><br><span data-ttu-id="71af2-269">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-269">
        - MatrixCoercion</span></span><br><span data-ttu-id="71af2-270">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-270">
        - PdfFile</span></span><br><span data-ttu-id="71af2-271">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-271">
        - Selection</span></span><br><span data-ttu-id="71af2-272">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-272">
        - Settings</span></span><br><span data-ttu-id="71af2-273">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-273">
        - TableBindings</span></span><br><span data-ttu-id="71af2-274">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-274">
        - TableCoercion</span></span><br><span data-ttu-id="71af2-275">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-275">
        - TextBindings</span></span><br><span data-ttu-id="71af2-276">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-276">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-277">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="71af2-277">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="71af2-278">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-278">- TaskPane</span></span><br><span data-ttu-id="71af2-279">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="71af2-279">
        - Content</span></span><br><span data-ttu-id="71af2-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="71af2-281">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-281">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="71af2-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="71af2-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="71af2-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="71af2-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="71af2-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="71af2-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="71af2-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="71af2-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="71af2-287">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="71af2-287">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="71af2-288">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="71af2-288">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="71af2-289">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-289">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="71af2-290">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-290">- BindingEvents</span></span><br><span data-ttu-id="71af2-291">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-291">
        - CompressedFile</span></span><br><span data-ttu-id="71af2-292">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-292">
        - DocumentEvents</span></span><br><span data-ttu-id="71af2-293">
        - File</span><span class="sxs-lookup"><span data-stu-id="71af2-293">
        - File</span></span><br><span data-ttu-id="71af2-294">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-294">
        - ImageCoercion</span></span><br><span data-ttu-id="71af2-295">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-295">
        - MatrixBindings</span></span><br><span data-ttu-id="71af2-296">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-296">
        - MatrixCoercion</span></span><br><span data-ttu-id="71af2-297">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-297">
        - PdfFile</span></span><br><span data-ttu-id="71af2-298">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-298">
        - Selection</span></span><br><span data-ttu-id="71af2-299">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-299">
        - Settings</span></span><br><span data-ttu-id="71af2-300">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-300">
        - TableBindings</span></span><br><span data-ttu-id="71af2-301">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-301">
        - TableCoercion</span></span><br><span data-ttu-id="71af2-302">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-302">
        - TextBindings</span></span><br><span data-ttu-id="71af2-303">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-303">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-304">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="71af2-304">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="71af2-305">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-305">- TaskPane</span></span><br><span data-ttu-id="71af2-306">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="71af2-306">
        - Content</span></span></td>
    <td><span data-ttu-id="71af2-307">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-307">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="71af2-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="71af2-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="71af2-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-309">- BindingEvents</span></span><br><span data-ttu-id="71af2-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-310">
        - CompressedFile</span></span><br><span data-ttu-id="71af2-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-311">
        - DocumentEvents</span></span><br><span data-ttu-id="71af2-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="71af2-312">
        - File</span></span><br><span data-ttu-id="71af2-313">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-313">
        - ImageCoercion</span></span><br><span data-ttu-id="71af2-314">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-314">
        - MatrixBindings</span></span><br><span data-ttu-id="71af2-315">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-315">
        - MatrixCoercion</span></span><br><span data-ttu-id="71af2-316">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-316">
        - PdfFile</span></span><br><span data-ttu-id="71af2-317">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-317">
        - Selection</span></span><br><span data-ttu-id="71af2-318">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-318">
        - Settings</span></span><br><span data-ttu-id="71af2-319">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-319">
        - TableBindings</span></span><br><span data-ttu-id="71af2-320">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-320">
        - TableCoercion</span></span><br><span data-ttu-id="71af2-321">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-321">
        - TextBindings</span></span><br><span data-ttu-id="71af2-322">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-322">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="71af2-323">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="71af2-323">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="71af2-324">Outlook</span><span class="sxs-lookup"><span data-stu-id="71af2-324">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="71af2-325">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="71af2-325">Platform</span></span></th>
    <th><span data-ttu-id="71af2-326">拡張点</span><span class="sxs-lookup"><span data-stu-id="71af2-326">Extension points</span></span></th>
    <th><span data-ttu-id="71af2-327">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="71af2-327">API requirement sets</span></span></th>
    <th><span data-ttu-id="71af2-328"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="71af2-328"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-329">Office Online</span><span class="sxs-lookup"><span data-stu-id="71af2-329">Office Online</span></span></td>
    <td> <span data-ttu-id="71af2-330">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="71af2-330">- Mail Read</span></span><br><span data-ttu-id="71af2-331">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="71af2-331">
      - Mail Compose</span></span><br><span data-ttu-id="71af2-332">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-332">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="71af2-333">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-333">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="71af2-334">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-334">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="71af2-335">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-335">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="71af2-336">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="71af2-336">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="71af2-337">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="71af2-337">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="71af2-338">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="71af2-338">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="71af2-339">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="71af2-339">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="71af2-340">利用不可</span><span class="sxs-lookup"><span data-stu-id="71af2-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-341">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="71af2-341">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="71af2-342">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="71af2-342">- Mail Read</span></span><br><span data-ttu-id="71af2-343">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="71af2-343">
      - Mail Compose</span></span><br><span data-ttu-id="71af2-344">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-344">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="71af2-345">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="71af2-345">
      - Modules</span></span></td>
    <td> <span data-ttu-id="71af2-346">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-346">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="71af2-347">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-347">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="71af2-348">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-348">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="71af2-349">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="71af2-349">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="71af2-350">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="71af2-350">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="71af2-351">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="71af2-351">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="71af2-352">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="71af2-352">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="71af2-353">利用不可</span><span class="sxs-lookup"><span data-stu-id="71af2-353">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-354">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="71af2-354">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="71af2-355">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="71af2-355">- Mail Read</span></span><br><span data-ttu-id="71af2-356">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="71af2-356">
      - Mail Compose</span></span><br><span data-ttu-id="71af2-357">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-357">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="71af2-358">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="71af2-358">
      - Modules</span></span></td>
    <td> <span data-ttu-id="71af2-359">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-359">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="71af2-360">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-360">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="71af2-361">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-361">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="71af2-362">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="71af2-362">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="71af2-363">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="71af2-363">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="71af2-364">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="71af2-364">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="71af2-365">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="71af2-365">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="71af2-366">利用不可</span><span class="sxs-lookup"><span data-stu-id="71af2-366">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-367">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="71af2-367">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="71af2-368">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="71af2-368">- Mail Read</span></span><br><span data-ttu-id="71af2-369">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="71af2-369">
      - Mail Compose</span></span><br><span data-ttu-id="71af2-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="71af2-371">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="71af2-371">
      - Modules</span></span></td>
    <td> <span data-ttu-id="71af2-372">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-372">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="71af2-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="71af2-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="71af2-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="71af2-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="71af2-376">利用不可</span><span class="sxs-lookup"><span data-stu-id="71af2-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-377">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="71af2-377">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="71af2-378">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="71af2-378">- Mail Read</span></span><br><span data-ttu-id="71af2-379">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="71af2-379">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="71af2-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="71af2-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="71af2-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="71af2-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="71af2-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="71af2-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="71af2-384">利用不可</span><span class="sxs-lookup"><span data-stu-id="71af2-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-385">Office 365 for iOS</span><span class="sxs-lookup"><span data-stu-id="71af2-385">Office 365 for iOS</span></span></td>
    <td> <span data-ttu-id="71af2-386">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="71af2-386">- Mail Read</span></span><br><span data-ttu-id="71af2-387">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-387">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="71af2-388">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-388">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="71af2-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="71af2-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="71af2-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="71af2-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="71af2-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="71af2-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="71af2-393">利用不可</span><span class="sxs-lookup"><span data-stu-id="71af2-393">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-394">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="71af2-394">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="71af2-395">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="71af2-395">- Mail Read</span></span><br><span data-ttu-id="71af2-396">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="71af2-396">
      - Mail Compose</span></span><br><span data-ttu-id="71af2-397">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-397">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="71af2-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="71af2-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="71af2-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="71af2-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="71af2-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="71af2-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="71af2-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="71af2-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="71af2-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="71af2-404">利用不可</span><span class="sxs-lookup"><span data-stu-id="71af2-404">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-405">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="71af2-405">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="71af2-406">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="71af2-406">- Mail Read</span></span><br><span data-ttu-id="71af2-407">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="71af2-407">
      - Mail Compose</span></span><br><span data-ttu-id="71af2-408">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-408">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="71af2-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="71af2-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="71af2-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="71af2-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="71af2-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="71af2-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="71af2-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="71af2-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="71af2-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="71af2-415">利用不可</span><span class="sxs-lookup"><span data-stu-id="71af2-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-416">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="71af2-416">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="71af2-417">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="71af2-417">- Mail Read</span></span><br><span data-ttu-id="71af2-418">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="71af2-418">
      - Mail Compose</span></span><br><span data-ttu-id="71af2-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="71af2-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="71af2-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="71af2-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="71af2-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="71af2-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="71af2-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="71af2-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="71af2-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="71af2-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="71af2-426">利用不可</span><span class="sxs-lookup"><span data-stu-id="71af2-426">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-427">Office 365 for Android</span><span class="sxs-lookup"><span data-stu-id="71af2-427">Office 365 for Android</span></span></td>
    <td> <span data-ttu-id="71af2-428">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="71af2-428">- Mail Read</span></span><br><span data-ttu-id="71af2-429">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-429">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="71af2-430">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-430">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="71af2-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="71af2-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="71af2-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="71af2-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="71af2-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="71af2-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="71af2-435">利用不可</span><span class="sxs-lookup"><span data-stu-id="71af2-435">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="71af2-436">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="71af2-436">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="71af2-437">Word</span><span class="sxs-lookup"><span data-stu-id="71af2-437">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="71af2-438">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="71af2-438">Platform</span></span></th>
    <th><span data-ttu-id="71af2-439">拡張点</span><span class="sxs-lookup"><span data-stu-id="71af2-439">Extension points</span></span></th>
    <th><span data-ttu-id="71af2-440">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="71af2-440">API requirement sets</span></span></th>
    <th><span data-ttu-id="71af2-441"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="71af2-441"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-442">Office Online</span><span class="sxs-lookup"><span data-stu-id="71af2-442">Office Online</span></span></td>
    <td> <span data-ttu-id="71af2-443">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-443">- TaskPane</span></span><br><span data-ttu-id="71af2-444">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-444">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="71af2-445">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-445">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="71af2-446">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-446">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="71af2-447">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-447">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="71af2-448">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-448">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="71af2-449">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-449">- BindingEvents</span></span><br><span data-ttu-id="71af2-450">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="71af2-450">
         - CustomXmlParts</span></span><br><span data-ttu-id="71af2-451">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-451">
         - DocumentEvents</span></span><br><span data-ttu-id="71af2-452">
         - File</span><span class="sxs-lookup"><span data-stu-id="71af2-452">
         - File</span></span><br><span data-ttu-id="71af2-453">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-453">
         - HtmlCoercion</span></span><br><span data-ttu-id="71af2-454">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-454">
         - ImageCoercion</span></span><br><span data-ttu-id="71af2-455">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-455">
         - MatrixBindings</span></span><br><span data-ttu-id="71af2-456">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-456">
         - MatrixCoercion</span></span><br><span data-ttu-id="71af2-457">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-457">
         - OoxmlCoercion</span></span><br><span data-ttu-id="71af2-458">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-458">
         - PdfFile</span></span><br><span data-ttu-id="71af2-459">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-459">
         - Selection</span></span><br><span data-ttu-id="71af2-460">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-460">
         - Settings</span></span><br><span data-ttu-id="71af2-461">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-461">
         - TableBindings</span></span><br><span data-ttu-id="71af2-462">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-462">
         - TableCoercion</span></span><br><span data-ttu-id="71af2-463">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-463">
         - TextBindings</span></span><br><span data-ttu-id="71af2-464">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-464">
         - TextCoercion</span></span><br><span data-ttu-id="71af2-465">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="71af2-465">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-466">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="71af2-466">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="71af2-467">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-467">- TaskPane</span></span><br><span data-ttu-id="71af2-468">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-468">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="71af2-469">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-469">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="71af2-470">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-470">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="71af2-471">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-471">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="71af2-472">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-472">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="71af2-473">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-473">- BindingEvents</span></span><br><span data-ttu-id="71af2-474">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-474">
         - CompressedFile</span></span><br><span data-ttu-id="71af2-475">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="71af2-475">
         - CustomXmlParts</span></span><br><span data-ttu-id="71af2-476">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-476">
         - DocumentEvents</span></span><br><span data-ttu-id="71af2-477">
         - File</span><span class="sxs-lookup"><span data-stu-id="71af2-477">
         - File</span></span><br><span data-ttu-id="71af2-478">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-478">
         - HtmlCoercion</span></span><br><span data-ttu-id="71af2-479">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-479">
         - ImageCoercion</span></span><br><span data-ttu-id="71af2-480">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-480">
         - MatrixBindings</span></span><br><span data-ttu-id="71af2-481">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-481">
         - MatrixCoercion</span></span><br><span data-ttu-id="71af2-482">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-482">
         - OoxmlCoercion</span></span><br><span data-ttu-id="71af2-483">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-483">
         - PdfFile</span></span><br><span data-ttu-id="71af2-484">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-484">
         - Selection</span></span><br><span data-ttu-id="71af2-485">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-485">
         - Settings</span></span><br><span data-ttu-id="71af2-486">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-486">
         - TableBindings</span></span><br><span data-ttu-id="71af2-487">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-487">
         - TableCoercion</span></span><br><span data-ttu-id="71af2-488">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-488">
         - TextBindings</span></span><br><span data-ttu-id="71af2-489">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-489">
         - TextCoercion</span></span><br><span data-ttu-id="71af2-490">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="71af2-490">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-491">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="71af2-491">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="71af2-492">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-492">- TaskPane</span></span><br><span data-ttu-id="71af2-493">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-493">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="71af2-494">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-494">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="71af2-495">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-495">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="71af2-496">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-496">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="71af2-497">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-497">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="71af2-498">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-498">- BindingEvents</span></span><br><span data-ttu-id="71af2-499">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-499">
         - CompressedFile</span></span><br><span data-ttu-id="71af2-500">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="71af2-500">
         - CustomXmlParts</span></span><br><span data-ttu-id="71af2-501">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-501">
         - DocumentEvents</span></span><br><span data-ttu-id="71af2-502">
         - File</span><span class="sxs-lookup"><span data-stu-id="71af2-502">
         - File</span></span><br><span data-ttu-id="71af2-503">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-503">
         - HtmlCoercion</span></span><br><span data-ttu-id="71af2-504">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-504">
         - ImageCoercion</span></span><br><span data-ttu-id="71af2-505">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-505">
         - MatrixBindings</span></span><br><span data-ttu-id="71af2-506">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-506">
         - MatrixCoercion</span></span><br><span data-ttu-id="71af2-507">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-507">
         - OoxmlCoercion</span></span><br><span data-ttu-id="71af2-508">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-508">
         - PdfFile</span></span><br><span data-ttu-id="71af2-509">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-509">
         - Selection</span></span><br><span data-ttu-id="71af2-510">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-510">
         - Settings</span></span><br><span data-ttu-id="71af2-511">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-511">
         - TableBindings</span></span><br><span data-ttu-id="71af2-512">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-512">
         - TableCoercion</span></span><br><span data-ttu-id="71af2-513">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-513">
         - TextBindings</span></span><br><span data-ttu-id="71af2-514">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-514">
         - TextCoercion</span></span><br><span data-ttu-id="71af2-515">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="71af2-515">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-516">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="71af2-516">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="71af2-517">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-517">- TaskPane</span></span></td>
    <td> <span data-ttu-id="71af2-518">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-518">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="71af2-519">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="71af2-519">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="71af2-520">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-520">- BindingEvents</span></span><br><span data-ttu-id="71af2-521">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-521">
         - CompressedFile</span></span><br><span data-ttu-id="71af2-522">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="71af2-522">
         - CustomXmlParts</span></span><br><span data-ttu-id="71af2-523">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-523">
         - DocumentEvents</span></span><br><span data-ttu-id="71af2-524">
         - File</span><span class="sxs-lookup"><span data-stu-id="71af2-524">
         - File</span></span><br><span data-ttu-id="71af2-525">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-525">
         - HtmlCoercion</span></span><br><span data-ttu-id="71af2-526">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-526">
         - ImageCoercion</span></span><br><span data-ttu-id="71af2-527">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-527">
         - MatrixBindings</span></span><br><span data-ttu-id="71af2-528">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-528">
         - MatrixCoercion</span></span><br><span data-ttu-id="71af2-529">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-529">
         - OoxmlCoercion</span></span><br><span data-ttu-id="71af2-530">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-530">
         - PdfFile</span></span><br><span data-ttu-id="71af2-531">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-531">
         - Selection</span></span><br><span data-ttu-id="71af2-532">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-532">
         - Settings</span></span><br><span data-ttu-id="71af2-533">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-533">
         - TableBindings</span></span><br><span data-ttu-id="71af2-534">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-534">
         - TableCoercion</span></span><br><span data-ttu-id="71af2-535">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-535">
         - TextBindings</span></span><br><span data-ttu-id="71af2-536">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-536">
         - TextCoercion</span></span><br><span data-ttu-id="71af2-537">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="71af2-537">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-538">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="71af2-538">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="71af2-539">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-539">- TaskPane</span></span></td>
    <td> <span data-ttu-id="71af2-540">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="71af2-540">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="71af2-541">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-541">- BindingEvents</span></span><br><span data-ttu-id="71af2-542">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-542">
         - CompressedFile</span></span><br><span data-ttu-id="71af2-543">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="71af2-543">
         - CustomXmlParts</span></span><br><span data-ttu-id="71af2-544">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-544">
         - DocumentEvents</span></span><br><span data-ttu-id="71af2-545">
         - File</span><span class="sxs-lookup"><span data-stu-id="71af2-545">
         - File</span></span><br><span data-ttu-id="71af2-546">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-546">
         - HtmlCoercion</span></span><br><span data-ttu-id="71af2-547">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-547">
         - ImageCoercion</span></span><br><span data-ttu-id="71af2-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-548">
         - MatrixBindings</span></span><br><span data-ttu-id="71af2-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="71af2-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="71af2-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-551">
         - PdfFile</span></span><br><span data-ttu-id="71af2-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-552">
         - Selection</span></span><br><span data-ttu-id="71af2-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-553">
         - Settings</span></span><br><span data-ttu-id="71af2-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-554">
         - TableBindings</span></span><br><span data-ttu-id="71af2-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-555">
         - TableCoercion</span></span><br><span data-ttu-id="71af2-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-556">
         - TextBindings</span></span><br><span data-ttu-id="71af2-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-557">
         - TextCoercion</span></span><br><span data-ttu-id="71af2-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="71af2-558">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-559">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="71af2-559">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="71af2-560">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-560">- TaskPane</span></span></td>
    <td> <span data-ttu-id="71af2-561">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-561">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="71af2-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="71af2-563">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-563">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="71af2-564">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="71af2-564">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="71af2-565">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-565">- BindingEvents</span></span><br><span data-ttu-id="71af2-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-566">
         - CompressedFile</span></span><br><span data-ttu-id="71af2-567">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="71af2-567">
         - CustomXmlParts</span></span><br><span data-ttu-id="71af2-568">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-568">
         - DocumentEvents</span></span><br><span data-ttu-id="71af2-569">
         - File</span><span class="sxs-lookup"><span data-stu-id="71af2-569">
         - File</span></span><br><span data-ttu-id="71af2-570">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-570">
         - HtmlCoercion</span></span><br><span data-ttu-id="71af2-571">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-571">
         - ImageCoercion</span></span><br><span data-ttu-id="71af2-572">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-572">
         - MatrixBindings</span></span><br><span data-ttu-id="71af2-573">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-573">
         - MatrixCoercion</span></span><br><span data-ttu-id="71af2-574">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-574">
         - OoxmlCoercion</span></span><br><span data-ttu-id="71af2-575">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-575">
         - PdfFile</span></span><br><span data-ttu-id="71af2-576">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-576">
         - Selection</span></span><br><span data-ttu-id="71af2-577">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-577">
         - Settings</span></span><br><span data-ttu-id="71af2-578">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-578">
         - TableBindings</span></span><br><span data-ttu-id="71af2-579">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-579">
         - TableCoercion</span></span><br><span data-ttu-id="71af2-580">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-580">
         - TextBindings</span></span><br><span data-ttu-id="71af2-581">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-581">
         - TextCoercion</span></span><br><span data-ttu-id="71af2-582">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="71af2-582">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-583">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="71af2-583">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="71af2-584">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-584">- TaskPane</span></span><br><span data-ttu-id="71af2-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="71af2-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="71af2-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="71af2-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="71af2-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="71af2-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="71af2-590">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-590">- BindingEvents</span></span><br><span data-ttu-id="71af2-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-591">
         - CompressedFile</span></span><br><span data-ttu-id="71af2-592">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="71af2-592">
         - CustomXmlParts</span></span><br><span data-ttu-id="71af2-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-593">
         - DocumentEvents</span></span><br><span data-ttu-id="71af2-594">
         - File</span><span class="sxs-lookup"><span data-stu-id="71af2-594">
         - File</span></span><br><span data-ttu-id="71af2-595">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-595">
         - HtmlCoercion</span></span><br><span data-ttu-id="71af2-596">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-596">
         - ImageCoercion</span></span><br><span data-ttu-id="71af2-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-597">
         - MatrixBindings</span></span><br><span data-ttu-id="71af2-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="71af2-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="71af2-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-600">
         - PdfFile</span></span><br><span data-ttu-id="71af2-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-601">
         - Selection</span></span><br><span data-ttu-id="71af2-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-602">
         - Settings</span></span><br><span data-ttu-id="71af2-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-603">
         - TableBindings</span></span><br><span data-ttu-id="71af2-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-604">
         - TableCoercion</span></span><br><span data-ttu-id="71af2-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-605">
         - TextBindings</span></span><br><span data-ttu-id="71af2-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-606">
         - TextCoercion</span></span><br><span data-ttu-id="71af2-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="71af2-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-608">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="71af2-608">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="71af2-609">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-609">- TaskPane</span></span><br><span data-ttu-id="71af2-610">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-610">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="71af2-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="71af2-612">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="71af2-612">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="71af2-613">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="71af2-613">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="71af2-614">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="71af2-614">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="71af2-615">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-615">- BindingEvents</span></span><br><span data-ttu-id="71af2-616">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-616">
         - CompressedFile</span></span><br><span data-ttu-id="71af2-617">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="71af2-617">
         - CustomXmlParts</span></span><br><span data-ttu-id="71af2-618">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-618">
         - DocumentEvents</span></span><br><span data-ttu-id="71af2-619">
         - File</span><span class="sxs-lookup"><span data-stu-id="71af2-619">
         - File</span></span><br><span data-ttu-id="71af2-620">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-620">
         - HtmlCoercion</span></span><br><span data-ttu-id="71af2-621">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-621">
         - ImageCoercion</span></span><br><span data-ttu-id="71af2-622">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-622">
         - MatrixBindings</span></span><br><span data-ttu-id="71af2-623">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-623">
         - MatrixCoercion</span></span><br><span data-ttu-id="71af2-624">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-624">
         - OoxmlCoercion</span></span><br><span data-ttu-id="71af2-625">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-625">
         - PdfFile</span></span><br><span data-ttu-id="71af2-626">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-626">
         - Selection</span></span><br><span data-ttu-id="71af2-627">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-627">
         - Settings</span></span><br><span data-ttu-id="71af2-628">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-628">
         - TableBindings</span></span><br><span data-ttu-id="71af2-629">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-629">
         - TableCoercion</span></span><br><span data-ttu-id="71af2-630">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-630">
         - TextBindings</span></span><br><span data-ttu-id="71af2-631">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-631">
         - TextCoercion</span></span><br><span data-ttu-id="71af2-632">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="71af2-632">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-633">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="71af2-633">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="71af2-634">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-634">- TaskPane</span></span></td>
    <td> <span data-ttu-id="71af2-635">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-635">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="71af2-636">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="71af2-636">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="71af2-637">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-637">- BindingEvents</span></span><br><span data-ttu-id="71af2-638">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-638">
         - CompressedFile</span></span><br><span data-ttu-id="71af2-639">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="71af2-639">
         - CustomXmlParts</span></span><br><span data-ttu-id="71af2-640">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-640">
         - DocumentEvents</span></span><br><span data-ttu-id="71af2-641">
         - File</span><span class="sxs-lookup"><span data-stu-id="71af2-641">
         - File</span></span><br><span data-ttu-id="71af2-642">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-642">
         - HtmlCoercion</span></span><br><span data-ttu-id="71af2-643">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-643">
         - ImageCoercion</span></span><br><span data-ttu-id="71af2-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-644">
         - MatrixBindings</span></span><br><span data-ttu-id="71af2-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="71af2-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="71af2-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-647">
         - PdfFile</span></span><br><span data-ttu-id="71af2-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-648">
         - Selection</span></span><br><span data-ttu-id="71af2-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-649">
         - Settings</span></span><br><span data-ttu-id="71af2-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-650">
         - TableBindings</span></span><br><span data-ttu-id="71af2-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-651">
         - TableCoercion</span></span><br><span data-ttu-id="71af2-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="71af2-652">
         - TextBindings</span></span><br><span data-ttu-id="71af2-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-653">
         - TextCoercion</span></span><br><span data-ttu-id="71af2-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="71af2-654">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="71af2-655">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="71af2-655">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="71af2-656">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="71af2-656">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="71af2-657">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="71af2-657">Platform</span></span></th>
    <th><span data-ttu-id="71af2-658">拡張点</span><span class="sxs-lookup"><span data-stu-id="71af2-658">Extension points</span></span></th>
    <th><span data-ttu-id="71af2-659">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="71af2-659">API requirement sets</span></span></th>
    <th><span data-ttu-id="71af2-660"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="71af2-660"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-661">Office Online</span><span class="sxs-lookup"><span data-stu-id="71af2-661">Office Online</span></span></td>
    <td> <span data-ttu-id="71af2-662">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="71af2-662">- Content</span></span><br><span data-ttu-id="71af2-663">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-663">
         - TaskPane</span></span><br><span data-ttu-id="71af2-664">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-664">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="71af2-665">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-665">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="71af2-666">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="71af2-666">- ActiveView</span></span><br><span data-ttu-id="71af2-667">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-667">
         - CompressedFile</span></span><br><span data-ttu-id="71af2-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-668">
         - DocumentEvents</span></span><br><span data-ttu-id="71af2-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="71af2-669">
         - File</span></span><br><span data-ttu-id="71af2-670">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-670">
         - ImageCoercion</span></span><br><span data-ttu-id="71af2-671">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-671">
         - PdfFile</span></span><br><span data-ttu-id="71af2-672">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-672">
         - Selection</span></span><br><span data-ttu-id="71af2-673">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-673">
         - Settings</span></span><br><span data-ttu-id="71af2-674">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-674">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-675">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="71af2-675">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="71af2-676">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="71af2-676">- Content</span></span><br><span data-ttu-id="71af2-677">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-677">
         - TaskPane</span></span><br><span data-ttu-id="71af2-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="71af2-679">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-679">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="71af2-680">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="71af2-680">- ActiveView</span></span><br><span data-ttu-id="71af2-681">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-681">
         - CompressedFile</span></span><br><span data-ttu-id="71af2-682">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-682">
         - DocumentEvents</span></span><br><span data-ttu-id="71af2-683">
         - File</span><span class="sxs-lookup"><span data-stu-id="71af2-683">
         - File</span></span><br><span data-ttu-id="71af2-684">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-684">
         - ImageCoercion</span></span><br><span data-ttu-id="71af2-685">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-685">
         - PdfFile</span></span><br><span data-ttu-id="71af2-686">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-686">
         - Selection</span></span><br><span data-ttu-id="71af2-687">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-687">
         - Settings</span></span><br><span data-ttu-id="71af2-688">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-688">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-689">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="71af2-689">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="71af2-690">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="71af2-690">- Content</span></span><br><span data-ttu-id="71af2-691">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-691">
         - TaskPane</span></span><br><span data-ttu-id="71af2-692">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-692">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="71af2-693">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-693">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="71af2-694">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="71af2-694">- ActiveView</span></span><br><span data-ttu-id="71af2-695">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-695">
         - CompressedFile</span></span><br><span data-ttu-id="71af2-696">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-696">
         - DocumentEvents</span></span><br><span data-ttu-id="71af2-697">
         - File</span><span class="sxs-lookup"><span data-stu-id="71af2-697">
         - File</span></span><br><span data-ttu-id="71af2-698">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-698">
         - ImageCoercion</span></span><br><span data-ttu-id="71af2-699">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-699">
         - PdfFile</span></span><br><span data-ttu-id="71af2-700">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-700">
         - Selection</span></span><br><span data-ttu-id="71af2-701">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-701">
         - Settings</span></span><br><span data-ttu-id="71af2-702">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-702">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-703">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="71af2-703">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="71af2-704">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="71af2-704">- Content</span></span><br><span data-ttu-id="71af2-705">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-705">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="71af2-706">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="71af2-706">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="71af2-707">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="71af2-707">- ActiveView</span></span><br><span data-ttu-id="71af2-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-708">
         - CompressedFile</span></span><br><span data-ttu-id="71af2-709">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-709">
         - DocumentEvents</span></span><br><span data-ttu-id="71af2-710">
         - File</span><span class="sxs-lookup"><span data-stu-id="71af2-710">
         - File</span></span><br><span data-ttu-id="71af2-711">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-711">
         - ImageCoercion</span></span><br><span data-ttu-id="71af2-712">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-712">
         - PdfFile</span></span><br><span data-ttu-id="71af2-713">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-713">
         - Selection</span></span><br><span data-ttu-id="71af2-714">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-714">
         - Settings</span></span><br><span data-ttu-id="71af2-715">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-715">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-716">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="71af2-716">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="71af2-717">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="71af2-717">- Content</span></span><br><span data-ttu-id="71af2-718">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-718">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="71af2-719">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="71af2-719">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="71af2-720">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="71af2-720">- ActiveView</span></span><br><span data-ttu-id="71af2-721">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-721">
         - CompressedFile</span></span><br><span data-ttu-id="71af2-722">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-722">
         - DocumentEvents</span></span><br><span data-ttu-id="71af2-723">
         - File</span><span class="sxs-lookup"><span data-stu-id="71af2-723">
         - File</span></span><br><span data-ttu-id="71af2-724">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-724">
         - ImageCoercion</span></span><br><span data-ttu-id="71af2-725">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-725">
         - PdfFile</span></span><br><span data-ttu-id="71af2-726">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-726">
         - Selection</span></span><br><span data-ttu-id="71af2-727">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-727">
         - Settings</span></span><br><span data-ttu-id="71af2-728">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-728">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-729">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="71af2-729">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="71af2-730">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="71af2-730">- Content</span></span><br><span data-ttu-id="71af2-731">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-731">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="71af2-732">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-732">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="71af2-733">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="71af2-733">- ActiveView</span></span><br><span data-ttu-id="71af2-734">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-734">
         - CompressedFile</span></span><br><span data-ttu-id="71af2-735">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-735">
         - DocumentEvents</span></span><br><span data-ttu-id="71af2-736">
         - File</span><span class="sxs-lookup"><span data-stu-id="71af2-736">
         - File</span></span><br><span data-ttu-id="71af2-737">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-737">
         - PdfFile</span></span><br><span data-ttu-id="71af2-738">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-738">
         - Selection</span></span><br><span data-ttu-id="71af2-739">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-739">
         - Settings</span></span><br><span data-ttu-id="71af2-740">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-740">
         - TextCoercion</span></span><br><span data-ttu-id="71af2-741">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-741">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-742">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="71af2-742">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="71af2-743">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="71af2-743">- Content</span></span><br><span data-ttu-id="71af2-744">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-744">
         - TaskPane</span></span><br><span data-ttu-id="71af2-745">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-745">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="71af2-746">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-746">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="71af2-747">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="71af2-747">- ActiveView</span></span><br><span data-ttu-id="71af2-748">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-748">
         - CompressedFile</span></span><br><span data-ttu-id="71af2-749">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-749">
         - DocumentEvents</span></span><br><span data-ttu-id="71af2-750">
         - File</span><span class="sxs-lookup"><span data-stu-id="71af2-750">
         - File</span></span><br><span data-ttu-id="71af2-751">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-751">
         - ImageCoercion</span></span><br><span data-ttu-id="71af2-752">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-752">
         - PdfFile</span></span><br><span data-ttu-id="71af2-753">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-753">
         - Selection</span></span><br><span data-ttu-id="71af2-754">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-754">
         - Settings</span></span><br><span data-ttu-id="71af2-755">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-755">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-756">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="71af2-756">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="71af2-757">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="71af2-757">- Content</span></span><br><span data-ttu-id="71af2-758">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-758">
         - TaskPane</span></span><br><span data-ttu-id="71af2-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="71af2-760">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-760">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="71af2-761">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="71af2-761">- ActiveView</span></span><br><span data-ttu-id="71af2-762">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-762">
         - CompressedFile</span></span><br><span data-ttu-id="71af2-763">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-763">
         - DocumentEvents</span></span><br><span data-ttu-id="71af2-764">
         - File</span><span class="sxs-lookup"><span data-stu-id="71af2-764">
         - File</span></span><br><span data-ttu-id="71af2-765">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-765">
         - ImageCoercion</span></span><br><span data-ttu-id="71af2-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-766">
         - PdfFile</span></span><br><span data-ttu-id="71af2-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-767">
         - Selection</span></span><br><span data-ttu-id="71af2-768">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-768">
         - Settings</span></span><br><span data-ttu-id="71af2-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-770">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="71af2-770">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="71af2-771">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="71af2-771">- Content</span></span><br><span data-ttu-id="71af2-772">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-772">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="71af2-773">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="71af2-773">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="71af2-774">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="71af2-774">- ActiveView</span></span><br><span data-ttu-id="71af2-775">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="71af2-775">
         - CompressedFile</span></span><br><span data-ttu-id="71af2-776">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-776">
         - DocumentEvents</span></span><br><span data-ttu-id="71af2-777">
         - File</span><span class="sxs-lookup"><span data-stu-id="71af2-777">
         - File</span></span><br><span data-ttu-id="71af2-778">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-778">
         - ImageCoercion</span></span><br><span data-ttu-id="71af2-779">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="71af2-779">
         - PdfFile</span></span><br><span data-ttu-id="71af2-780">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-780">
         - Selection</span></span><br><span data-ttu-id="71af2-781">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-781">
         - Settings</span></span><br><span data-ttu-id="71af2-782">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-782">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="71af2-783">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="71af2-783">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="71af2-784">OneNote</span><span class="sxs-lookup"><span data-stu-id="71af2-784">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="71af2-785">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="71af2-785">Platform</span></span></th>
    <th><span data-ttu-id="71af2-786">拡張点</span><span class="sxs-lookup"><span data-stu-id="71af2-786">Extension points</span></span></th>
    <th><span data-ttu-id="71af2-787">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="71af2-787">API requirement sets</span></span></th>
    <th><span data-ttu-id="71af2-788"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="71af2-788"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-789">Office Online</span><span class="sxs-lookup"><span data-stu-id="71af2-789">Office Online</span></span></td>
    <td> <span data-ttu-id="71af2-790">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="71af2-790">- Content</span></span><br><span data-ttu-id="71af2-791">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-791">
         - TaskPane</span></span><br><span data-ttu-id="71af2-792">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="71af2-792">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="71af2-793">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-793">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="71af2-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="71af2-795">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="71af2-795">- DocumentEvents</span></span><br><span data-ttu-id="71af2-796">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-796">
         - HtmlCoercion</span></span><br><span data-ttu-id="71af2-797">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-797">
         - ImageCoercion</span></span><br><span data-ttu-id="71af2-798">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="71af2-798">
         - Settings</span></span><br><span data-ttu-id="71af2-799">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-799">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="71af2-800">Project</span><span class="sxs-lookup"><span data-stu-id="71af2-800">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="71af2-801">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="71af2-801">Platform</span></span></th>
    <th><span data-ttu-id="71af2-802">拡張点</span><span class="sxs-lookup"><span data-stu-id="71af2-802">Extension points</span></span></th>
    <th><span data-ttu-id="71af2-803">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="71af2-803">API requirement sets</span></span></th>
    <th><span data-ttu-id="71af2-804"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="71af2-804"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-805">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="71af2-805">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="71af2-806">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-806">- TaskPane</span></span></td>
    <td> <span data-ttu-id="71af2-807">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-807">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="71af2-808">- Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-808">- Selection</span></span><br><span data-ttu-id="71af2-809">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-809">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-810">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="71af2-810">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="71af2-811">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-811">- TaskPane</span></span></td>
    <td> <span data-ttu-id="71af2-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="71af2-813">- Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-813">- Selection</span></span><br><span data-ttu-id="71af2-814">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-814">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="71af2-815">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="71af2-815">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="71af2-816">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="71af2-816">- TaskPane</span></span></td>
    <td> <span data-ttu-id="71af2-817">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="71af2-817">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="71af2-818">- Selection</span><span class="sxs-lookup"><span data-stu-id="71af2-818">- Selection</span></span><br><span data-ttu-id="71af2-819">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="71af2-819">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="71af2-820">関連項目</span><span class="sxs-lookup"><span data-stu-id="71af2-820">See also</span></span>

- [<span data-ttu-id="71af2-821">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="71af2-821">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="71af2-822">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="71af2-822">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="71af2-823">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="71af2-823">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="71af2-824">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="71af2-824">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
