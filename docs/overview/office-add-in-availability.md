---
title: Office アドインのホストとプラットフォームの可用性
description: Excel、Word、Outlook、PowerPoint、および OneNote のサポートされる要件セット。
ms.date: 10/03/2018
ms.openlocfilehash: 39a80f322c282e29e6e8c4363f0c82522b33b75d
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579927"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="2072c-103">Office アドインのホストとプラットフォームの可用性</span><span class="sxs-lookup"><span data-stu-id="2072c-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="2072c-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API の要件セットが含まれています。</span><span class="sxs-lookup"><span data-stu-id="2072c-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="2072c-p102">表のセルにアスタリスク ( \* ) が含まれる場合は、準備中を意味します。Project または Access の要件セットについては、「[Office の共有要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2072c-p102">If a table cell contains an asterisk ( \* ), that means we’re working on it. For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="2072c-p103">MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。このバージョンには、ExcelApi 1.1、WordApi 1.1、および共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="2072c-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="2072c-110">Excel</span><span class="sxs-lookup"><span data-stu-id="2072c-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="2072c-111">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="2072c-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="2072c-112">拡張点</span><span class="sxs-lookup"><span data-stu-id="2072c-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="2072c-113">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="2072c-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="2072c-114"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="2072c-114"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="2072c-115">Office Online</span></span></td>
    <td> <span data-ttu-id="2072c-116">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-116">- Taskpane</span></span><br><span data-ttu-id="2072c-117">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="2072c-117">
        - Content</span></span><br><span data-ttu-id="2072c-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="2072c-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="2072c-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2072c-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2072c-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="2072c-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2072c-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="2072c-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2072c-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="2072c-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2072c-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="2072c-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2072c-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="2072c-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2072c-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="2072c-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="2072c-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="2072c-127">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-127">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="2072c-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-128">
        -BindingEvents</span></span><br><span data-ttu-id="2072c-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2072c-129">
        -CompressedFile</span></span><br><span data-ttu-id="2072c-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-130">
        -DocumentEvents</span></span><br><span data-ttu-id="2072c-131">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-131">
        - File</span></span><br><span data-ttu-id="2072c-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-132">
        -MatrixBindings</span></span><br><span data-ttu-id="2072c-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-133">
        -MatrixCoercion</span></span><br><span data-ttu-id="2072c-134">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-134">
        - Selection</span></span><br><span data-ttu-id="2072c-135">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-135">
        - Settings</span></span><br><span data-ttu-id="2072c-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-136">
        -TableBindings</span></span><br><span data-ttu-id="2072c-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-137">
        -TableCoercion</span></span><br><span data-ttu-id="2072c-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-138">
        -TextBindings</span></span><br><span data-ttu-id="2072c-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-139">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-140">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="2072c-140">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="2072c-141">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-141">
        - Taskpane</span></span><br><span data-ttu-id="2072c-142">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="2072c-142">
        - Content</span></span></td>
    <td>  <span data-ttu-id="2072c-143">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-143">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="2072c-144">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-144">
        -BindingEvents</span></span><br><span data-ttu-id="2072c-145">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2072c-145">
        -CompressedFile</span></span><br><span data-ttu-id="2072c-146">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-146">
        -DocumentEvents</span></span><br><span data-ttu-id="2072c-147">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-147">
        - File</span></span><br><span data-ttu-id="2072c-148">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-148">
        -ImageCoercion</span></span><br><span data-ttu-id="2072c-149">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-149">
        -MatrixBindings</span></span><br><span data-ttu-id="2072c-150">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-150">
        -MatrixCoercion</span></span><br><span data-ttu-id="2072c-151">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-151">
        - Selection</span></span><br><span data-ttu-id="2072c-152">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-152">
        - Settings</span></span><br><span data-ttu-id="2072c-153">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-153">
        -TableBindings</span></span><br><span data-ttu-id="2072c-154">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-154">
        -TableCoercion</span></span><br><span data-ttu-id="2072c-155">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-155">
        -TextBindings</span></span><br><span data-ttu-id="2072c-156">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-156">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-157">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2072c-157">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="2072c-158">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-158">- Taskpane</span></span><br><span data-ttu-id="2072c-159">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="2072c-159">
        - Content</span></span><br><span data-ttu-id="2072c-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="2072c-161">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-161">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2072c-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2072c-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="2072c-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2072c-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="2072c-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2072c-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="2072c-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2072c-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="2072c-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2072c-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="2072c-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2072c-167">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="2072c-168">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="2072c-168">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="2072c-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="2072c-170">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-170">-BindingEvents</span></span><br><span data-ttu-id="2072c-171">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2072c-171">
        -CompressedFile</span></span><br><span data-ttu-id="2072c-172">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-172">
        -DocumentEvents</span></span><br><span data-ttu-id="2072c-173">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-173">
        - File</span></span><br><span data-ttu-id="2072c-174">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-174">
        -ImageCoercion</span></span><br><span data-ttu-id="2072c-175">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-175">
        -MatrixBindings</span></span><br><span data-ttu-id="2072c-176">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-176">
        -MatrixCoercion</span></span><br><span data-ttu-id="2072c-177">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-177">
        - Selection</span></span><br><span data-ttu-id="2072c-178">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-178">
        - Settings</span></span><br><span data-ttu-id="2072c-179">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-179">
        -TableBindings</span></span><br><span data-ttu-id="2072c-180">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-180">
        -TableCoercion</span></span><br><span data-ttu-id="2072c-181">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-181">
        -TextBindings</span></span><br><span data-ttu-id="2072c-182">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-182">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-183">Windows 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2072c-183">Office for Windows Desktop.</span></span></td>
    <td><span data-ttu-id="2072c-184">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-184">- Taskpane</span></span><br><span data-ttu-id="2072c-185">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="2072c-185">
        - Content</span></span><br><span data-ttu-id="2072c-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="2072c-187">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-187">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2072c-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2072c-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="2072c-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2072c-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="2072c-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2072c-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="2072c-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2072c-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="2072c-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2072c-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="2072c-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2072c-193">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="2072c-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="2072c-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="2072c-195">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-195">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="2072c-196">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-196">-BindingEvents</span></span><br><span data-ttu-id="2072c-197">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2072c-197">
        -CompressedFile</span></span><br><span data-ttu-id="2072c-198">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-198">
        -DocumentEvents</span></span><br><span data-ttu-id="2072c-199">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-199">
        - File</span></span><br><span data-ttu-id="2072c-200">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-200">
        -ImageCoercion</span></span><br><span data-ttu-id="2072c-201">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-201">
        -MatrixBindings</span></span><br><span data-ttu-id="2072c-202">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-202">
        -MatrixCoercion</span></span><br><span data-ttu-id="2072c-203">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-203">
        - Selection</span></span><br><span data-ttu-id="2072c-204">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-204">
        - Settings</span></span><br><span data-ttu-id="2072c-205">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-205">
        -TableBindings</span></span><br><span data-ttu-id="2072c-206">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-206">
        -TableCoercion</span></span><br><span data-ttu-id="2072c-207">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-207">
        -TextBindings</span></span><br><span data-ttu-id="2072c-208">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-208">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-209">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="2072c-209">Office for iOS</span></span></td>
    <td><span data-ttu-id="2072c-210">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-210">- Taskpane</span></span><br><span data-ttu-id="2072c-211">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="2072c-211">
        - Content</span></span></td>
    <td><span data-ttu-id="2072c-212">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-212">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2072c-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2072c-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="2072c-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2072c-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="2072c-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2072c-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="2072c-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2072c-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="2072c-217">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2072c-217">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="2072c-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2072c-218">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="2072c-219">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="2072c-219">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="2072c-220">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-220">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="2072c-221">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-221">-BindingEvents</span></span><br><span data-ttu-id="2072c-222">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2072c-222">
        -CompressedFile</span></span><br><span data-ttu-id="2072c-223">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-223">
        -DocumentEvents</span></span><br><span data-ttu-id="2072c-224">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-224">
        - File</span></span><br><span data-ttu-id="2072c-225">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-225">
        -ImageCoercion</span></span><br><span data-ttu-id="2072c-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-226">
        -MatrixBindings</span></span><br><span data-ttu-id="2072c-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-227">
        -MatrixCoercion</span></span><br><span data-ttu-id="2072c-228">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-228">
        - Selection</span></span><br><span data-ttu-id="2072c-229">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-229">
        - Settings</span></span><br><span data-ttu-id="2072c-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-230">
        -TableBindings</span></span><br><span data-ttu-id="2072c-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-231">
        -TableCoercion</span></span><br><span data-ttu-id="2072c-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-232">
        -TextBindings</span></span><br><span data-ttu-id="2072c-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-233">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-234">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="2072c-234">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="2072c-235">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-235">- Taskpane</span></span><br><span data-ttu-id="2072c-236">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="2072c-236">
        - Content</span></span><br><span data-ttu-id="2072c-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="2072c-238">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-238">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2072c-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2072c-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="2072c-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2072c-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="2072c-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2072c-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="2072c-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2072c-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="2072c-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2072c-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="2072c-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2072c-244">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="2072c-245">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="2072c-245">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="2072c-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="2072c-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-247">-BindingEvents</span></span><br><span data-ttu-id="2072c-248">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2072c-248">
        -CompressedFile</span></span><br><span data-ttu-id="2072c-249">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-249">
        -DocumentEvents</span></span><br><span data-ttu-id="2072c-250">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-250">
        - File</span></span><br><span data-ttu-id="2072c-251">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-251">
        -ImageCoercion</span></span><br><span data-ttu-id="2072c-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-252">
        -MatrixBindings</span></span><br><span data-ttu-id="2072c-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-253">
        -MatrixCoercion</span></span><br><span data-ttu-id="2072c-254">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2072c-254">
        -PdfFile</span></span><br><span data-ttu-id="2072c-255">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-255">
        - Selection</span></span><br><span data-ttu-id="2072c-256">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-256">
        - Settings</span></span><br><span data-ttu-id="2072c-257">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-257">
        -TableBindings</span></span><br><span data-ttu-id="2072c-258">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-258">
        -TableCoercion</span></span><br><span data-ttu-id="2072c-259">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-259">
        -TextBindings</span></span><br><span data-ttu-id="2072c-260">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-260">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-261">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="2072c-261">Office for Mac</span></span></td>
    <td><span data-ttu-id="2072c-262">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-262">- Taskpane</span></span><br><span data-ttu-id="2072c-263">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="2072c-263">
        - Content</span></span><br><span data-ttu-id="2072c-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="2072c-265">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-265">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="2072c-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2072c-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="2072c-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2072c-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="2072c-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2072c-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="2072c-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2072c-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="2072c-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2072c-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="2072c-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2072c-271">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="2072c-272">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="2072c-272">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="2072c-273">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-273">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="2072c-274">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-274">-BindingEvents</span></span><br><span data-ttu-id="2072c-275">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2072c-275">
        -CompressedFile</span></span><br><span data-ttu-id="2072c-276">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-276">
        -DocumentEvents</span></span><br><span data-ttu-id="2072c-277">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-277">
        - File</span></span><br><span data-ttu-id="2072c-278">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-278">
        -ImageCoercion</span></span><br><span data-ttu-id="2072c-279">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-279">
        -MatrixBindings</span></span><br><span data-ttu-id="2072c-280">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-280">
        -MatrixCoercion</span></span><br><span data-ttu-id="2072c-281">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2072c-281">
        -PdfFile</span></span><br><span data-ttu-id="2072c-282">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-282">
        - Selection</span></span><br><span data-ttu-id="2072c-283">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-283">
        - Settings</span></span><br><span data-ttu-id="2072c-284">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-284">
        -TableBindings</span></span><br><span data-ttu-id="2072c-285">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-285">
        -TableCoercion</span></span><br><span data-ttu-id="2072c-286">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-286">
        -TextBindings</span></span><br><span data-ttu-id="2072c-287">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-287">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="2072c-288">Outlook</span><span class="sxs-lookup"><span data-stu-id="2072c-288">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="2072c-289">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="2072c-289">Platform</span></span></th>
    <th><span data-ttu-id="2072c-290">拡張点</span><span class="sxs-lookup"><span data-stu-id="2072c-290">Extension points</span></span></th>
    <th><span data-ttu-id="2072c-291">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="2072c-291">API requirement sets</span></span></th>
    <th><span data-ttu-id="2072c-292"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="2072c-292"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-293">Office Online</span><span class="sxs-lookup"><span data-stu-id="2072c-293">Office Online</span></span></td>
    <td> <span data-ttu-id="2072c-294">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="2072c-294">- Mail Read</span></span><br><span data-ttu-id="2072c-295">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="2072c-295">
      - Mail Compose</span></span><br><span data-ttu-id="2072c-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2072c-297">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-297">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2072c-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2072c-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2072c-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2072c-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2072c-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2072c-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2072c-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2072c-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="2072c-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2072c-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="2072c-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2072c-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="2072c-304">使用不可</span><span class="sxs-lookup"><span data-stu-id="2072c-304">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-305">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="2072c-305">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="2072c-306">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="2072c-306">- Mail Read</span></span><br><span data-ttu-id="2072c-307">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="2072c-307">
      - Mail Compose</span></span><br><span data-ttu-id="2072c-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2072c-309">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-309">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2072c-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2072c-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2072c-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2072c-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2072c-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2072c-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="2072c-313">使用不可</span><span class="sxs-lookup"><span data-stu-id="2072c-313">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-314">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2072c-314">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="2072c-315">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="2072c-315">- Mail Read</span></span><br><span data-ttu-id="2072c-316">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="2072c-316">
      - Mail Compose</span></span><br><span data-ttu-id="2072c-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="2072c-318">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="2072c-318">
      - Modules</span></span></td>
    <td> <span data-ttu-id="2072c-319">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-319">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2072c-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2072c-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2072c-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2072c-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2072c-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2072c-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2072c-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2072c-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="2072c-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2072c-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="2072c-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2072c-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="2072c-326">使用不可</span><span class="sxs-lookup"><span data-stu-id="2072c-326">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-327">Windows 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2072c-327">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="2072c-328">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="2072c-328">- Mail Read</span></span><br><span data-ttu-id="2072c-329">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="2072c-329">
      - Mail Compose</span></span><br><span data-ttu-id="2072c-330">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-330">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="2072c-331">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="2072c-331">
      - Modules</span></span></td>
    <td> <span data-ttu-id="2072c-332">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-332">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2072c-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2072c-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2072c-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2072c-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2072c-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2072c-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2072c-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2072c-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="2072c-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2072c-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="2072c-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2072c-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="2072c-339">使用不可</span><span class="sxs-lookup"><span data-stu-id="2072c-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-340">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="2072c-340">Office for iOS</span></span></td>
    <td> <span data-ttu-id="2072c-341">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="2072c-341">- Mail Read</span></span><br><span data-ttu-id="2072c-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2072c-343">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-343">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2072c-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2072c-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2072c-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2072c-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2072c-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2072c-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2072c-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2072c-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="2072c-348">使用不可</span><span class="sxs-lookup"><span data-stu-id="2072c-348">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-349">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="2072c-349">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="2072c-350">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="2072c-350">- Mail Read</span></span><br><span data-ttu-id="2072c-351">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="2072c-351">
      - Mail Compose</span></span><br><span data-ttu-id="2072c-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2072c-353">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-353">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2072c-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2072c-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2072c-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2072c-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2072c-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2072c-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2072c-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2072c-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="2072c-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2072c-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="2072c-359">使用不可</span><span class="sxs-lookup"><span data-stu-id="2072c-359">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-360">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="2072c-360">Office for Mac</span></span></td>
    <td> <span data-ttu-id="2072c-361">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="2072c-361">- Mail Read</span></span><br><span data-ttu-id="2072c-362">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="2072c-362">
      - Mail Compose</span></span><br><span data-ttu-id="2072c-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2072c-364">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-364">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2072c-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2072c-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2072c-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2072c-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2072c-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2072c-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2072c-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2072c-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="2072c-369">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="2072c-369">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="2072c-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="2072c-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="2072c-371">使用不可</span><span class="sxs-lookup"><span data-stu-id="2072c-371">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-372">Android 用 Office</span><span class="sxs-lookup"><span data-stu-id="2072c-372">Office for Android</span></span></td>
    <td> <span data-ttu-id="2072c-373">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="2072c-373">- Mail Read</span></span><br><span data-ttu-id="2072c-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2072c-375">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-375">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="2072c-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2072c-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="2072c-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2072c-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="2072c-378">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="2072c-378">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="2072c-379">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="2072c-379">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="2072c-380">使用不可</span><span class="sxs-lookup"><span data-stu-id="2072c-380">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="2072c-381">Word</span><span class="sxs-lookup"><span data-stu-id="2072c-381">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="2072c-382">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="2072c-382">Platform</span></span></th>
    <th><span data-ttu-id="2072c-383">拡張点</span><span class="sxs-lookup"><span data-stu-id="2072c-383">Extension points</span></span></th>
    <th><span data-ttu-id="2072c-384">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="2072c-384">API requirement sets</span></span></th>
    <th><span data-ttu-id="2072c-385"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="2072c-385"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-386">Office Online</span><span class="sxs-lookup"><span data-stu-id="2072c-386">Office Online</span></span></td>
    <td> <span data-ttu-id="2072c-387">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-387">- Taskpane</span></span><br><span data-ttu-id="2072c-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2072c-389">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-389">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2072c-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2072c-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="2072c-391">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2072c-391">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="2072c-392">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-392">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="2072c-393">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-393">-BindingEvents</span></span><br><span data-ttu-id="2072c-394">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="2072c-394">
         -CustomXmlParts</span></span><br><span data-ttu-id="2072c-395">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-395">
         -DocumentEvents</span></span><br><span data-ttu-id="2072c-396">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-396">
         - File</span></span><br><span data-ttu-id="2072c-397">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-397">
         -HtmlCoercion</span></span><br><span data-ttu-id="2072c-398">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-398">
         -ImageCoercion</span></span><br><span data-ttu-id="2072c-399">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-399">
         -MatrixBindings</span></span><br><span data-ttu-id="2072c-400">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-400">
         -MatrixCoercion</span></span><br><span data-ttu-id="2072c-401">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-401">
         -OoxmlCoercion</span></span><br><span data-ttu-id="2072c-402">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2072c-402">
         -PdfFile</span></span><br><span data-ttu-id="2072c-403">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-403">
         - Selection</span></span><br><span data-ttu-id="2072c-404">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-404">
         - Settings</span></span><br><span data-ttu-id="2072c-405">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-405">
         -TableBindings</span></span><br><span data-ttu-id="2072c-406">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-406">
         -TableCoercion</span></span><br><span data-ttu-id="2072c-407">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-407">
         -TextBindings</span></span><br><span data-ttu-id="2072c-408">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-408">
         -TextCoercion</span></span><br><span data-ttu-id="2072c-409">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2072c-409">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-410">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="2072c-410">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="2072c-411">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-411">- Taskpane</span></span></td>
    <td> <span data-ttu-id="2072c-412">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-412">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="2072c-413">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-413">-BindingEvents</span></span><br><span data-ttu-id="2072c-414">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2072c-414">
         -CompressedFile</span></span><br><span data-ttu-id="2072c-415">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="2072c-415">
         -CustomXmlParts</span></span><br><span data-ttu-id="2072c-416">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-416">
         -DocumentEvents</span></span><br><span data-ttu-id="2072c-417">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-417">
         - File</span></span><br><span data-ttu-id="2072c-418">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-418">
         -HtmlCoercion</span></span><br><span data-ttu-id="2072c-419">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-419">
         -ImageCoercion</span></span><br><span data-ttu-id="2072c-420">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-420">
         -MatrixBindings</span></span><br><span data-ttu-id="2072c-421">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-421">
         -MatrixCoercion</span></span><br><span data-ttu-id="2072c-422">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-422">
         -OoxmlCoercion</span></span><br><span data-ttu-id="2072c-423">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2072c-423">
         -PdfFile</span></span><br><span data-ttu-id="2072c-424">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-424">
         - Selection</span></span><br><span data-ttu-id="2072c-425">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-425">
         - Settings</span></span><br><span data-ttu-id="2072c-426">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-426">
         -TableBindings</span></span><br><span data-ttu-id="2072c-427">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-427">
         -TableCoercion</span></span><br><span data-ttu-id="2072c-428">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-428">
         -TextBindings</span></span><br><span data-ttu-id="2072c-429">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-429">
         -TextCoercion</span></span><br><span data-ttu-id="2072c-430">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2072c-430">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-431">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2072c-431">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="2072c-432">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-432">- Taskpane</span></span><br><span data-ttu-id="2072c-433">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-433">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2072c-434">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-434">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2072c-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2072c-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="2072c-436">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2072c-436">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="2072c-437">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-437">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="2072c-438">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-438">-BindingEvents</span></span><br><span data-ttu-id="2072c-439">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2072c-439">
         -CompressedFile</span></span><br><span data-ttu-id="2072c-440">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="2072c-440">
         -CustomXmlParts</span></span><br><span data-ttu-id="2072c-441">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-441">
         -DocumentEvents</span></span><br><span data-ttu-id="2072c-442">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-442">
         - File</span></span><br><span data-ttu-id="2072c-443">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-443">
         -HtmlCoercion</span></span><br><span data-ttu-id="2072c-444">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-444">
         -ImageCoercion</span></span><br><span data-ttu-id="2072c-445">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-445">
         -MatrixBindings</span></span><br><span data-ttu-id="2072c-446">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-446">
         -MatrixCoercion</span></span><br><span data-ttu-id="2072c-447">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-447">
         -OoxmlCoercion</span></span><br><span data-ttu-id="2072c-448">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2072c-448">
         -PdfFile</span></span><br><span data-ttu-id="2072c-449">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-449">
         - Selection</span></span><br><span data-ttu-id="2072c-450">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-450">
         - Settings</span></span><br><span data-ttu-id="2072c-451">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-451">
         -TableBindings</span></span><br><span data-ttu-id="2072c-452">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-452">
         -TableCoercion</span></span><br><span data-ttu-id="2072c-453">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-453">
         -TextBindings</span></span><br><span data-ttu-id="2072c-454">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-454">
         -TextCoercion</span></span><br><span data-ttu-id="2072c-455">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2072c-455">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-456">Windows 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2072c-456">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="2072c-457">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-457">- Taskpane</span></span><br><span data-ttu-id="2072c-458">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-458">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2072c-459">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-459">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2072c-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2072c-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="2072c-461">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2072c-461">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="2072c-462">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-462">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="2072c-463">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-463">-BindingEvents</span></span><br><span data-ttu-id="2072c-464">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2072c-464">
         -CompressedFile</span></span><br><span data-ttu-id="2072c-465">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="2072c-465">
         -CustomXmlParts</span></span><br><span data-ttu-id="2072c-466">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-466">
         -DocumentEvents</span></span><br><span data-ttu-id="2072c-467">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-467">
         - File</span></span><br><span data-ttu-id="2072c-468">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-468">
         -HtmlCoercion</span></span><br><span data-ttu-id="2072c-469">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-469">
         -ImageCoercion</span></span><br><span data-ttu-id="2072c-470">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-470">
         -MatrixBindings</span></span><br><span data-ttu-id="2072c-471">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-471">
         -MatrixCoercion</span></span><br><span data-ttu-id="2072c-472">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-472">
         -OoxmlCoercion</span></span><br><span data-ttu-id="2072c-473">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2072c-473">
         -PdfFile</span></span><br><span data-ttu-id="2072c-474">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-474">
         - Selection</span></span><br><span data-ttu-id="2072c-475">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-475">
         - Settings</span></span><br><span data-ttu-id="2072c-476">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-476">
         -TableBindings</span></span><br><span data-ttu-id="2072c-477">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-477">
         -TableCoercion</span></span><br><span data-ttu-id="2072c-478">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-478">
         -TextBindings</span></span><br><span data-ttu-id="2072c-479">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-479">
         -TextCoercion</span></span><br><span data-ttu-id="2072c-480">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2072c-480">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-481">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="2072c-481">Office for iOS</span></span></td>
    <td> <span data-ttu-id="2072c-482">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-482">- Taskpane</span></span></td>
    <td> <span data-ttu-id="2072c-483">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-483">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2072c-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2072c-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="2072c-485">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2072c-485">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="2072c-486">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="2072c-486">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="2072c-487">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-487">-BindingEvents</span></span><br><span data-ttu-id="2072c-488">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2072c-488">
         -CompressedFile</span></span><br><span data-ttu-id="2072c-489">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="2072c-489">
         -CustomXmlParts</span></span><br><span data-ttu-id="2072c-490">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-490">
         -DocumentEvents</span></span><br><span data-ttu-id="2072c-491">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-491">
         - File</span></span><br><span data-ttu-id="2072c-492">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-492">
         -HtmlCoercion</span></span><br><span data-ttu-id="2072c-493">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-493">
         -ImageCoercion</span></span><br><span data-ttu-id="2072c-494">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-494">
         -MatrixBindings</span></span><br><span data-ttu-id="2072c-495">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-495">
         -MatrixCoercion</span></span><br><span data-ttu-id="2072c-496">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-496">
         -OoxmlCoercion</span></span><br><span data-ttu-id="2072c-497">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2072c-497">
         -PdfFile</span></span><br><span data-ttu-id="2072c-498">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-498">
         - Selection</span></span><br><span data-ttu-id="2072c-499">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-499">
         - Settings</span></span><br><span data-ttu-id="2072c-500">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-500">
         -TableBindings</span></span><br><span data-ttu-id="2072c-501">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-501">
         -TableCoercion</span></span><br><span data-ttu-id="2072c-502">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-502">
         -TextBindings</span></span><br><span data-ttu-id="2072c-503">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-503">
         -TextCoercion</span></span><br><span data-ttu-id="2072c-504">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2072c-504">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-505">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="2072c-505">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="2072c-506">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-506">- Taskpane</span></span><br><span data-ttu-id="2072c-507">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-507">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2072c-508">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-508">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2072c-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2072c-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="2072c-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2072c-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="2072c-511">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="2072c-511">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="2072c-512">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-512">-BindingEvents</span></span><br><span data-ttu-id="2072c-513">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2072c-513">
         -CompressedFile</span></span><br><span data-ttu-id="2072c-514">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="2072c-514">
         -CustomXmlParts</span></span><br><span data-ttu-id="2072c-515">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-515">
         -DocumentEvents</span></span><br><span data-ttu-id="2072c-516">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-516">
         - File</span></span><br><span data-ttu-id="2072c-517">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-517">
         -HtmlCoercion</span></span><br><span data-ttu-id="2072c-518">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-518">
         -ImageCoercion</span></span><br><span data-ttu-id="2072c-519">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-519">
         -MatrixBindings</span></span><br><span data-ttu-id="2072c-520">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-520">
         -MatrixCoercion</span></span><br><span data-ttu-id="2072c-521">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-521">
         -OoxmlCoercion</span></span><br><span data-ttu-id="2072c-522">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2072c-522">
         -PdfFile</span></span><br><span data-ttu-id="2072c-523">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-523">
         - Selection</span></span><br><span data-ttu-id="2072c-524">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-524">
         - Settings</span></span><br><span data-ttu-id="2072c-525">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-525">
         -TableBindings</span></span><br><span data-ttu-id="2072c-526">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-526">
         -TableCoercion</span></span><br><span data-ttu-id="2072c-527">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-527">
         -TextBindings</span></span><br><span data-ttu-id="2072c-528">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-528">
         -TextCoercion</span></span><br><span data-ttu-id="2072c-529">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2072c-529">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-530">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="2072c-530">Office for Mac</span></span></td>
    <td> <span data-ttu-id="2072c-531">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-531">- Taskpane</span></span><br><span data-ttu-id="2072c-532">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-532">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2072c-533">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-533">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="2072c-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="2072c-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="2072c-535">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="2072c-535">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="2072c-536">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="2072c-536">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="2072c-537">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-537">-BindingEvents</span></span><br><span data-ttu-id="2072c-538">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2072c-538">
         -CompressedFile</span></span><br><span data-ttu-id="2072c-539">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="2072c-539">
         -CustomXmlParts</span></span><br><span data-ttu-id="2072c-540">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-540">
         -DocumentEvents</span></span><br><span data-ttu-id="2072c-541">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-541">
         - File</span></span><br><span data-ttu-id="2072c-542">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-542">
         -HtmlCoercion</span></span><br><span data-ttu-id="2072c-543">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-543">
         -ImageCoercion</span></span><br><span data-ttu-id="2072c-544">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-544">
         -MatrixBindings</span></span><br><span data-ttu-id="2072c-545">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-545">
         -MatrixCoercion</span></span><br><span data-ttu-id="2072c-546">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-546">
         -OoxmlCoercion</span></span><br><span data-ttu-id="2072c-547">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2072c-547">
         -PdfFile</span></span><br><span data-ttu-id="2072c-548">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-548">
         - Selection</span></span><br><span data-ttu-id="2072c-549">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-549">
         - Settings</span></span><br><span data-ttu-id="2072c-550">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-550">
         -TableBindings</span></span><br><span data-ttu-id="2072c-551">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-551">
         -TableCoercion</span></span><br><span data-ttu-id="2072c-552">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="2072c-552">
         -TextBindings</span></span><br><span data-ttu-id="2072c-553">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-553">
         -TextCoercion</span></span><br><span data-ttu-id="2072c-554">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="2072c-554">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="2072c-555">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="2072c-555">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="2072c-556">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="2072c-556">Platform</span></span></th>
    <th><span data-ttu-id="2072c-557">拡張点</span><span class="sxs-lookup"><span data-stu-id="2072c-557">Extension points</span></span></th>
    <th><span data-ttu-id="2072c-558">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="2072c-558">API requirement sets</span></span></th>
    <th><span data-ttu-id="2072c-559"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="2072c-559"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-560">Office Online</span><span class="sxs-lookup"><span data-stu-id="2072c-560">Office Online</span></span></td>
    <td> <span data-ttu-id="2072c-561">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="2072c-561">- Content</span></span><br><span data-ttu-id="2072c-562">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-562">
         - Taskpane</span></span><br><span data-ttu-id="2072c-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2072c-564">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-564">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="2072c-565">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2072c-565">-ActiveView</span></span><br><span data-ttu-id="2072c-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2072c-566">
         -CompressedFile</span></span><br><span data-ttu-id="2072c-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-567">
         -DocumentEvents</span></span><br><span data-ttu-id="2072c-568">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-568">
         - File</span></span><br><span data-ttu-id="2072c-569">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-569">
         -ImageCoercion</span></span><br><span data-ttu-id="2072c-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2072c-570">
         -PdfFile</span></span><br><span data-ttu-id="2072c-571">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-571">
         - Selection</span></span><br><span data-ttu-id="2072c-572">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-572">
         - Settings</span></span><br><span data-ttu-id="2072c-573">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-573">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-574">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="2072c-574">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="2072c-575">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="2072c-575">- Content</span></span><br><span data-ttu-id="2072c-576">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-576">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="2072c-577">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="2072c-577">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="2072c-578">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2072c-578">-ActiveView</span></span><br><span data-ttu-id="2072c-579">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2072c-579">
         -CompressedFile</span></span><br><span data-ttu-id="2072c-580">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-580">
         -DocumentEvents</span></span><br><span data-ttu-id="2072c-581">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-581">
         - File</span></span><br><span data-ttu-id="2072c-582">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-582">
         -ImageCoercion</span></span><br><span data-ttu-id="2072c-583">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2072c-583">
         -PdfFile</span></span><br><span data-ttu-id="2072c-584">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-584">
         - Selection</span></span><br><span data-ttu-id="2072c-585">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-585">
         - Settings</span></span><br><span data-ttu-id="2072c-586">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-586">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-587">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="2072c-587">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="2072c-588">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="2072c-588">- Content</span></span><br><span data-ttu-id="2072c-589">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-589">
         - Taskpane</span></span><br><span data-ttu-id="2072c-590">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-590">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2072c-591">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-591">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="2072c-592">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2072c-592">-ActiveView</span></span><br><span data-ttu-id="2072c-593">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2072c-593">
         -CompressedFile</span></span><br><span data-ttu-id="2072c-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-594">
         -DocumentEvents</span></span><br><span data-ttu-id="2072c-595">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-595">
         - File</span></span><br><span data-ttu-id="2072c-596">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-596">
         -ImageCoercion</span></span><br><span data-ttu-id="2072c-597">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2072c-597">
         -PdfFile</span></span><br><span data-ttu-id="2072c-598">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-598">
         - Selection</span></span><br><span data-ttu-id="2072c-599">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-599">
         - Settings</span></span><br><span data-ttu-id="2072c-600">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-600">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-601">Windows 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="2072c-601">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="2072c-602">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="2072c-602">- Content</span></span><br><span data-ttu-id="2072c-603">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-603">
         - Taskpane</span></span><br><span data-ttu-id="2072c-604">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-604">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2072c-605">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-605">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="2072c-606">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2072c-606">-ActiveView</span></span><br><span data-ttu-id="2072c-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2072c-607">
         -CompressedFile</span></span><br><span data-ttu-id="2072c-608">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-608">
         -DocumentEvents</span></span><br><span data-ttu-id="2072c-609">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-609">
         - File</span></span><br><span data-ttu-id="2072c-610">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-610">
         -ImageCoercion</span></span><br><span data-ttu-id="2072c-611">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2072c-611">
         -PdfFile</span></span><br><span data-ttu-id="2072c-612">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-612">
         - Selection</span></span><br><span data-ttu-id="2072c-613">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-613">
         - Settings</span></span><br><span data-ttu-id="2072c-614">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-614">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-615">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="2072c-615">Office for iOS</span></span></td>
    <td> <span data-ttu-id="2072c-616">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="2072c-616">- Content</span></span><br><span data-ttu-id="2072c-617">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-617">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="2072c-618">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-618">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="2072c-619">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2072c-619">-ActiveView</span></span><br><span data-ttu-id="2072c-620">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2072c-620">
         -CompressedFile</span></span><br><span data-ttu-id="2072c-621">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-621">
         -DocumentEvents</span></span><br><span data-ttu-id="2072c-622">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-622">
         - File</span></span><br><span data-ttu-id="2072c-623">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2072c-623">
         -PdfFile</span></span><br><span data-ttu-id="2072c-624">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-624">
         - Selection</span></span><br><span data-ttu-id="2072c-625">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-625">
         - Settings</span></span><br><span data-ttu-id="2072c-626">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-626">
         -TextCoercion</span></span><br><span data-ttu-id="2072c-627">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-627">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-628">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="2072c-628">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="2072c-629">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="2072c-629">- Content</span></span><br><span data-ttu-id="2072c-630">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-630">
         - Taskpane</span></span><br><span data-ttu-id="2072c-631">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-631">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2072c-632">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-632">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="2072c-633">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2072c-633">-ActiveView</span></span><br><span data-ttu-id="2072c-634">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2072c-634">
         -CompressedFile</span></span><br><span data-ttu-id="2072c-635">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-635">
         -DocumentEvents</span></span><br><span data-ttu-id="2072c-636">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-636">
         - File</span></span><br><span data-ttu-id="2072c-637">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-637">
         -ImageCoercion</span></span><br><span data-ttu-id="2072c-638">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2072c-638">
         -PdfFile</span></span><br><span data-ttu-id="2072c-639">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-639">
         - Selection</span></span><br><span data-ttu-id="2072c-640">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-640">
         - Settings</span></span><br><span data-ttu-id="2072c-641">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-641">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-642">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="2072c-642">Office for Mac</span></span></td>
    <td> <span data-ttu-id="2072c-643">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="2072c-643">- Content</span></span><br><span data-ttu-id="2072c-644">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-644">
         - Taskpane</span></span><br><span data-ttu-id="2072c-645">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-645">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2072c-646">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-646">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="2072c-647">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="2072c-647">-ActiveView</span></span><br><span data-ttu-id="2072c-648">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="2072c-648">
         -CompressedFile</span></span><br><span data-ttu-id="2072c-649">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-649">
         -DocumentEvents</span></span><br><span data-ttu-id="2072c-650">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="2072c-650">
         - File</span></span><br><span data-ttu-id="2072c-651">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-651">
         -ImageCoercion</span></span><br><span data-ttu-id="2072c-652">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="2072c-652">
         -PdfFile</span></span><br><span data-ttu-id="2072c-653">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="2072c-653">
         - Selection</span></span><br><span data-ttu-id="2072c-654">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-654">
         - Settings</span></span><br><span data-ttu-id="2072c-655">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-655">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="2072c-656">OneNote</span><span class="sxs-lookup"><span data-stu-id="2072c-656">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="2072c-657">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="2072c-657">Platform</span></span></th>
    <th><span data-ttu-id="2072c-658">拡張点</span><span class="sxs-lookup"><span data-stu-id="2072c-658">Extension points</span></span></th>
    <th><span data-ttu-id="2072c-659">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="2072c-659">API requirement sets</span></span></th>
    <th><span data-ttu-id="2072c-660"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="2072c-660"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="2072c-661">Office Online</span><span class="sxs-lookup"><span data-stu-id="2072c-661">Office Online</span></span></td>
    <td> <span data-ttu-id="2072c-662">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="2072c-662">- Content</span></span><br><span data-ttu-id="2072c-663">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="2072c-663">
         - Taskpane</span></span><br><span data-ttu-id="2072c-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="2072c-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="2072c-665">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-665">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="2072c-666">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="2072c-666">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="2072c-667">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="2072c-667">-DocumentEvents</span></span><br><span data-ttu-id="2072c-668">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-668">
         -HtmlCoercion</span></span><br><span data-ttu-id="2072c-669">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-669">
         -ImageCoercion</span></span><br><span data-ttu-id="2072c-670">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="2072c-670">
         - Settings</span></span><br><span data-ttu-id="2072c-671">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="2072c-671">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="2072c-672">関連項目</span><span class="sxs-lookup"><span data-stu-id="2072c-672">See also</span></span>

- [<span data-ttu-id="2072c-673">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="2072c-673">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="2072c-674">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="2072c-674">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="2072c-675">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="2072c-675">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="2072c-676">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="2072c-676">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
