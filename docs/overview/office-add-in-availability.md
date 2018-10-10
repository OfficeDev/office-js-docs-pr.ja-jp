---
title: Office アドインのホストとプラットフォームの可用性
description: Excel、Word、Outlook、PowerPoint、および OneNote のサポートされる要件セット。
ms.date: 10/03/2018
ms.openlocfilehash: 6f7b5b565773457e6cd8a9eee69eb304784a29a9
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459316"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="c77e2-103">Office アドインのホストとプラットフォームの可用性</span><span class="sxs-lookup"><span data-stu-id="c77e2-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="c77e2-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API の要件セットが含まれています。</span><span class="sxs-lookup"><span data-stu-id="c77e2-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="c77e2-p102">表のセルにアスタリスク ( \* ) が含まれる場合は、準備中を意味します。Project または Access の要件セットについては、「[Office の共有要件セット](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c77e2-p102">If a table cell contains an asterisk ( \* ), that means we’re working on it. For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="c77e2-p103">MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。このバージョンには、ExcelApi 1.1、WordApi 1.1、および共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="c77e2-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="c77e2-110">Excel</span><span class="sxs-lookup"><span data-stu-id="c77e2-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="c77e2-111">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="c77e2-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="c77e2-112">拡張点</span><span class="sxs-lookup"><span data-stu-id="c77e2-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="c77e2-113">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="c77e2-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="c77e2-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c77e2-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="c77e2-115">Office Online</span></span></td>
    <td> <span data-ttu-id="c77e2-116">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-116">- Taskpane</span></span><br><span data-ttu-id="c77e2-117">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c77e2-117">
        - Content</span></span><br><span data-ttu-id="c77e2-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="c77e2-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="c77e2-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c77e2-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c77e2-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c77e2-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c77e2-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c77e2-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c77e2-125">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="c77e2-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c77e2-127">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-127">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c77e2-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-128">
        -BindingEvents</span></span><br><span data-ttu-id="c77e2-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-129">
        -CompressedFile</span></span><br><span data-ttu-id="c77e2-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-130">
        -DocumentEvents</span></span><br><span data-ttu-id="c77e2-131">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="c77e2-131">
        - File</span></span><br><span data-ttu-id="c77e2-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-132">
        -MatrixBindings</span></span><br><span data-ttu-id="c77e2-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-133">
        -MatrixCoercion</span></span><br><span data-ttu-id="c77e2-134">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="c77e2-134">
        - Selection</span></span><br><span data-ttu-id="c77e2-135">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-135">
        - Settings</span></span><br><span data-ttu-id="c77e2-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-136">
        -TableBindings</span></span><br><span data-ttu-id="c77e2-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-137">
        -TableCoercion</span></span><br><span data-ttu-id="c77e2-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-138">
        -TextBindings</span></span><br><span data-ttu-id="c77e2-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-139">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-140">Windows 用 Office 2013</span><span class="sxs-lookup"><span data-stu-id="c77e2-140">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="c77e2-141">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-141">
        - Taskpane</span></span><br><span data-ttu-id="c77e2-142">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c77e2-142">
        - Content</span></span></td>
    <td>  <span data-ttu-id="c77e2-143">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-143">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c77e2-144">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-144">
        -BindingEvents</span></span><br><span data-ttu-id="c77e2-145">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-145">
        -CompressedFile</span></span><br><span data-ttu-id="c77e2-146">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-146">
        -DocumentEvents</span></span><br><span data-ttu-id="c77e2-147">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="c77e2-147">
        - File</span></span><br><span data-ttu-id="c77e2-148">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-148">
        -ImageCoercion</span></span><br><span data-ttu-id="c77e2-149">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-149">
        -MatrixBindings</span></span><br><span data-ttu-id="c77e2-150">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-150">
        -MatrixCoercion</span></span><br><span data-ttu-id="c77e2-151">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="c77e2-151">
        - Selection</span></span><br><span data-ttu-id="c77e2-152">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-152">
        - Settings</span></span><br><span data-ttu-id="c77e2-153">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-153">
        -TableBindings</span></span><br><span data-ttu-id="c77e2-154">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-154">
        -TableCoercion</span></span><br><span data-ttu-id="c77e2-155">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-155">
        -TextBindings</span></span><br><span data-ttu-id="c77e2-156">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-156">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-157">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c77e2-157">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="c77e2-158">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-158">- Taskpane</span></span><br><span data-ttu-id="c77e2-159">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c77e2-159">
        - Content</span></span><br><span data-ttu-id="c77e2-160">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-160">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c77e2-161">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-161">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c77e2-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c77e2-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c77e2-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c77e2-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c77e2-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c77e2-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-167">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="c77e2-168">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-168">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c77e2-169">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-169">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c77e2-170">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-170">-BindingEvents</span></span><br><span data-ttu-id="c77e2-171">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-171">
        -CompressedFile</span></span><br><span data-ttu-id="c77e2-172">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-172">
        -DocumentEvents</span></span><br><span data-ttu-id="c77e2-173">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="c77e2-173">
        - File</span></span><br><span data-ttu-id="c77e2-174">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-174">
        -ImageCoercion</span></span><br><span data-ttu-id="c77e2-175">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-175">
        -MatrixBindings</span></span><br><span data-ttu-id="c77e2-176">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-176">
        -MatrixCoercion</span></span><br><span data-ttu-id="c77e2-177">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="c77e2-177">
        - Selection</span></span><br><span data-ttu-id="c77e2-178">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-178">
        - Settings</span></span><br><span data-ttu-id="c77e2-179">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-179">
        -TableBindings</span></span><br><span data-ttu-id="c77e2-180">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-180">
        -TableCoercion</span></span><br><span data-ttu-id="c77e2-181">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-181">
        -TextBindings</span></span><br><span data-ttu-id="c77e2-182">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-182">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-183">Windows 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c77e2-183">Office for Windows Desktop.</span></span></td>
    <td><span data-ttu-id="c77e2-184">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-184">- Taskpane</span></span><br><span data-ttu-id="c77e2-185">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c77e2-185">
        - Content</span></span><br><span data-ttu-id="c77e2-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c77e2-187">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-187">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c77e2-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c77e2-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c77e2-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c77e2-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c77e2-192">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-192">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c77e2-193">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-193">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="c77e2-194">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-194">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c77e2-195">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-195">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c77e2-196">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-196">-BindingEvents</span></span><br><span data-ttu-id="c77e2-197">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-197">
        -CompressedFile</span></span><br><span data-ttu-id="c77e2-198">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-198">
        -DocumentEvents</span></span><br><span data-ttu-id="c77e2-199">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="c77e2-199">
        - File</span></span><br><span data-ttu-id="c77e2-200">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-200">
        -ImageCoercion</span></span><br><span data-ttu-id="c77e2-201">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-201">
        -MatrixBindings</span></span><br><span data-ttu-id="c77e2-202">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-202">
        -MatrixCoercion</span></span><br><span data-ttu-id="c77e2-203">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="c77e2-203">
        - Selection</span></span><br><span data-ttu-id="c77e2-204">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-204">
        - Settings</span></span><br><span data-ttu-id="c77e2-205">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-205">
        -TableBindings</span></span><br><span data-ttu-id="c77e2-206">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-206">
        -TableCoercion</span></span><br><span data-ttu-id="c77e2-207">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-207">
        -TextBindings</span></span><br><span data-ttu-id="c77e2-208">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-208">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-209">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="c77e2-209">Office for iOS</span></span></td>
    <td><span data-ttu-id="c77e2-210">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-210">- Taskpane</span></span><br><span data-ttu-id="c77e2-211">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c77e2-211">
        - Content</span></span></td>
    <td><span data-ttu-id="c77e2-212">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-212">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c77e2-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c77e2-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c77e2-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c77e2-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c77e2-217">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-217">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c77e2-218">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-218">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="c77e2-219">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-219">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c77e2-220">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-220">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c77e2-221">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-221">-BindingEvents</span></span><br><span data-ttu-id="c77e2-222">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-222">
        -CompressedFile</span></span><br><span data-ttu-id="c77e2-223">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-223">
        -DocumentEvents</span></span><br><span data-ttu-id="c77e2-224">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="c77e2-224">
        - File</span></span><br><span data-ttu-id="c77e2-225">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-225">
        -ImageCoercion</span></span><br><span data-ttu-id="c77e2-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-226">
        -MatrixBindings</span></span><br><span data-ttu-id="c77e2-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-227">
        -MatrixCoercion</span></span><br><span data-ttu-id="c77e2-228">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="c77e2-228">
        - Selection</span></span><br><span data-ttu-id="c77e2-229">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-229">
        - Settings</span></span><br><span data-ttu-id="c77e2-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-230">
        -TableBindings</span></span><br><span data-ttu-id="c77e2-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-231">
        -TableCoercion</span></span><br><span data-ttu-id="c77e2-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-232">
        -TextBindings</span></span><br><span data-ttu-id="c77e2-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-233">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-234">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c77e2-234">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="c77e2-235">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-235">- Taskpane</span></span><br><span data-ttu-id="c77e2-236">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c77e2-236">
        - Content</span></span><br><span data-ttu-id="c77e2-237">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-237">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c77e2-238">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-238">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c77e2-239">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-239">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c77e2-240">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-240">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c77e2-241">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-241">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c77e2-242">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-242">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c77e2-243">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-243">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c77e2-244">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-244">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="c77e2-245">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-245">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c77e2-246">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-246">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c77e2-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-247">-BindingEvents</span></span><br><span data-ttu-id="c77e2-248">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-248">
        -CompressedFile</span></span><br><span data-ttu-id="c77e2-249">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-249">
        -DocumentEvents</span></span><br><span data-ttu-id="c77e2-250">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="c77e2-250">
        - File</span></span><br><span data-ttu-id="c77e2-251">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-251">
        -ImageCoercion</span></span><br><span data-ttu-id="c77e2-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-252">
        -MatrixBindings</span></span><br><span data-ttu-id="c77e2-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-253">
        -MatrixCoercion</span></span><br><span data-ttu-id="c77e2-254">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-254">
        -PdfFile</span></span><br><span data-ttu-id="c77e2-255">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c77e2-255">
        - Selection</span></span><br><span data-ttu-id="c77e2-256">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-256">
        - Settings</span></span><br><span data-ttu-id="c77e2-257">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-257">
        -TableBindings</span></span><br><span data-ttu-id="c77e2-258">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-258">
        -TableCoercion</span></span><br><span data-ttu-id="c77e2-259">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-259">
        -TextBindings</span></span><br><span data-ttu-id="c77e2-260">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-260">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-261">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c77e2-261">Office for Mac</span></span></td>
    <td><span data-ttu-id="c77e2-262">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-262">- Taskpane</span></span><br><span data-ttu-id="c77e2-263">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c77e2-263">
        - Content</span></span><br><span data-ttu-id="c77e2-264">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-264">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c77e2-265">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-265">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c77e2-266">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-266">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c77e2-267">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-267">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c77e2-268">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-268">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c77e2-269">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-269">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c77e2-270">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-270">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c77e2-271">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-271">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="c77e2-272">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-272">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c77e2-273">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-273">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c77e2-274">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-274">-BindingEvents</span></span><br><span data-ttu-id="c77e2-275">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-275">
        -CompressedFile</span></span><br><span data-ttu-id="c77e2-276">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-276">
        -DocumentEvents</span></span><br><span data-ttu-id="c77e2-277">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="c77e2-277">
        - File</span></span><br><span data-ttu-id="c77e2-278">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-278">
        -ImageCoercion</span></span><br><span data-ttu-id="c77e2-279">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-279">
        -MatrixBindings</span></span><br><span data-ttu-id="c77e2-280">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-280">
        -MatrixCoercion</span></span><br><span data-ttu-id="c77e2-281">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-281">
        -PdfFile</span></span><br><span data-ttu-id="c77e2-282">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c77e2-282">
        - Selection</span></span><br><span data-ttu-id="c77e2-283">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-283">
        - Settings</span></span><br><span data-ttu-id="c77e2-284">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-284">
        -TableBindings</span></span><br><span data-ttu-id="c77e2-285">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-285">
        -TableCoercion</span></span><br><span data-ttu-id="c77e2-286">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-286">
        -TextBindings</span></span><br><span data-ttu-id="c77e2-287">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-287">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="c77e2-288">Outlook</span><span class="sxs-lookup"><span data-stu-id="c77e2-288">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c77e2-289">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="c77e2-289">Platform</span></span></th>
    <th><span data-ttu-id="c77e2-290">拡張点</span><span class="sxs-lookup"><span data-stu-id="c77e2-290">Extension points</span></span></th>
    <th><span data-ttu-id="c77e2-291">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="c77e2-291">API requirement sets</span></span></th>
    <th><span data-ttu-id="c77e2-292"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c77e2-292"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-293">Office Online</span><span class="sxs-lookup"><span data-stu-id="c77e2-293">Office Online</span></span></td>
    <td> <span data-ttu-id="c77e2-294">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="c77e2-294">- Mail Read</span></span><br><span data-ttu-id="c77e2-295">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="c77e2-295">
      - Mail Compose</span></span><br><span data-ttu-id="c77e2-296">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-296">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c77e2-297">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-297">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c77e2-298">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-298">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c77e2-299">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-299">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c77e2-300">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-300">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c77e2-301">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-301">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c77e2-302">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-302">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c77e2-303">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-303">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="c77e2-304">使用不可</span><span class="sxs-lookup"><span data-stu-id="c77e2-304">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-305">Windows 用 Office 2013</span><span class="sxs-lookup"><span data-stu-id="c77e2-305">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="c77e2-306">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="c77e2-306">- Mail Read</span></span><br><span data-ttu-id="c77e2-307">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="c77e2-307">
      - Mail Compose</span></span><br><span data-ttu-id="c77e2-308">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-308">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c77e2-309">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-309">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c77e2-310">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-310">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c77e2-311">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-311">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c77e2-312">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-312">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="c77e2-313">使用不可</span><span class="sxs-lookup"><span data-stu-id="c77e2-313">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-314">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c77e2-314">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="c77e2-315">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="c77e2-315">- Mail Read</span></span><br><span data-ttu-id="c77e2-316">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="c77e2-316">
      - Mail Compose</span></span><br><span data-ttu-id="c77e2-317">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-317">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c77e2-318">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="c77e2-318">
      - Modules</span></span></td>
    <td> <span data-ttu-id="c77e2-319">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-319">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c77e2-320">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-320">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c77e2-321">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-321">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c77e2-322">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-322">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c77e2-323">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-323">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c77e2-324">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-324">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c77e2-325">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-325">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="c77e2-326">使用不可</span><span class="sxs-lookup"><span data-stu-id="c77e2-326">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-327">Windows 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c77e2-327">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="c77e2-328">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="c77e2-328">- Mail Read</span></span><br><span data-ttu-id="c77e2-329">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="c77e2-329">
      - Mail Compose</span></span><br><span data-ttu-id="c77e2-330">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-330">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c77e2-331">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="c77e2-331">
      - Modules</span></span></td>
    <td> <span data-ttu-id="c77e2-332">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-332">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c77e2-333">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-333">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c77e2-334">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-334">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c77e2-335">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-335">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c77e2-336">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-336">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c77e2-337">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-337">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c77e2-338">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-338">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="c77e2-339">使用不可</span><span class="sxs-lookup"><span data-stu-id="c77e2-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-340">iOS 用 Office</span><span class="sxs-lookup"><span data-stu-id="c77e2-340">Office for iOS</span></span></td>
    <td> <span data-ttu-id="c77e2-341">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="c77e2-341">- Mail Read</span></span><br><span data-ttu-id="c77e2-342">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-342">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c77e2-343">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-343">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c77e2-344">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-344">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c77e2-345">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-345">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c77e2-346">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-346">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c77e2-347">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-347">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="c77e2-348">使用不可</span><span class="sxs-lookup"><span data-stu-id="c77e2-348">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-349">Mac 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c77e2-349">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="c77e2-350">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="c77e2-350">- Mail Read</span></span><br><span data-ttu-id="c77e2-351">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="c77e2-351">
      - Mail Compose</span></span><br><span data-ttu-id="c77e2-352">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-352">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c77e2-353">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-353">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c77e2-354">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-354">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c77e2-355">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-355">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c77e2-356">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-356">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c77e2-357">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-357">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c77e2-358">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-358">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c77e2-359">使用不可</span><span class="sxs-lookup"><span data-stu-id="c77e2-359">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-360">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c77e2-360">Office for Mac</span></span></td>
    <td> <span data-ttu-id="c77e2-361">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="c77e2-361">- Mail Read</span></span><br><span data-ttu-id="c77e2-362">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="c77e2-362">
      - Mail Compose</span></span><br><span data-ttu-id="c77e2-363">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-363">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c77e2-364">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-364">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c77e2-365">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-365">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c77e2-366">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-366">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c77e2-367">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-367">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c77e2-368">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-368">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c77e2-369">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-369">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c77e2-370">使用不可</span><span class="sxs-lookup"><span data-stu-id="c77e2-370">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-371">Android 用 Office</span><span class="sxs-lookup"><span data-stu-id="c77e2-371">Office for Android</span></span></td>
    <td> <span data-ttu-id="c77e2-372">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="c77e2-372">- Mail Read</span></span><br><span data-ttu-id="c77e2-373">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-373">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c77e2-374">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-374">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c77e2-375">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-375">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c77e2-376">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-376">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c77e2-377">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-377">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c77e2-378">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-378">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="c77e2-379">使用不可</span><span class="sxs-lookup"><span data-stu-id="c77e2-379">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="c77e2-380">Word</span><span class="sxs-lookup"><span data-stu-id="c77e2-380">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c77e2-381">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="c77e2-381">Platform</span></span></th>
    <th><span data-ttu-id="c77e2-382">拡張点</span><span class="sxs-lookup"><span data-stu-id="c77e2-382">Extension points</span></span></th>
    <th><span data-ttu-id="c77e2-383">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="c77e2-383">API requirement sets</span></span></th>
    <th><span data-ttu-id="c77e2-384"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c77e2-384"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-385">Office Online</span><span class="sxs-lookup"><span data-stu-id="c77e2-385">Office Online</span></span></td>
    <td> <span data-ttu-id="c77e2-386">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-386">- Taskpane</span></span><br><span data-ttu-id="c77e2-387">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-387">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c77e2-388">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-388">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c77e2-389">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-389">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="c77e2-390">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-390">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="c77e2-391">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-391">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c77e2-392">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-392">-BindingEvents</span></span><br><span data-ttu-id="c77e2-393">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="c77e2-393">
         -</span></span><br><span data-ttu-id="c77e2-394">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-394">
         -DocumentEvents</span></span><br><span data-ttu-id="c77e2-395">
         - File</span><span class="sxs-lookup"><span data-stu-id="c77e2-395">
         - File</span></span><br><span data-ttu-id="c77e2-396">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-396">
         -HtmlCoercion</span></span><br><span data-ttu-id="c77e2-397">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-397">
         -ImageCoercion</span></span><br><span data-ttu-id="c77e2-398">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-398">
         -MatrixBindings</span></span><br><span data-ttu-id="c77e2-399">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-399">
         -MatrixCoercion</span></span><br><span data-ttu-id="c77e2-400">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-400">
         -OoxmlCoercion</span></span><br><span data-ttu-id="c77e2-401">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-401">
         -PdfFile</span></span><br><span data-ttu-id="c77e2-402">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c77e2-402">
         - Selection</span></span><br><span data-ttu-id="c77e2-403">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-403">
         - Settings</span></span><br><span data-ttu-id="c77e2-404">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-404">
         -TableBindings</span></span><br><span data-ttu-id="c77e2-405">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-405">
         -TableCoercion</span></span><br><span data-ttu-id="c77e2-406">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-406">
         -TextBindings</span></span><br><span data-ttu-id="c77e2-407">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-407">
         -TextCoercion</span></span><br><span data-ttu-id="c77e2-408">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-408">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-409">Windows 用 Office 2013</span><span class="sxs-lookup"><span data-stu-id="c77e2-409">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="c77e2-410">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-410">- Taskpane</span></span></td>
    <td> <span data-ttu-id="c77e2-411">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-411">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c77e2-412">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-412">-BindingEvents</span></span><br><span data-ttu-id="c77e2-413">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-413">
         -CompressedFile</span></span><br><span data-ttu-id="c77e2-414">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="c77e2-414">
         -</span></span><br><span data-ttu-id="c77e2-415">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-415">
         -DocumentEvents</span></span><br><span data-ttu-id="c77e2-416">
         - File</span><span class="sxs-lookup"><span data-stu-id="c77e2-416">
         - File</span></span><br><span data-ttu-id="c77e2-417">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-417">
         -HtmlCoercion</span></span><br><span data-ttu-id="c77e2-418">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-418">
         -ImageCoercion</span></span><br><span data-ttu-id="c77e2-419">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-419">
         -MatrixBindings</span></span><br><span data-ttu-id="c77e2-420">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-420">
         -MatrixCoercion</span></span><br><span data-ttu-id="c77e2-421">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-421">
         -OoxmlCoercion</span></span><br><span data-ttu-id="c77e2-422">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-422">
         -PdfFile</span></span><br><span data-ttu-id="c77e2-423">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c77e2-423">
         - Selection</span></span><br><span data-ttu-id="c77e2-424">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-424">
         - Settings</span></span><br><span data-ttu-id="c77e2-425">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-425">
         -TableBindings</span></span><br><span data-ttu-id="c77e2-426">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-426">
         -TableCoercion</span></span><br><span data-ttu-id="c77e2-427">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-427">
         -TextBindings</span></span><br><span data-ttu-id="c77e2-428">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-428">
         -TextCoercion</span></span><br><span data-ttu-id="c77e2-429">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-429">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-430">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c77e2-430">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="c77e2-431">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-431">- Taskpane</span></span><br><span data-ttu-id="c77e2-432">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-432">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c77e2-433">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-433">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c77e2-434">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-434">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="c77e2-435">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-435">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="c77e2-436">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-436">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c77e2-437">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-437">-BindingEvents</span></span><br><span data-ttu-id="c77e2-438">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-438">
         -CompressedFile</span></span><br><span data-ttu-id="c77e2-439">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="c77e2-439">
         -</span></span><br><span data-ttu-id="c77e2-440">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-440">
         -DocumentEvents</span></span><br><span data-ttu-id="c77e2-441">
         - File</span><span class="sxs-lookup"><span data-stu-id="c77e2-441">
         - File</span></span><br><span data-ttu-id="c77e2-442">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-442">
         -HtmlCoercion</span></span><br><span data-ttu-id="c77e2-443">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-443">
         -ImageCoercion</span></span><br><span data-ttu-id="c77e2-444">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-444">
         -MatrixBindings</span></span><br><span data-ttu-id="c77e2-445">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-445">
         -MatrixCoercion</span></span><br><span data-ttu-id="c77e2-446">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-446">
         -OoxmlCoercion</span></span><br><span data-ttu-id="c77e2-447">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-447">
         -PdfFile</span></span><br><span data-ttu-id="c77e2-448">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c77e2-448">
         - Selection</span></span><br><span data-ttu-id="c77e2-449">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-449">
         - Settings</span></span><br><span data-ttu-id="c77e2-450">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-450">
         -TableBindings</span></span><br><span data-ttu-id="c77e2-451">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-451">
         -TableCoercion</span></span><br><span data-ttu-id="c77e2-452">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-452">
         -TextBindings</span></span><br><span data-ttu-id="c77e2-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-453">
         -TextCoercion</span></span><br><span data-ttu-id="c77e2-454">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-454">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-455">Windows 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c77e2-455">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="c77e2-456">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-456">- Taskpane</span></span><br><span data-ttu-id="c77e2-457">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-457">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c77e2-458">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-458">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c77e2-459">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-459">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="c77e2-460">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-460">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="c77e2-461">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-461">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c77e2-462">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-462">-BindingEvents</span></span><br><span data-ttu-id="c77e2-463">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-463">
         -CompressedFile</span></span><br><span data-ttu-id="c77e2-464">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="c77e2-464">
         -</span></span><br><span data-ttu-id="c77e2-465">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-465">
         -DocumentEvents</span></span><br><span data-ttu-id="c77e2-466">
         - File</span><span class="sxs-lookup"><span data-stu-id="c77e2-466">
         - File</span></span><br><span data-ttu-id="c77e2-467">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-467">
         -HtmlCoercion</span></span><br><span data-ttu-id="c77e2-468">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-468">
         -ImageCoercion</span></span><br><span data-ttu-id="c77e2-469">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-469">
         -MatrixBindings</span></span><br><span data-ttu-id="c77e2-470">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-470">
         -MatrixCoercion</span></span><br><span data-ttu-id="c77e2-471">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-471">
         -OoxmlCoercion</span></span><br><span data-ttu-id="c77e2-472">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-472">
         -PdfFile</span></span><br><span data-ttu-id="c77e2-473">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c77e2-473">
         - Selection</span></span><br><span data-ttu-id="c77e2-474">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-474">
         - Settings</span></span><br><span data-ttu-id="c77e2-475">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-475">
         -TableBindings</span></span><br><span data-ttu-id="c77e2-476">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-476">
         -TableCoercion</span></span><br><span data-ttu-id="c77e2-477">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-477">
         -TextBindings</span></span><br><span data-ttu-id="c77e2-478">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-478">
         -TextCoercion</span></span><br><span data-ttu-id="c77e2-479">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-479">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-480">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="c77e2-480">Office for iOS</span></span></td>
    <td> <span data-ttu-id="c77e2-481">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-481">- Taskpane</span></span></td>
    <td> <span data-ttu-id="c77e2-482">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-482">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c77e2-483">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-483">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="c77e2-484">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-484">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="c77e2-485">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="c77e2-485">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="c77e2-486">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-486">-BindingEvents</span></span><br><span data-ttu-id="c77e2-487">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-487">
         -CompressedFile</span></span><br><span data-ttu-id="c77e2-488">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="c77e2-488">
         -</span></span><br><span data-ttu-id="c77e2-489">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-489">
         -DocumentEvents</span></span><br><span data-ttu-id="c77e2-490">
         - File</span><span class="sxs-lookup"><span data-stu-id="c77e2-490">
         - File</span></span><br><span data-ttu-id="c77e2-491">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-491">
         -HtmlCoercion</span></span><br><span data-ttu-id="c77e2-492">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-492">
         -ImageCoercion</span></span><br><span data-ttu-id="c77e2-493">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-493">
         -MatrixBindings</span></span><br><span data-ttu-id="c77e2-494">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-494">
         -MatrixCoercion</span></span><br><span data-ttu-id="c77e2-495">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-495">
         -OoxmlCoercion</span></span><br><span data-ttu-id="c77e2-496">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-496">
         -PdfFile</span></span><br><span data-ttu-id="c77e2-497">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c77e2-497">
         - Selection</span></span><br><span data-ttu-id="c77e2-498">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-498">
         - Settings</span></span><br><span data-ttu-id="c77e2-499">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-499">
         -TableBindings</span></span><br><span data-ttu-id="c77e2-500">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-500">
         -TableCoercion</span></span><br><span data-ttu-id="c77e2-501">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-501">
         -TextBindings</span></span><br><span data-ttu-id="c77e2-502">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-502">
         -TextCoercion</span></span><br><span data-ttu-id="c77e2-503">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-503">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-504">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c77e2-504">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="c77e2-505">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-505">- Taskpane</span></span><br><span data-ttu-id="c77e2-506">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-506">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c77e2-507">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-507">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c77e2-508">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-508">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="c77e2-509">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-509">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="c77e2-510">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="c77e2-510">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="c77e2-511">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-511">-BindingEvents</span></span><br><span data-ttu-id="c77e2-512">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-512">
         -CompressedFile</span></span><br><span data-ttu-id="c77e2-513">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="c77e2-513">
         -</span></span><br><span data-ttu-id="c77e2-514">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-514">
         -DocumentEvents</span></span><br><span data-ttu-id="c77e2-515">
         - File</span><span class="sxs-lookup"><span data-stu-id="c77e2-515">
         - File</span></span><br><span data-ttu-id="c77e2-516">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-516">
         -HtmlCoercion</span></span><br><span data-ttu-id="c77e2-517">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-517">
         -ImageCoercion</span></span><br><span data-ttu-id="c77e2-518">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-518">
         -MatrixBindings</span></span><br><span data-ttu-id="c77e2-519">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-519">
         -MatrixCoercion</span></span><br><span data-ttu-id="c77e2-520">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-520">
         -OoxmlCoercion</span></span><br><span data-ttu-id="c77e2-521">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-521">
         -PdfFile</span></span><br><span data-ttu-id="c77e2-522">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c77e2-522">
         - Selection</span></span><br><span data-ttu-id="c77e2-523">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-523">
         - Settings</span></span><br><span data-ttu-id="c77e2-524">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-524">
         -TableBindings</span></span><br><span data-ttu-id="c77e2-525">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-525">
         -TableCoercion</span></span><br><span data-ttu-id="c77e2-526">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-526">
         -TextBindings</span></span><br><span data-ttu-id="c77e2-527">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-527">
         -TextCoercion</span></span><br><span data-ttu-id="c77e2-528">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-528">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-529">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c77e2-529">Office for Mac</span></span></td>
    <td> <span data-ttu-id="c77e2-530">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-530">- Taskpane</span></span><br><span data-ttu-id="c77e2-531">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-531">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c77e2-532">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-532">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="c77e2-533">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-533">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="c77e2-534">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-534">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="c77e2-535">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="c77e2-535">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="c77e2-536">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-536">-BindingEvents</span></span><br><span data-ttu-id="c77e2-537">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-537">
         -CompressedFile</span></span><br><span data-ttu-id="c77e2-538">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="c77e2-538">
         -</span></span><br><span data-ttu-id="c77e2-539">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-539">
         -DocumentEvents</span></span><br><span data-ttu-id="c77e2-540">
         - File</span><span class="sxs-lookup"><span data-stu-id="c77e2-540">
         - File</span></span><br><span data-ttu-id="c77e2-541">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-541">
         -HtmlCoercion</span></span><br><span data-ttu-id="c77e2-542">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-542">
         -ImageCoercion</span></span><br><span data-ttu-id="c77e2-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-543">
         -MatrixBindings</span></span><br><span data-ttu-id="c77e2-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-544">
         -MatrixCoercion</span></span><br><span data-ttu-id="c77e2-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-545">
         -OoxmlCoercion</span></span><br><span data-ttu-id="c77e2-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-546">
         -PdfFile</span></span><br><span data-ttu-id="c77e2-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c77e2-547">
         - Selection</span></span><br><span data-ttu-id="c77e2-548">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-548">
         - Settings</span></span><br><span data-ttu-id="c77e2-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-549">
         -TableBindings</span></span><br><span data-ttu-id="c77e2-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-550">
         -TableCoercion</span></span><br><span data-ttu-id="c77e2-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c77e2-551">
         -TextBindings</span></span><br><span data-ttu-id="c77e2-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-552">
         -TextCoercion</span></span><br><span data-ttu-id="c77e2-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-553">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="c77e2-554">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c77e2-554">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c77e2-555">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="c77e2-555">Platform</span></span></th>
    <th><span data-ttu-id="c77e2-556">拡張点</span><span class="sxs-lookup"><span data-stu-id="c77e2-556">Extension points</span></span></th>
    <th><span data-ttu-id="c77e2-557">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="c77e2-557">API requirement sets</span></span></th>
    <th><span data-ttu-id="c77e2-558"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c77e2-558"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-559">Office Online</span><span class="sxs-lookup"><span data-stu-id="c77e2-559">Office Online</span></span></td>
    <td> <span data-ttu-id="c77e2-560">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c77e2-560">- Content</span></span><br><span data-ttu-id="c77e2-561">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-561">
         - Taskpane</span></span><br><span data-ttu-id="c77e2-562">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-562">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c77e2-563">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-563">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c77e2-564">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c77e2-564">-ActiveView</span></span><br><span data-ttu-id="c77e2-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-565">
         -CompressedFile</span></span><br><span data-ttu-id="c77e2-566">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-566">
         -DocumentEvents</span></span><br><span data-ttu-id="c77e2-567">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="c77e2-567">
         - File</span></span><br><span data-ttu-id="c77e2-568">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-568">
         -ImageCoercion</span></span><br><span data-ttu-id="c77e2-569">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-569">
         -PdfFile</span></span><br><span data-ttu-id="c77e2-570">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c77e2-570">
         - Selection</span></span><br><span data-ttu-id="c77e2-571">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-571">
         - Settings</span></span><br><span data-ttu-id="c77e2-572">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-572">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-573">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="c77e2-573">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="c77e2-574">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c77e2-574">- Content</span></span><br><span data-ttu-id="c77e2-575">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-575">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="c77e2-576">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="c77e2-576">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="c77e2-577">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c77e2-577">-ActiveView</span></span><br><span data-ttu-id="c77e2-578">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-578">
         -CompressedFile</span></span><br><span data-ttu-id="c77e2-579">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-579">
         -DocumentEvents</span></span><br><span data-ttu-id="c77e2-580">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="c77e2-580">
         - File</span></span><br><span data-ttu-id="c77e2-581">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-581">
         -ImageCoercion</span></span><br><span data-ttu-id="c77e2-582">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-582">
         -PdfFile</span></span><br><span data-ttu-id="c77e2-583">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c77e2-583">
         - Selection</span></span><br><span data-ttu-id="c77e2-584">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-584">
         - Settings</span></span><br><span data-ttu-id="c77e2-585">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-585">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-586">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c77e2-586">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="c77e2-587">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c77e2-587">- Content</span></span><br><span data-ttu-id="c77e2-588">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-588">
         - Taskpane</span></span><br><span data-ttu-id="c77e2-589">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-589">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c77e2-590">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-590">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c77e2-591">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c77e2-591">-ActiveView</span></span><br><span data-ttu-id="c77e2-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-592">
         -CompressedFile</span></span><br><span data-ttu-id="c77e2-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-593">
         -DocumentEvents</span></span><br><span data-ttu-id="c77e2-594">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="c77e2-594">
         - File</span></span><br><span data-ttu-id="c77e2-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-595">
         -ImageCoercion</span></span><br><span data-ttu-id="c77e2-596">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-596">
         -PdfFile</span></span><br><span data-ttu-id="c77e2-597">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c77e2-597">
         - Selection</span></span><br><span data-ttu-id="c77e2-598">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-598">
         - Settings</span></span><br><span data-ttu-id="c77e2-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-599">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-600">Windows 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c77e2-600">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="c77e2-601">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c77e2-601">- Content</span></span><br><span data-ttu-id="c77e2-602">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-602">
         - Taskpane</span></span><br><span data-ttu-id="c77e2-603">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-603">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c77e2-604">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-604">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c77e2-605">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c77e2-605">-ActiveView</span></span><br><span data-ttu-id="c77e2-606">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-606">
         -CompressedFile</span></span><br><span data-ttu-id="c77e2-607">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-607">
         -DocumentEvents</span></span><br><span data-ttu-id="c77e2-608">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="c77e2-608">
         - File</span></span><br><span data-ttu-id="c77e2-609">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-609">
         -ImageCoercion</span></span><br><span data-ttu-id="c77e2-610">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-610">
         -PdfFile</span></span><br><span data-ttu-id="c77e2-611">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c77e2-611">
         - Selection</span></span><br><span data-ttu-id="c77e2-612">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-612">
         - Settings</span></span><br><span data-ttu-id="c77e2-613">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-613">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-614">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="c77e2-614">Office for iOS</span></span></td>
    <td> <span data-ttu-id="c77e2-615">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c77e2-615">- Content</span></span><br><span data-ttu-id="c77e2-616">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-616">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="c77e2-617">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-617">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="c77e2-618">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c77e2-618">-ActiveView</span></span><br><span data-ttu-id="c77e2-619">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-619">
         -CompressedFile</span></span><br><span data-ttu-id="c77e2-620">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-620">
         -DocumentEvents</span></span><br><span data-ttu-id="c77e2-621">
         - File</span><span class="sxs-lookup"><span data-stu-id="c77e2-621">
         - File</span></span><br><span data-ttu-id="c77e2-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-622">
         -PdfFile</span></span><br><span data-ttu-id="c77e2-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c77e2-623">
         - Selection</span></span><br><span data-ttu-id="c77e2-624">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-624">
         - Settings</span></span><br><span data-ttu-id="c77e2-625">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-625">
         -TextCoercion</span></span><br><span data-ttu-id="c77e2-626">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-626">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-627">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c77e2-627">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="c77e2-628">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c77e2-628">- Content</span></span><br><span data-ttu-id="c77e2-629">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-629">
         - Taskpane</span></span><br><span data-ttu-id="c77e2-630">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-630">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c77e2-631">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-631">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c77e2-632">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c77e2-632">-ActiveView</span></span><br><span data-ttu-id="c77e2-633">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-633">
         -CompressedFile</span></span><br><span data-ttu-id="c77e2-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-634">
         -DocumentEvents</span></span><br><span data-ttu-id="c77e2-635">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="c77e2-635">
         - File</span></span><br><span data-ttu-id="c77e2-636">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-636">
         -ImageCoercion</span></span><br><span data-ttu-id="c77e2-637">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-637">
         -PdfFile</span></span><br><span data-ttu-id="c77e2-638">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c77e2-638">
         - Selection</span></span><br><span data-ttu-id="c77e2-639">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-639">
         - Settings</span></span><br><span data-ttu-id="c77e2-640">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-640">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-641">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c77e2-641">Office for Mac</span></span></td>
    <td> <span data-ttu-id="c77e2-642">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c77e2-642">- Content</span></span><br><span data-ttu-id="c77e2-643">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-643">
         - Taskpane</span></span><br><span data-ttu-id="c77e2-644">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-644">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c77e2-645">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-645">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c77e2-646">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c77e2-646">-ActiveView</span></span><br><span data-ttu-id="c77e2-647">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-647">
         -CompressedFile</span></span><br><span data-ttu-id="c77e2-648">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-648">
         -DocumentEvents</span></span><br><span data-ttu-id="c77e2-649">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="c77e2-649">
         - File</span></span><br><span data-ttu-id="c77e2-650">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-650">
         -ImageCoercion</span></span><br><span data-ttu-id="c77e2-651">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c77e2-651">
         -PdfFile</span></span><br><span data-ttu-id="c77e2-652">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c77e2-652">
         - Selection</span></span><br><span data-ttu-id="c77e2-653">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-653">
         - Settings</span></span><br><span data-ttu-id="c77e2-654">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-654">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="c77e2-655">OneNote</span><span class="sxs-lookup"><span data-stu-id="c77e2-655">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c77e2-656">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="c77e2-656">Platform</span></span></th>
    <th><span data-ttu-id="c77e2-657">拡張点</span><span class="sxs-lookup"><span data-stu-id="c77e2-657">Extension points</span></span></th>
    <th><span data-ttu-id="c77e2-658">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="c77e2-658">API requirement sets</span></span></th>
    <th><span data-ttu-id="c77e2-659"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c77e2-659"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="c77e2-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="c77e2-660">Office Online</span></span></td>
    <td> <span data-ttu-id="c77e2-661">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c77e2-661">- Content</span></span><br><span data-ttu-id="c77e2-662">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c77e2-662">
         - Taskpane</span></span><br><span data-ttu-id="c77e2-663">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-663">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c77e2-664">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-664">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="c77e2-665">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c77e2-665">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c77e2-666">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c77e2-666">-DocumentEvents</span></span><br><span data-ttu-id="c77e2-667">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-667">
         -HtmlCoercion</span></span><br><span data-ttu-id="c77e2-668">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-668">
         -ImageCoercion</span></span><br><span data-ttu-id="c77e2-669">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="c77e2-669">
         - Settings</span></span><br><span data-ttu-id="c77e2-670">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c77e2-670">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="c77e2-671">関連項目</span><span class="sxs-lookup"><span data-stu-id="c77e2-671">See also</span></span>

- [<span data-ttu-id="c77e2-672">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="c77e2-672">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="c77e2-673">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="c77e2-673">Common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="c77e2-674">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="c77e2-674">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="c77e2-675">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="c77e2-675">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office)
