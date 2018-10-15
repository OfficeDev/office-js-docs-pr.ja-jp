---
title: Office アドインのホストとプラットフォームの可用性
description: Excel、Word、Outlook、PowerPoint、および OneNote のサポートされる要件セット。
ms.date: 10/03/2018
ms.openlocfilehash: bc7ac5c97c041a546c160c05cffc2c80db1ff1b1
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506351"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="5934c-103">Office アドインのホストとプラットフォームの可用性</span><span class="sxs-lookup"><span data-stu-id="5934c-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="5934c-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API の要件セットが含まれています。</span><span class="sxs-lookup"><span data-stu-id="5934c-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="5934c-p102">表のセルにアスタリスク ( \* ) が含まれる場合は、準備中を意味します。Project または Access の要件セットについては、「[Office の共有要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5934c-p102">If a table cell contains an asterisk ( \* ), that means we’re working on it. For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="5934c-p103">MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。このバージョンには、ExcelApi 1.1、WordApi 1.1、および共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="5934c-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="5934c-110">Excel</span><span class="sxs-lookup"><span data-stu-id="5934c-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="5934c-111">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5934c-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="5934c-112">拡張点</span><span class="sxs-lookup"><span data-stu-id="5934c-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="5934c-113">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="5934c-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="5934c-114"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5934c-114"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="5934c-115">Office Online</span></span></td>
    <td> <span data-ttu-id="5934c-116">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-116">- Taskpane</span></span><br><span data-ttu-id="5934c-117">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5934c-117">
        - Content</span></span><br><span data-ttu-id="5934c-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="5934c-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="5934c-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5934c-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5934c-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5934c-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5934c-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5934c-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5934c-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5934c-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5934c-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5934c-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5934c-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5934c-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5934c-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="5934c-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5934c-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5934c-127">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-127">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5934c-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-128">
        -BindingEvents</span></span><br><span data-ttu-id="5934c-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5934c-129">
        -CompressedFile</span></span><br><span data-ttu-id="5934c-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-130">
        -DocumentEvents</span></span><br><span data-ttu-id="5934c-131">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="5934c-131">
        - File</span></span><br><span data-ttu-id="5934c-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-132">
        -MatrixBindings</span></span><br><span data-ttu-id="5934c-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-133">
        -MatrixCoercion</span></span><br><span data-ttu-id="5934c-134">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="5934c-134">
        - Selection</span></span><br><span data-ttu-id="5934c-135">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-135">
        - Settings</span></span><br><span data-ttu-id="5934c-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-136">
        -TableBindings</span></span><br><span data-ttu-id="5934c-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-137">
        -TableCoercion</span></span><br><span data-ttu-id="5934c-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-138">
        -TextBindings</span></span><br><span data-ttu-id="5934c-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-139">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-140">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="5934c-140">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="5934c-141">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-141">
        - Taskpane</span></span><br><span data-ttu-id="5934c-142">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5934c-142">
        - Content</span></span></td>
    <td>  <span data-ttu-id="5934c-143">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-143">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5934c-144">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-144">
        -BindingEvents</span></span><br><span data-ttu-id="5934c-145">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5934c-145">
        -CompressedFile</span></span><br><span data-ttu-id="5934c-146">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-146">
        -DocumentEvents</span></span><br><span data-ttu-id="5934c-147">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="5934c-147">
        - File</span></span><br><span data-ttu-id="5934c-148">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-148">
        -ImageCoercion</span></span><br><span data-ttu-id="5934c-149">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-149">
        -MatrixBindings</span></span><br><span data-ttu-id="5934c-150">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-150">
        -MatrixCoercion</span></span><br><span data-ttu-id="5934c-151">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="5934c-151">
        - Selection</span></span><br><span data-ttu-id="5934c-152">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-152">
        - Settings</span></span><br><span data-ttu-id="5934c-153">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-153">
        -TableBindings</span></span><br><span data-ttu-id="5934c-154">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-154">
        -TableCoercion</span></span><br><span data-ttu-id="5934c-155">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-155">
        -TextBindings</span></span><br><span data-ttu-id="5934c-156">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-156">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-157">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="5934c-157">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="5934c-158">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-158">- Taskpane</span></span><br><span data-ttu-id="5934c-159">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5934c-159">
        - Content</span></span><br><span data-ttu-id="5934c-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5934c-161">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-161">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5934c-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5934c-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5934c-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5934c-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5934c-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5934c-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5934c-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5934c-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5934c-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5934c-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5934c-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5934c-167">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="5934c-168">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5934c-168">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5934c-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5934c-170">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-170">-BindingEvents</span></span><br><span data-ttu-id="5934c-171">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5934c-171">
        -CompressedFile</span></span><br><span data-ttu-id="5934c-172">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-172">
        -DocumentEvents</span></span><br><span data-ttu-id="5934c-173">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="5934c-173">
        - File</span></span><br><span data-ttu-id="5934c-174">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-174">
        -ImageCoercion</span></span><br><span data-ttu-id="5934c-175">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-175">
        -MatrixBindings</span></span><br><span data-ttu-id="5934c-176">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-176">
        -MatrixCoercion</span></span><br><span data-ttu-id="5934c-177">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="5934c-177">
        - Selection</span></span><br><span data-ttu-id="5934c-178">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-178">
        - Settings</span></span><br><span data-ttu-id="5934c-179">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-179">
        -TableBindings</span></span><br><span data-ttu-id="5934c-180">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-180">
        -TableCoercion</span></span><br><span data-ttu-id="5934c-181">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-181">
        -TextBindings</span></span><br><span data-ttu-id="5934c-182">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-182">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-183">Windows 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="5934c-183">Office for Windows Desktop.</span></span></td>
    <td><span data-ttu-id="5934c-184">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-184">- Taskpane</span></span><br><span data-ttu-id="5934c-185">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5934c-185">
        - Content</span></span><br><span data-ttu-id="5934c-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5934c-187">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-187">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5934c-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5934c-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5934c-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5934c-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5934c-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5934c-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5934c-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5934c-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5934c-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5934c-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5934c-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5934c-193">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="5934c-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5934c-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5934c-195">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-195">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5934c-196">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-196">-BindingEvents</span></span><br><span data-ttu-id="5934c-197">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5934c-197">
        -CompressedFile</span></span><br><span data-ttu-id="5934c-198">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-198">
        -DocumentEvents</span></span><br><span data-ttu-id="5934c-199">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="5934c-199">
        - File</span></span><br><span data-ttu-id="5934c-200">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-200">
        -ImageCoercion</span></span><br><span data-ttu-id="5934c-201">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-201">
        -MatrixBindings</span></span><br><span data-ttu-id="5934c-202">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-202">
        -MatrixCoercion</span></span><br><span data-ttu-id="5934c-203">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="5934c-203">
        - Selection</span></span><br><span data-ttu-id="5934c-204">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-204">
        - Settings</span></span><br><span data-ttu-id="5934c-205">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-205">
        -TableBindings</span></span><br><span data-ttu-id="5934c-206">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-206">
        -TableCoercion</span></span><br><span data-ttu-id="5934c-207">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-207">
        -TextBindings</span></span><br><span data-ttu-id="5934c-208">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-208">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-209">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="5934c-209">Office for iOS</span></span></td>
    <td><span data-ttu-id="5934c-210">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-210">- Taskpane</span></span><br><span data-ttu-id="5934c-211">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5934c-211">
        - Content</span></span></td>
    <td><span data-ttu-id="5934c-212">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-212">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5934c-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5934c-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5934c-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5934c-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5934c-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5934c-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5934c-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5934c-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5934c-217">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5934c-217">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5934c-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5934c-218">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="5934c-219">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5934c-219">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5934c-220">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-220">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5934c-221">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-221">-BindingEvents</span></span><br><span data-ttu-id="5934c-222">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5934c-222">
        -CompressedFile</span></span><br><span data-ttu-id="5934c-223">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-223">
        -DocumentEvents</span></span><br><span data-ttu-id="5934c-224">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="5934c-224">
        - File</span></span><br><span data-ttu-id="5934c-225">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-225">
        -ImageCoercion</span></span><br><span data-ttu-id="5934c-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-226">
        -MatrixBindings</span></span><br><span data-ttu-id="5934c-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-227">
        -MatrixCoercion</span></span><br><span data-ttu-id="5934c-228">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="5934c-228">
        - Selection</span></span><br><span data-ttu-id="5934c-229">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-229">
        - Settings</span></span><br><span data-ttu-id="5934c-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-230">
        -TableBindings</span></span><br><span data-ttu-id="5934c-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-231">
        -TableCoercion</span></span><br><span data-ttu-id="5934c-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-232">
        -TextBindings</span></span><br><span data-ttu-id="5934c-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-233">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-234">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="5934c-234">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="5934c-235">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-235">- Taskpane</span></span><br><span data-ttu-id="5934c-236">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5934c-236">
        - Content</span></span><br><span data-ttu-id="5934c-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5934c-238">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-238">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5934c-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5934c-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5934c-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5934c-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5934c-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5934c-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5934c-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5934c-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5934c-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5934c-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5934c-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5934c-244">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="5934c-245">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5934c-245">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5934c-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5934c-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-247">-BindingEvents</span></span><br><span data-ttu-id="5934c-248">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5934c-248">
        -CompressedFile</span></span><br><span data-ttu-id="5934c-249">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-249">
        -DocumentEvents</span></span><br><span data-ttu-id="5934c-250">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="5934c-250">
        - File</span></span><br><span data-ttu-id="5934c-251">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-251">
        -ImageCoercion</span></span><br><span data-ttu-id="5934c-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-252">
        -MatrixBindings</span></span><br><span data-ttu-id="5934c-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-253">
        -MatrixCoercion</span></span><br><span data-ttu-id="5934c-254">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5934c-254">
        -PdfFile</span></span><br><span data-ttu-id="5934c-255">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5934c-255">
        - Selection</span></span><br><span data-ttu-id="5934c-256">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-256">
        - Settings</span></span><br><span data-ttu-id="5934c-257">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-257">
        -TableBindings</span></span><br><span data-ttu-id="5934c-258">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-258">
        -TableCoercion</span></span><br><span data-ttu-id="5934c-259">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-259">
        -TextBindings</span></span><br><span data-ttu-id="5934c-260">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-260">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-261">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="5934c-261">Office for Mac</span></span></td>
    <td><span data-ttu-id="5934c-262">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-262">- Taskpane</span></span><br><span data-ttu-id="5934c-263">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5934c-263">
        - Content</span></span><br><span data-ttu-id="5934c-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5934c-265">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-265">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5934c-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5934c-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5934c-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5934c-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5934c-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5934c-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5934c-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5934c-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5934c-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5934c-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5934c-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5934c-271">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="5934c-272">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5934c-272">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5934c-273">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-273">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5934c-274">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-274">-BindingEvents</span></span><br><span data-ttu-id="5934c-275">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5934c-275">
        -CompressedFile</span></span><br><span data-ttu-id="5934c-276">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-276">
        -DocumentEvents</span></span><br><span data-ttu-id="5934c-277">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="5934c-277">
        - File</span></span><br><span data-ttu-id="5934c-278">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-278">
        -ImageCoercion</span></span><br><span data-ttu-id="5934c-279">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-279">
        -MatrixBindings</span></span><br><span data-ttu-id="5934c-280">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-280">
        -MatrixCoercion</span></span><br><span data-ttu-id="5934c-281">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5934c-281">
        -PdfFile</span></span><br><span data-ttu-id="5934c-282">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5934c-282">
        - Selection</span></span><br><span data-ttu-id="5934c-283">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-283">
        - Settings</span></span><br><span data-ttu-id="5934c-284">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-284">
        -TableBindings</span></span><br><span data-ttu-id="5934c-285">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-285">
        -TableCoercion</span></span><br><span data-ttu-id="5934c-286">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-286">
        -TextBindings</span></span><br><span data-ttu-id="5934c-287">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-287">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="5934c-288">Outlook</span><span class="sxs-lookup"><span data-stu-id="5934c-288">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5934c-289">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5934c-289">Platform</span></span></th>
    <th><span data-ttu-id="5934c-290">拡張点</span><span class="sxs-lookup"><span data-stu-id="5934c-290">Extension points</span></span></th>
    <th><span data-ttu-id="5934c-291">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="5934c-291">API requirement sets</span></span></th>
    <th><span data-ttu-id="5934c-292"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5934c-292"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-293">Office Online</span><span class="sxs-lookup"><span data-stu-id="5934c-293">Office Online</span></span></td>
    <td> <span data-ttu-id="5934c-294">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5934c-294">- Mail Read</span></span><br><span data-ttu-id="5934c-295">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5934c-295">
      - Mail Compose</span></span><br><span data-ttu-id="5934c-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5934c-297">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-297">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5934c-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5934c-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5934c-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5934c-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5934c-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5934c-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5934c-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5934c-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5934c-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5934c-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5934c-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5934c-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="5934c-304">使用不可</span><span class="sxs-lookup"><span data-stu-id="5934c-304">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-305">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="5934c-305">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="5934c-306">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5934c-306">- Mail Read</span></span><br><span data-ttu-id="5934c-307">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5934c-307">
      - Mail Compose</span></span><br><span data-ttu-id="5934c-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5934c-309">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-309">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5934c-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5934c-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5934c-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5934c-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5934c-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5934c-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="5934c-313">使用不可</span><span class="sxs-lookup"><span data-stu-id="5934c-313">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-314">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="5934c-314">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="5934c-315">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5934c-315">- Mail Read</span></span><br><span data-ttu-id="5934c-316">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5934c-316">
      - Mail Compose</span></span><br><span data-ttu-id="5934c-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="5934c-318">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="5934c-318">
      - Modules</span></span></td>
    <td> <span data-ttu-id="5934c-319">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-319">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5934c-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5934c-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5934c-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5934c-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5934c-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5934c-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5934c-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5934c-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5934c-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5934c-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5934c-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5934c-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="5934c-326">使用不可</span><span class="sxs-lookup"><span data-stu-id="5934c-326">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-327">Windows 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="5934c-327">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="5934c-328">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5934c-328">- Mail Read</span></span><br><span data-ttu-id="5934c-329">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5934c-329">
      - Mail Compose</span></span><br><span data-ttu-id="5934c-330">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-330">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="5934c-331">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="5934c-331">
      - Modules</span></span></td>
    <td> <span data-ttu-id="5934c-332">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-332">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5934c-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5934c-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5934c-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5934c-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5934c-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5934c-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5934c-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5934c-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5934c-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5934c-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5934c-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5934c-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="5934c-339">使用不可</span><span class="sxs-lookup"><span data-stu-id="5934c-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-340">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="5934c-340">Office for iOS</span></span></td>
    <td> <span data-ttu-id="5934c-341">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5934c-341">- Mail Read</span></span><br><span data-ttu-id="5934c-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5934c-343">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-343">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5934c-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5934c-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5934c-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5934c-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5934c-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5934c-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5934c-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5934c-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="5934c-348">使用不可</span><span class="sxs-lookup"><span data-stu-id="5934c-348">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-349">Mac 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="5934c-349">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="5934c-350">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5934c-350">- Mail Read</span></span><br><span data-ttu-id="5934c-351">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5934c-351">
      - Mail Compose</span></span><br><span data-ttu-id="5934c-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5934c-353">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-353">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5934c-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5934c-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5934c-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5934c-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5934c-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5934c-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5934c-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5934c-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5934c-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5934c-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="5934c-359">使用不可</span><span class="sxs-lookup"><span data-stu-id="5934c-359">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-360">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="5934c-360">Office for Mac</span></span></td>
    <td> <span data-ttu-id="5934c-361">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5934c-361">- Mail Read</span></span><br><span data-ttu-id="5934c-362">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5934c-362">
      - Mail Compose</span></span><br><span data-ttu-id="5934c-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5934c-364">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-364">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5934c-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5934c-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5934c-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5934c-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5934c-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5934c-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5934c-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5934c-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5934c-369">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5934c-369">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="5934c-370">使用不可</span><span class="sxs-lookup"><span data-stu-id="5934c-370">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-371">Android 用 Office</span><span class="sxs-lookup"><span data-stu-id="5934c-371">Office for Android</span></span></td>
    <td> <span data-ttu-id="5934c-372">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5934c-372">- Mail Read</span></span><br><span data-ttu-id="5934c-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5934c-374">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-374">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5934c-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5934c-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5934c-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5934c-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5934c-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5934c-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5934c-378">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5934c-378">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="5934c-379">使用不可</span><span class="sxs-lookup"><span data-stu-id="5934c-379">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="5934c-380">Word</span><span class="sxs-lookup"><span data-stu-id="5934c-380">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5934c-381">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5934c-381">Platform</span></span></th>
    <th><span data-ttu-id="5934c-382">拡張点</span><span class="sxs-lookup"><span data-stu-id="5934c-382">Extension points</span></span></th>
    <th><span data-ttu-id="5934c-383">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="5934c-383">API requirement sets</span></span></th>
    <th><span data-ttu-id="5934c-384"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5934c-384"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-385">Office Online</span><span class="sxs-lookup"><span data-stu-id="5934c-385">Office Online</span></span></td>
    <td> <span data-ttu-id="5934c-386">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-386">- Taskpane</span></span><br><span data-ttu-id="5934c-387">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-387">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5934c-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5934c-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5934c-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5934c-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5934c-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5934c-391">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-391">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5934c-392">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-392">-BindingEvents</span></span><br><span data-ttu-id="5934c-393">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="5934c-393">
         -CustomXmlParts</span></span><br><span data-ttu-id="5934c-394">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-394">
         -DocumentEvents</span></span><br><span data-ttu-id="5934c-395">
         - File</span><span class="sxs-lookup"><span data-stu-id="5934c-395">
         - File</span></span><br><span data-ttu-id="5934c-396">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-396">
         -HtmlCoercion</span></span><br><span data-ttu-id="5934c-397">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-397">
         -ImageCoercion</span></span><br><span data-ttu-id="5934c-398">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-398">
         -MatrixBindings</span></span><br><span data-ttu-id="5934c-399">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-399">
         -MatrixCoercion</span></span><br><span data-ttu-id="5934c-400">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-400">
         -OoxmlCoercion</span></span><br><span data-ttu-id="5934c-401">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5934c-401">
         -PdfFile</span></span><br><span data-ttu-id="5934c-402">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5934c-402">
         - Selection</span></span><br><span data-ttu-id="5934c-403">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-403">
         - Settings</span></span><br><span data-ttu-id="5934c-404">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-404">
         -TableBindings</span></span><br><span data-ttu-id="5934c-405">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-405">
         -TableCoercion</span></span><br><span data-ttu-id="5934c-406">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-406">
         -TextBindings</span></span><br><span data-ttu-id="5934c-407">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-407">
         -TextCoercion</span></span><br><span data-ttu-id="5934c-408">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5934c-408">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-409">Windows 用 Office 2013</span><span class="sxs-lookup"><span data-stu-id="5934c-409">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="5934c-410">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-410">- Taskpane</span></span></td>
    <td> <span data-ttu-id="5934c-411">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-411">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5934c-412">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-412">-BindingEvents</span></span><br><span data-ttu-id="5934c-413">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5934c-413">
         -CompressedFile</span></span><br><span data-ttu-id="5934c-414">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="5934c-414">
         -CustomXmlParts</span></span><br><span data-ttu-id="5934c-415">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-415">
         -DocumentEvents</span></span><br><span data-ttu-id="5934c-416">
         - File</span><span class="sxs-lookup"><span data-stu-id="5934c-416">
         - File</span></span><br><span data-ttu-id="5934c-417">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-417">
         -HtmlCoercion</span></span><br><span data-ttu-id="5934c-418">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-418">
         -ImageCoercion</span></span><br><span data-ttu-id="5934c-419">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-419">
         -MatrixBindings</span></span><br><span data-ttu-id="5934c-420">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-420">
         -MatrixCoercion</span></span><br><span data-ttu-id="5934c-421">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-421">
         -OoxmlCoercion</span></span><br><span data-ttu-id="5934c-422">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5934c-422">
         -PdfFile</span></span><br><span data-ttu-id="5934c-423">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5934c-423">
         - Selection</span></span><br><span data-ttu-id="5934c-424">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-424">
         - Settings</span></span><br><span data-ttu-id="5934c-425">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-425">
         -TableBindings</span></span><br><span data-ttu-id="5934c-426">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-426">
         -TableCoercion</span></span><br><span data-ttu-id="5934c-427">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-427">
         -TextBindings</span></span><br><span data-ttu-id="5934c-428">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-428">
         -TextCoercion</span></span><br><span data-ttu-id="5934c-429">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5934c-429">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-430">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="5934c-430">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="5934c-431">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-431">- Taskpane</span></span><br><span data-ttu-id="5934c-432">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-432">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5934c-433">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-433">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5934c-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5934c-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5934c-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5934c-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5934c-436">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-436">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5934c-437">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-437">-BindingEvents</span></span><br><span data-ttu-id="5934c-438">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5934c-438">
         -CompressedFile</span></span><br><span data-ttu-id="5934c-439">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="5934c-439">
         -CustomXmlParts</span></span><br><span data-ttu-id="5934c-440">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-440">
         -DocumentEvents</span></span><br><span data-ttu-id="5934c-441">
         - File</span><span class="sxs-lookup"><span data-stu-id="5934c-441">
         - File</span></span><br><span data-ttu-id="5934c-442">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-442">
         -HtmlCoercion</span></span><br><span data-ttu-id="5934c-443">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-443">
         -ImageCoercion</span></span><br><span data-ttu-id="5934c-444">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-444">
         -MatrixBindings</span></span><br><span data-ttu-id="5934c-445">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-445">
         -MatrixCoercion</span></span><br><span data-ttu-id="5934c-446">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-446">
         -OoxmlCoercion</span></span><br><span data-ttu-id="5934c-447">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5934c-447">
         -PdfFile</span></span><br><span data-ttu-id="5934c-448">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5934c-448">
         - Selection</span></span><br><span data-ttu-id="5934c-449">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-449">
         - Settings</span></span><br><span data-ttu-id="5934c-450">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-450">
         -TableBindings</span></span><br><span data-ttu-id="5934c-451">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-451">
         -TableCoercion</span></span><br><span data-ttu-id="5934c-452">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-452">
         -TextBindings</span></span><br><span data-ttu-id="5934c-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-453">
         -TextCoercion</span></span><br><span data-ttu-id="5934c-454">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5934c-454">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-455">Windows 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="5934c-455">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="5934c-456">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-456">- Taskpane</span></span><br><span data-ttu-id="5934c-457">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-457">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5934c-458">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-458">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5934c-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5934c-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5934c-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5934c-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5934c-461">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-461">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5934c-462">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-462">-BindingEvents</span></span><br><span data-ttu-id="5934c-463">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5934c-463">
         -CompressedFile</span></span><br><span data-ttu-id="5934c-464">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="5934c-464">
         -CustomXmlParts</span></span><br><span data-ttu-id="5934c-465">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-465">
         -DocumentEvents</span></span><br><span data-ttu-id="5934c-466">
         - File</span><span class="sxs-lookup"><span data-stu-id="5934c-466">
         - File</span></span><br><span data-ttu-id="5934c-467">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-467">
         -HtmlCoercion</span></span><br><span data-ttu-id="5934c-468">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-468">
         -ImageCoercion</span></span><br><span data-ttu-id="5934c-469">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-469">
         -MatrixBindings</span></span><br><span data-ttu-id="5934c-470">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-470">
         -MatrixCoercion</span></span><br><span data-ttu-id="5934c-471">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-471">
         -OoxmlCoercion</span></span><br><span data-ttu-id="5934c-472">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5934c-472">
         -PdfFile</span></span><br><span data-ttu-id="5934c-473">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5934c-473">
         - Selection</span></span><br><span data-ttu-id="5934c-474">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-474">
         - Settings</span></span><br><span data-ttu-id="5934c-475">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-475">
         -TableBindings</span></span><br><span data-ttu-id="5934c-476">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-476">
         -TableCoercion</span></span><br><span data-ttu-id="5934c-477">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-477">
         -TextBindings</span></span><br><span data-ttu-id="5934c-478">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-478">
         -TextCoercion</span></span><br><span data-ttu-id="5934c-479">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5934c-479">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-480">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="5934c-480">Office for iOS</span></span></td>
    <td> <span data-ttu-id="5934c-481">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-481">- Taskpane</span></span></td>
    <td> <span data-ttu-id="5934c-482">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-482">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5934c-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5934c-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5934c-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5934c-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5934c-485">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="5934c-485">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="5934c-486">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-486">-BindingEvents</span></span><br><span data-ttu-id="5934c-487">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5934c-487">
         -CompressedFile</span></span><br><span data-ttu-id="5934c-488">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="5934c-488">
         -CustomXmlParts</span></span><br><span data-ttu-id="5934c-489">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-489">
         -DocumentEvents</span></span><br><span data-ttu-id="5934c-490">
         - File</span><span class="sxs-lookup"><span data-stu-id="5934c-490">
         - File</span></span><br><span data-ttu-id="5934c-491">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-491">
         -HtmlCoercion</span></span><br><span data-ttu-id="5934c-492">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-492">
         -ImageCoercion</span></span><br><span data-ttu-id="5934c-493">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-493">
         -MatrixBindings</span></span><br><span data-ttu-id="5934c-494">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-494">
         -MatrixCoercion</span></span><br><span data-ttu-id="5934c-495">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-495">
         -OoxmlCoercion</span></span><br><span data-ttu-id="5934c-496">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5934c-496">
         -PdfFile</span></span><br><span data-ttu-id="5934c-497">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5934c-497">
         - Selection</span></span><br><span data-ttu-id="5934c-498">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-498">
         - Settings</span></span><br><span data-ttu-id="5934c-499">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-499">
         -TableBindings</span></span><br><span data-ttu-id="5934c-500">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-500">
         -TableCoercion</span></span><br><span data-ttu-id="5934c-501">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-501">
         -TextBindings</span></span><br><span data-ttu-id="5934c-502">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-502">
         -TextCoercion</span></span><br><span data-ttu-id="5934c-503">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5934c-503">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-504">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="5934c-504">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="5934c-505">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-505">- Taskpane</span></span><br><span data-ttu-id="5934c-506">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-506">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5934c-507">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-507">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5934c-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5934c-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5934c-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5934c-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5934c-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="5934c-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="5934c-511">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-511">-BindingEvents</span></span><br><span data-ttu-id="5934c-512">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5934c-512">
         -CompressedFile</span></span><br><span data-ttu-id="5934c-513">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="5934c-513">
         -CustomXmlParts</span></span><br><span data-ttu-id="5934c-514">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-514">
         -DocumentEvents</span></span><br><span data-ttu-id="5934c-515">
         - File</span><span class="sxs-lookup"><span data-stu-id="5934c-515">
         - File</span></span><br><span data-ttu-id="5934c-516">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-516">
         -HtmlCoercion</span></span><br><span data-ttu-id="5934c-517">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-517">
         -ImageCoercion</span></span><br><span data-ttu-id="5934c-518">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-518">
         -MatrixBindings</span></span><br><span data-ttu-id="5934c-519">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-519">
         -MatrixCoercion</span></span><br><span data-ttu-id="5934c-520">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-520">
         -OoxmlCoercion</span></span><br><span data-ttu-id="5934c-521">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5934c-521">
         -PdfFile</span></span><br><span data-ttu-id="5934c-522">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5934c-522">
         - Selection</span></span><br><span data-ttu-id="5934c-523">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-523">
         - Settings</span></span><br><span data-ttu-id="5934c-524">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-524">
         -TableBindings</span></span><br><span data-ttu-id="5934c-525">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-525">
         -TableCoercion</span></span><br><span data-ttu-id="5934c-526">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-526">
         -TextBindings</span></span><br><span data-ttu-id="5934c-527">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-527">
         -TextCoercion</span></span><br><span data-ttu-id="5934c-528">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5934c-528">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-529">Mac 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="5934c-529">Office for Mac</span></span></td>
    <td> <span data-ttu-id="5934c-530">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-530">- Taskpane</span></span><br><span data-ttu-id="5934c-531">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-531">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5934c-532">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-532">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5934c-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5934c-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5934c-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5934c-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5934c-535">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="5934c-535">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="5934c-536">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-536">-BindingEvents</span></span><br><span data-ttu-id="5934c-537">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5934c-537">
         -CompressedFile</span></span><br><span data-ttu-id="5934c-538">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="5934c-538">
         -CustomXmlParts</span></span><br><span data-ttu-id="5934c-539">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-539">
         -DocumentEvents</span></span><br><span data-ttu-id="5934c-540">
         - File</span><span class="sxs-lookup"><span data-stu-id="5934c-540">
         - File</span></span><br><span data-ttu-id="5934c-541">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-541">
         -HtmlCoercion</span></span><br><span data-ttu-id="5934c-542">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-542">
         -ImageCoercion</span></span><br><span data-ttu-id="5934c-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-543">
         -MatrixBindings</span></span><br><span data-ttu-id="5934c-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-544">
         -MatrixCoercion</span></span><br><span data-ttu-id="5934c-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-545">
         -OoxmlCoercion</span></span><br><span data-ttu-id="5934c-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5934c-546">
         -PdfFile</span></span><br><span data-ttu-id="5934c-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5934c-547">
         - Selection</span></span><br><span data-ttu-id="5934c-548">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-548">
         - Settings</span></span><br><span data-ttu-id="5934c-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-549">
         -TableBindings</span></span><br><span data-ttu-id="5934c-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-550">
         -TableCoercion</span></span><br><span data-ttu-id="5934c-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5934c-551">
         -TextBindings</span></span><br><span data-ttu-id="5934c-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-552">
         -TextCoercion</span></span><br><span data-ttu-id="5934c-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5934c-553">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="5934c-554">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="5934c-554">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5934c-555">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5934c-555">Platform</span></span></th>
    <th><span data-ttu-id="5934c-556">拡張点</span><span class="sxs-lookup"><span data-stu-id="5934c-556">Extension points</span></span></th>
    <th><span data-ttu-id="5934c-557">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="5934c-557">API requirement sets</span></span></th>
    <th><span data-ttu-id="5934c-558"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5934c-558"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-559">Office Online</span><span class="sxs-lookup"><span data-stu-id="5934c-559">Office Online</span></span></td>
    <td> <span data-ttu-id="5934c-560">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5934c-560">- Content</span></span><br><span data-ttu-id="5934c-561">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-561">
         - Taskpane</span></span><br><span data-ttu-id="5934c-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5934c-563">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-563">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5934c-564">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5934c-564">-ActiveView</span></span><br><span data-ttu-id="5934c-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5934c-565">
         -CompressedFile</span></span><br><span data-ttu-id="5934c-566">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-566">
         -DocumentEvents</span></span><br><span data-ttu-id="5934c-567">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="5934c-567">
         - File</span></span><br><span data-ttu-id="5934c-568">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-568">
         -ImageCoercion</span></span><br><span data-ttu-id="5934c-569">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5934c-569">
         -PdfFile</span></span><br><span data-ttu-id="5934c-570">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5934c-570">
         - Selection</span></span><br><span data-ttu-id="5934c-571">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-571">
         - Settings</span></span><br><span data-ttu-id="5934c-572">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-572">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-573">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="5934c-573">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="5934c-574">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5934c-574">- Content</span></span><br><span data-ttu-id="5934c-575">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-575">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="5934c-576">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="5934c-576">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="5934c-577">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5934c-577">-ActiveView</span></span><br><span data-ttu-id="5934c-578">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5934c-578">
         -CompressedFile</span></span><br><span data-ttu-id="5934c-579">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-579">
         -DocumentEvents</span></span><br><span data-ttu-id="5934c-580">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="5934c-580">
         - File</span></span><br><span data-ttu-id="5934c-581">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-581">
         -ImageCoercion</span></span><br><span data-ttu-id="5934c-582">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5934c-582">
         -PdfFile</span></span><br><span data-ttu-id="5934c-583">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5934c-583">
         - Selection</span></span><br><span data-ttu-id="5934c-584">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-584">
         - Settings</span></span><br><span data-ttu-id="5934c-585">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-585">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-586">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="5934c-586">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="5934c-587">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5934c-587">- Content</span></span><br><span data-ttu-id="5934c-588">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-588">
         - Taskpane</span></span><br><span data-ttu-id="5934c-589">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-589">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5934c-590">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-590">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5934c-591">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5934c-591">-ActiveView</span></span><br><span data-ttu-id="5934c-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5934c-592">
         -CompressedFile</span></span><br><span data-ttu-id="5934c-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-593">
         -DocumentEvents</span></span><br><span data-ttu-id="5934c-594">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="5934c-594">
         - File</span></span><br><span data-ttu-id="5934c-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-595">
         -ImageCoercion</span></span><br><span data-ttu-id="5934c-596">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5934c-596">
         -PdfFile</span></span><br><span data-ttu-id="5934c-597">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5934c-597">
         - Selection</span></span><br><span data-ttu-id="5934c-598">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-598">
         - Settings</span></span><br><span data-ttu-id="5934c-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-599">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-600">Windows 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="5934c-600">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="5934c-601">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5934c-601">- Content</span></span><br><span data-ttu-id="5934c-602">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-602">
         - Taskpane</span></span><br><span data-ttu-id="5934c-603">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-603">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5934c-604">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-604">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5934c-605">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5934c-605">-ActiveView</span></span><br><span data-ttu-id="5934c-606">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5934c-606">
         -CompressedFile</span></span><br><span data-ttu-id="5934c-607">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-607">
         -DocumentEvents</span></span><br><span data-ttu-id="5934c-608">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="5934c-608">
         - File</span></span><br><span data-ttu-id="5934c-609">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-609">
         -ImageCoercion</span></span><br><span data-ttu-id="5934c-610">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5934c-610">
         -PdfFile</span></span><br><span data-ttu-id="5934c-611">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5934c-611">
         - Selection</span></span><br><span data-ttu-id="5934c-612">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-612">
         - Settings</span></span><br><span data-ttu-id="5934c-613">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-613">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-614">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="5934c-614">Office for iOS</span></span></td>
    <td> <span data-ttu-id="5934c-615">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5934c-615">- Content</span></span><br><span data-ttu-id="5934c-616">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-616">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="5934c-617">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-617">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="5934c-618">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5934c-618">-ActiveView</span></span><br><span data-ttu-id="5934c-619">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5934c-619">
         -CompressedFile</span></span><br><span data-ttu-id="5934c-620">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-620">
         -DocumentEvents</span></span><br><span data-ttu-id="5934c-621">
         - File</span><span class="sxs-lookup"><span data-stu-id="5934c-621">
         - File</span></span><br><span data-ttu-id="5934c-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5934c-622">
         -PdfFile</span></span><br><span data-ttu-id="5934c-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5934c-623">
         - Selection</span></span><br><span data-ttu-id="5934c-624">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-624">
         - Settings</span></span><br><span data-ttu-id="5934c-625">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-625">
         -TextCoercion</span></span><br><span data-ttu-id="5934c-626">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-626">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-627">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="5934c-627">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="5934c-628">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5934c-628">- Content</span></span><br><span data-ttu-id="5934c-629">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-629">
         - Taskpane</span></span><br><span data-ttu-id="5934c-630">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-630">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5934c-631">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-631">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5934c-632">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5934c-632">-ActiveView</span></span><br><span data-ttu-id="5934c-633">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5934c-633">
         -CompressedFile</span></span><br><span data-ttu-id="5934c-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-634">
         -DocumentEvents</span></span><br><span data-ttu-id="5934c-635">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="5934c-635">
         - File</span></span><br><span data-ttu-id="5934c-636">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-636">
         -ImageCoercion</span></span><br><span data-ttu-id="5934c-637">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5934c-637">
         -PdfFile</span></span><br><span data-ttu-id="5934c-638">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5934c-638">
         - Selection</span></span><br><span data-ttu-id="5934c-639">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-639">
         - Settings</span></span><br><span data-ttu-id="5934c-640">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-640">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-641">Mac 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="5934c-641">Office for Mac</span></span></td>
    <td> <span data-ttu-id="5934c-642">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5934c-642">- Content</span></span><br><span data-ttu-id="5934c-643">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-643">
         - Taskpane</span></span><br><span data-ttu-id="5934c-644">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-644">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5934c-645">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-645">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5934c-646">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5934c-646">-ActiveView</span></span><br><span data-ttu-id="5934c-647">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5934c-647">
         -CompressedFile</span></span><br><span data-ttu-id="5934c-648">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-648">
         -DocumentEvents</span></span><br><span data-ttu-id="5934c-649">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="5934c-649">
         - File</span></span><br><span data-ttu-id="5934c-650">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-650">
         -ImageCoercion</span></span><br><span data-ttu-id="5934c-651">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5934c-651">
         -PdfFile</span></span><br><span data-ttu-id="5934c-652">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5934c-652">
         - Selection</span></span><br><span data-ttu-id="5934c-653">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-653">
         - Settings</span></span><br><span data-ttu-id="5934c-654">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-654">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="5934c-655">OneNote</span><span class="sxs-lookup"><span data-stu-id="5934c-655">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5934c-656">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5934c-656">Platform</span></span></th>
    <th><span data-ttu-id="5934c-657">拡張点</span><span class="sxs-lookup"><span data-stu-id="5934c-657">Extension points</span></span></th>
    <th><span data-ttu-id="5934c-658">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="5934c-658">API requirement sets</span></span></th>
    <th><span data-ttu-id="5934c-659"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5934c-659"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="5934c-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="5934c-660">Office Online</span></span></td>
    <td> <span data-ttu-id="5934c-661">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5934c-661">- Content</span></span><br><span data-ttu-id="5934c-662">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5934c-662">
         - Taskpane</span></span><br><span data-ttu-id="5934c-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5934c-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5934c-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="5934c-665">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5934c-665">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5934c-666">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5934c-666">-DocumentEvents</span></span><br><span data-ttu-id="5934c-667">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-667">
         -HtmlCoercion</span></span><br><span data-ttu-id="5934c-668">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-668">
         -ImageCoercion</span></span><br><span data-ttu-id="5934c-669">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="5934c-669">
         - Settings</span></span><br><span data-ttu-id="5934c-670">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5934c-670">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="5934c-671">関連項目</span><span class="sxs-lookup"><span data-stu-id="5934c-671">See also</span></span>

- [<span data-ttu-id="5934c-672">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="5934c-672">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="5934c-673">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="5934c-673">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="5934c-674">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="5934c-674">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="5934c-675">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="5934c-675">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
