---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、Word、Outlook、PowerPoint、および OneNote のサポートされる要件セット。
ms.date: 09/19/2018
ms.openlocfilehash: 09fb72c88bd0496c413f94b7ba4149192380d664
ms.sourcegitcommit: e7e4d08569a01c69168bb005188e9a1e628304b9
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/22/2018
ms.locfileid: "24967705"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="9d8d2-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="9d8d2-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="9d8d2-104">Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。</span><span class="sxs-lookup"><span data-stu-id="9d8d2-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="9d8d2-105">次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API の要件セットが含まれています。</span><span class="sxs-lookup"><span data-stu-id="9d8d2-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="9d8d2-106">表のセルにアスタリスク ( \* ) が含まれる場合は、準備中です。</span><span class="sxs-lookup"><span data-stu-id="9d8d2-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="9d8d2-107">Project または Access の要件セットについては、「[Office の共有要件セット](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9d8d2-107">For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="9d8d2-p103">MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。このバージョンには、ExcelApi 1.1、WordApi 1.1、および共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="9d8d2-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="9d8d2-110">Excel</span><span class="sxs-lookup"><span data-stu-id="9d8d2-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="9d8d2-111">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="9d8d2-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="9d8d2-112">拡張点</span><span class="sxs-lookup"><span data-stu-id="9d8d2-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="9d8d2-113">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="9d8d2-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="9d8d2-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="9d8d2-115">Office Online</span></span></td>
    <td> <span data-ttu-id="9d8d2-116">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-116">- Taskpane</span></span><br><span data-ttu-id="9d8d2-117">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-117">
        - Content</span></span><br><span data-ttu-id="9d8d2-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="9d8d2-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="9d8d2-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9d8d2-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9d8d2-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9d8d2-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9d8d2-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9d8d2-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9d8d2-125">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="9d8d2-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9d8d2-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-127">
        -BindingEvents</span></span><br><span data-ttu-id="9d8d2-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-128">
        -CompressedFile</span></span><br><span data-ttu-id="9d8d2-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-129">
        -DocumentEvents</span></span><br><span data-ttu-id="9d8d2-130">
        - File</span><span class="sxs-lookup"><span data-stu-id="9d8d2-130">
        - File</span></span><br><span data-ttu-id="9d8d2-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-131">
        -MatrixBindings</span></span><br><span data-ttu-id="9d8d2-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-132">
        -MatrixCoercion</span></span><br><span data-ttu-id="9d8d2-133">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9d8d2-133">
        - Selection</span></span><br><span data-ttu-id="9d8d2-134">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-134">
        - Settings</span></span><br><span data-ttu-id="9d8d2-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-135">
        -TableBindings</span></span><br><span data-ttu-id="9d8d2-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-136">
        -TableCoercion</span></span><br><span data-ttu-id="9d8d2-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-137">
        -TextBindings</span></span><br><span data-ttu-id="9d8d2-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-138">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-139">Windows 用 Office 2013</span><span class="sxs-lookup"><span data-stu-id="9d8d2-139">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="9d8d2-140">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-140">
        - Taskpane</span></span><br><span data-ttu-id="9d8d2-141">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-141">
        - Content</span></span></td>
    <td>  <span data-ttu-id="9d8d2-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9d8d2-143">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-143">
        -BindingEvents</span></span><br><span data-ttu-id="9d8d2-144">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-144">
        -CompressedFile</span></span><br><span data-ttu-id="9d8d2-145">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-145">
        -DocumentEvents</span></span><br><span data-ttu-id="9d8d2-146">
        - File</span><span class="sxs-lookup"><span data-stu-id="9d8d2-146">
        - File</span></span><br><span data-ttu-id="9d8d2-147">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-147">
        -ImageCoercion</span></span><br><span data-ttu-id="9d8d2-148">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-148">
        -MatrixBindings</span></span><br><span data-ttu-id="9d8d2-149">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-149">
        -MatrixCoercion</span></span><br><span data-ttu-id="9d8d2-150">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9d8d2-150">
        - Selection</span></span><br><span data-ttu-id="9d8d2-151">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-151">
        - Settings</span></span><br><span data-ttu-id="9d8d2-152">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-152">
        -TableBindings</span></span><br><span data-ttu-id="9d8d2-153">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-153">
        -TableCoercion</span></span><br><span data-ttu-id="9d8d2-154">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-154">
        -TextBindings</span></span><br><span data-ttu-id="9d8d2-155">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-155">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-156">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="9d8d2-156">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="9d8d2-157">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-157">- Taskpane</span></span><br><span data-ttu-id="9d8d2-158">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-158">
        - Content</span></span><br><span data-ttu-id="9d8d2-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="9d8d2-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9d8d2-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9d8d2-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9d8d2-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9d8d2-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9d8d2-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9d8d2-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-166">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="9d8d2-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9d8d2-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-168">-BindingEvents</span></span><br><span data-ttu-id="9d8d2-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-169">
        -CompressedFile</span></span><br><span data-ttu-id="9d8d2-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-170">
        -DocumentEvents</span></span><br><span data-ttu-id="9d8d2-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="9d8d2-171">
        - File</span></span><br><span data-ttu-id="9d8d2-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-172">
        -ImageCoercion</span></span><br><span data-ttu-id="9d8d2-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-173">
        -MatrixBindings</span></span><br><span data-ttu-id="9d8d2-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-174">
        -MatrixCoercion</span></span><br><span data-ttu-id="9d8d2-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="9d8d2-175">
        - Selection</span></span><br><span data-ttu-id="9d8d2-176">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="9d8d2-176">
        - Settings</span></span><br><span data-ttu-id="9d8d2-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-177">
        -TableBindings</span></span><br><span data-ttu-id="9d8d2-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-178">
        -TableCoercion</span></span><br><span data-ttu-id="9d8d2-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-179">
        -TextBindings</span></span><br><span data-ttu-id="9d8d2-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-181">iOS 用 Office</span><span class="sxs-lookup"><span data-stu-id="9d8d2-181">Office for iOS</span></span></td>
    <td><span data-ttu-id="9d8d2-182">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-182">- Taskpane</span></span><br><span data-ttu-id="9d8d2-183">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-183">
        - Content</span></span></td>
    <td><span data-ttu-id="9d8d2-184">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-184">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9d8d2-185">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-185">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9d8d2-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9d8d2-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9d8d2-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9d8d2-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9d8d2-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-190">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="9d8d2-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9d8d2-192">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-192">-BindingEvents</span></span><br><span data-ttu-id="9d8d2-193">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-193">
        -CompressedFile</span></span><br><span data-ttu-id="9d8d2-194">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-194">
        -DocumentEvents</span></span><br><span data-ttu-id="9d8d2-195">
        - File</span><span class="sxs-lookup"><span data-stu-id="9d8d2-195">
        - File</span></span><br><span data-ttu-id="9d8d2-196">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-196">
        -ImageCoercion</span></span><br><span data-ttu-id="9d8d2-197">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-197">
        -MatrixBindings</span></span><br><span data-ttu-id="9d8d2-198">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-198">
        -MatrixCoercion</span></span><br><span data-ttu-id="9d8d2-199">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="9d8d2-199">
        - Selection</span></span><br><span data-ttu-id="9d8d2-200">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="9d8d2-200">
        - Settings</span></span><br><span data-ttu-id="9d8d2-201">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-201">
        -TableBindings</span></span><br><span data-ttu-id="9d8d2-202">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-202">
        -TableCoercion</span></span><br><span data-ttu-id="9d8d2-203">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-203">
        -TextBindings</span></span><br><span data-ttu-id="9d8d2-204">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-204">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-205">Mac 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="9d8d2-205">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="9d8d2-206">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-206">- Taskpane</span></span><br><span data-ttu-id="9d8d2-207">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-207">
        - Content</span></span><br><span data-ttu-id="9d8d2-208">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-208">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="9d8d2-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="9d8d2-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="9d8d2-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="9d8d2-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="9d8d2-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="9d8d2-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="9d8d2-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-215">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="9d8d2-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="9d8d2-217">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-217">-BindingEvents</span></span><br><span data-ttu-id="9d8d2-218">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-218">
        -CompressedFile</span></span><br><span data-ttu-id="9d8d2-219">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-219">
        -DocumentEvents</span></span><br><span data-ttu-id="9d8d2-220">
        - File</span><span class="sxs-lookup"><span data-stu-id="9d8d2-220">
        - File</span></span><br><span data-ttu-id="9d8d2-221">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-221">
        -ImageCoercion</span></span><br><span data-ttu-id="9d8d2-222">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-222">
        -MatrixBindings</span></span><br><span data-ttu-id="9d8d2-223">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-223">
        -MatrixCoercion</span></span><br><span data-ttu-id="9d8d2-224">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-224">
        -PdfFile</span></span><br><span data-ttu-id="9d8d2-225">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="9d8d2-225">
        - Selection</span></span><br><span data-ttu-id="9d8d2-226">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="9d8d2-226">
        - Settings</span></span><br><span data-ttu-id="9d8d2-227">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-227">
        -TableBindings</span></span><br><span data-ttu-id="9d8d2-228">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-228">
        -TableCoercion</span></span><br><span data-ttu-id="9d8d2-229">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-229">
        -TextBindings</span></span><br><span data-ttu-id="9d8d2-230">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-230">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="9d8d2-231">Outlook</span><span class="sxs-lookup"><span data-stu-id="9d8d2-231">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9d8d2-232">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="9d8d2-232">Platform</span></span></th>
    <th><span data-ttu-id="9d8d2-233">拡張点</span><span class="sxs-lookup"><span data-stu-id="9d8d2-233">Extension points</span></span></th>
    <th><span data-ttu-id="9d8d2-234">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="9d8d2-234">API requirement sets</span></span></th>
    <th><span data-ttu-id="9d8d2-235"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-235"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-236">Office Online</span><span class="sxs-lookup"><span data-stu-id="9d8d2-236">Office Online</span></span></td>
    <td> <span data-ttu-id="9d8d2-237">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="9d8d2-237">- Mail Read</span></span><br><span data-ttu-id="9d8d2-238">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="9d8d2-238">
      - Mail Compose</span></span><br><span data-ttu-id="9d8d2-239">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-239">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9d8d2-240">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-240">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9d8d2-241">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-241">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9d8d2-242">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-242">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9d8d2-243">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-243">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9d8d2-244">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-244">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9d8d2-245">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-245">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="9d8d2-246">使用不可</span><span class="sxs-lookup"><span data-stu-id="9d8d2-246">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-247">Windows 用 Office 2013</span><span class="sxs-lookup"><span data-stu-id="9d8d2-247">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="9d8d2-248">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="9d8d2-248">- Mail Read</span></span><br><span data-ttu-id="9d8d2-249">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="9d8d2-249">
      - Mail Compose</span></span><br><span data-ttu-id="9d8d2-250">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-250">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9d8d2-251">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-251">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9d8d2-252">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-252">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9d8d2-253">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-253">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9d8d2-254">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-254">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="9d8d2-255">使用不可</span><span class="sxs-lookup"><span data-stu-id="9d8d2-255">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-256">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="9d8d2-256">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="9d8d2-257">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="9d8d2-257">- Mail Read</span></span><br><span data-ttu-id="9d8d2-258">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="9d8d2-258">
      - Mail Compose</span></span><br><span data-ttu-id="9d8d2-259">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-259">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="9d8d2-260">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="9d8d2-260">
      - Modules</span></span></td>
    <td> <span data-ttu-id="9d8d2-261">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-261">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9d8d2-262">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-262">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9d8d2-263">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-263">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9d8d2-264">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-264">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9d8d2-265">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-265">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9d8d2-266">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-266">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="9d8d2-267">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-267">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="9d8d2-268">使用不可</span><span class="sxs-lookup"><span data-stu-id="9d8d2-268">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-269">iOS 用 Office</span><span class="sxs-lookup"><span data-stu-id="9d8d2-269">Office for iOS</span></span></td>
    <td> <span data-ttu-id="9d8d2-270">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="9d8d2-270">- Mail Read</span></span><br><span data-ttu-id="9d8d2-271">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-271">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9d8d2-272">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-272">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9d8d2-273">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-273">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9d8d2-274">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-274">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9d8d2-275">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-275">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9d8d2-276">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-276">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="9d8d2-277">使用不可</span><span class="sxs-lookup"><span data-stu-id="9d8d2-277">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-278">Mac 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="9d8d2-278">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="9d8d2-279">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="9d8d2-279">- Mail Read</span></span><br><span data-ttu-id="9d8d2-280">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="9d8d2-280">
      - Mail Compose</span></span><br><span data-ttu-id="9d8d2-281">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-281">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9d8d2-282">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-282">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9d8d2-283">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-283">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9d8d2-284">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-284">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9d8d2-285">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-285">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9d8d2-286">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-286">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="9d8d2-287">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-287">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="9d8d2-288">使用不可</span><span class="sxs-lookup"><span data-stu-id="9d8d2-288">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-289">Android 用 Office</span><span class="sxs-lookup"><span data-stu-id="9d8d2-289">Office for Android</span></span></td>
    <td> <span data-ttu-id="9d8d2-290">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="9d8d2-290">- Mail Read</span></span><br><span data-ttu-id="9d8d2-291">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-291">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9d8d2-292">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-292">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="9d8d2-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="9d8d2-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="9d8d2-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="9d8d2-296">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-296">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="9d8d2-297">使用不可</span><span class="sxs-lookup"><span data-stu-id="9d8d2-297">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="9d8d2-298">Word</span><span class="sxs-lookup"><span data-stu-id="9d8d2-298">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9d8d2-299">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="9d8d2-299">Platform</span></span></th>
    <th><span data-ttu-id="9d8d2-300">拡張点</span><span class="sxs-lookup"><span data-stu-id="9d8d2-300">Extension points</span></span></th>
    <th><span data-ttu-id="9d8d2-301">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="9d8d2-301">API requirement sets</span></span></th>
    <th><span data-ttu-id="9d8d2-302"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-302"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-303">Office Online</span><span class="sxs-lookup"><span data-stu-id="9d8d2-303">Office Online</span></span></td>
    <td> <span data-ttu-id="9d8d2-304">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-304">- Taskpane</span></span><br><span data-ttu-id="9d8d2-305">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-305">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9d8d2-306">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-306">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9d8d2-307">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-307">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9d8d2-308">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-308">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9d8d2-309">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-309">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9d8d2-310">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-310">-BindingEvents</span></span><br><span data-ttu-id="9d8d2-311">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="9d8d2-311">
         -</span></span><br><span data-ttu-id="9d8d2-312">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-312">
         -DocumentEvents</span></span><br><span data-ttu-id="9d8d2-313">
         - File</span><span class="sxs-lookup"><span data-stu-id="9d8d2-313">
         - File</span></span><br><span data-ttu-id="9d8d2-314">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-314">
         -HtmlCoercion</span></span><br><span data-ttu-id="9d8d2-315">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-315">
         -ImageCoercion</span></span><br><span data-ttu-id="9d8d2-316">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-316">
         -MatrixBindings</span></span><br><span data-ttu-id="9d8d2-317">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-317">
         -MatrixCoercion</span></span><br><span data-ttu-id="9d8d2-318">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-318">
         -OoxmlCoercion</span></span><br><span data-ttu-id="9d8d2-319">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-319">
         -PdfFile</span></span><br><span data-ttu-id="9d8d2-320">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="9d8d2-320">
         - Selection</span></span><br><span data-ttu-id="9d8d2-321">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="9d8d2-321">
         - Settings</span></span><br><span data-ttu-id="9d8d2-322">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-322">
         -TableBindings</span></span><br><span data-ttu-id="9d8d2-323">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-323">
         -TableCoercion</span></span><br><span data-ttu-id="9d8d2-324">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-324">
         -TextBindings</span></span><br><span data-ttu-id="9d8d2-325">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-325">
         -TextCoercion</span></span><br><span data-ttu-id="9d8d2-326">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-326">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-327">Windows 用 Office 2013</span><span class="sxs-lookup"><span data-stu-id="9d8d2-327">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="9d8d2-328">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-328">- Taskpane</span></span></td>
    <td> <span data-ttu-id="9d8d2-329">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-329">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9d8d2-330">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-330">-BindingEvents</span></span><br><span data-ttu-id="9d8d2-331">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-331">
         -CompressedFile</span></span><br><span data-ttu-id="9d8d2-332">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="9d8d2-332">
         -</span></span><br><span data-ttu-id="9d8d2-333">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-333">
         -DocumentEvents</span></span><br><span data-ttu-id="9d8d2-334">
         - File</span><span class="sxs-lookup"><span data-stu-id="9d8d2-334">
         - File</span></span><br><span data-ttu-id="9d8d2-335">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-335">
         -HtmlCoercion</span></span><br><span data-ttu-id="9d8d2-336">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-336">
         -ImageCoercion</span></span><br><span data-ttu-id="9d8d2-337">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-337">
         -MatrixBindings</span></span><br><span data-ttu-id="9d8d2-338">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-338">
         -MatrixCoercion</span></span><br><span data-ttu-id="9d8d2-339">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-339">
         -OoxmlCoercion</span></span><br><span data-ttu-id="9d8d2-340">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-340">
         -PdfFile</span></span><br><span data-ttu-id="9d8d2-341">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="9d8d2-341">
         - Selection</span></span><br><span data-ttu-id="9d8d2-342">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="9d8d2-342">
         - Settings</span></span><br><span data-ttu-id="9d8d2-343">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-343">
         -TableBindings</span></span><br><span data-ttu-id="9d8d2-344">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-344">
         -TableCoercion</span></span><br><span data-ttu-id="9d8d2-345">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-345">
         -TextBindings</span></span><br><span data-ttu-id="9d8d2-346">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-346">
         -TextCoercion</span></span><br><span data-ttu-id="9d8d2-347">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-347">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-348">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="9d8d2-348">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="9d8d2-349">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-349">- Taskpane</span></span><br><span data-ttu-id="9d8d2-350">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-350">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9d8d2-351">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-351">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9d8d2-352">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-352">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9d8d2-353">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-353">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9d8d2-354">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-354">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9d8d2-355">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-355">-BindingEvents</span></span><br><span data-ttu-id="9d8d2-356">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-356">
         -CompressedFile</span></span><br><span data-ttu-id="9d8d2-357">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="9d8d2-357">
         -</span></span><br><span data-ttu-id="9d8d2-358">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-358">
         -DocumentEvents</span></span><br><span data-ttu-id="9d8d2-359">
         - File</span><span class="sxs-lookup"><span data-stu-id="9d8d2-359">
         - File</span></span><br><span data-ttu-id="9d8d2-360">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-360">
         -HtmlCoercion</span></span><br><span data-ttu-id="9d8d2-361">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-361">
         -ImageCoercion</span></span><br><span data-ttu-id="9d8d2-362">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-362">
         -MatrixBindings</span></span><br><span data-ttu-id="9d8d2-363">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-363">
         -MatrixCoercion</span></span><br><span data-ttu-id="9d8d2-364">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-364">
         -OoxmlCoercion</span></span><br><span data-ttu-id="9d8d2-365">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-365">
         -PdfFile</span></span><br><span data-ttu-id="9d8d2-366">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="9d8d2-366">
         - Selection</span></span><br><span data-ttu-id="9d8d2-367">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="9d8d2-367">
         - Settings</span></span><br><span data-ttu-id="9d8d2-368">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-368">
         -TableBindings</span></span><br><span data-ttu-id="9d8d2-369">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-369">
         -TableCoercion</span></span><br><span data-ttu-id="9d8d2-370">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-370">
         -TextBindings</span></span><br><span data-ttu-id="9d8d2-371">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-371">
         -TextCoercion</span></span><br><span data-ttu-id="9d8d2-372">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-372">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-373">iOS 用 Office</span><span class="sxs-lookup"><span data-stu-id="9d8d2-373">Office for iOS</span></span></td>
    <td> <span data-ttu-id="9d8d2-374">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-374">- Taskpane</span></span></td>
    <td> <span data-ttu-id="9d8d2-375">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-375">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9d8d2-376">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-376">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9d8d2-377">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-377">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9d8d2-378">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="9d8d2-378">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="9d8d2-379">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-379">-BindingEvents</span></span><br><span data-ttu-id="9d8d2-380">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-380">
         -CompressedFile</span></span><br><span data-ttu-id="9d8d2-381">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="9d8d2-381">
         -</span></span><br><span data-ttu-id="9d8d2-382">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-382">
         -DocumentEvents</span></span><br><span data-ttu-id="9d8d2-383">
         - File</span><span class="sxs-lookup"><span data-stu-id="9d8d2-383">
         - File</span></span><br><span data-ttu-id="9d8d2-384">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-384">
         -HtmlCoercion</span></span><br><span data-ttu-id="9d8d2-385">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-385">
         -ImageCoercion</span></span><br><span data-ttu-id="9d8d2-386">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-386">
         -MatrixBindings</span></span><br><span data-ttu-id="9d8d2-387">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-387">
         -MatrixCoercion</span></span><br><span data-ttu-id="9d8d2-388">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-388">
         -OoxmlCoercion</span></span><br><span data-ttu-id="9d8d2-389">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-389">
         -PdfFile</span></span><br><span data-ttu-id="9d8d2-390">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="9d8d2-390">
         - Selection</span></span><br><span data-ttu-id="9d8d2-391">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="9d8d2-391">
         - Settings</span></span><br><span data-ttu-id="9d8d2-392">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-392">
         -TableBindings</span></span><br><span data-ttu-id="9d8d2-393">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-393">
         -TableCoercion</span></span><br><span data-ttu-id="9d8d2-394">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-394">
         -TextBindings</span></span><br><span data-ttu-id="9d8d2-395">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-395">
         -TextCoercion</span></span><br><span data-ttu-id="9d8d2-396">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-396">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-397">Mac 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="9d8d2-397">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="9d8d2-398">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-398">- Taskpane</span></span><br><span data-ttu-id="9d8d2-399">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-399">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9d8d2-400">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-400">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="9d8d2-401">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-401">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="9d8d2-402">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-402">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="9d8d2-403">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="9d8d2-403">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="9d8d2-404">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-404">-BindingEvents</span></span><br><span data-ttu-id="9d8d2-405">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-405">
         -CompressedFile</span></span><br><span data-ttu-id="9d8d2-406">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="9d8d2-406">
         -</span></span><br><span data-ttu-id="9d8d2-407">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-407">
         -DocumentEvents</span></span><br><span data-ttu-id="9d8d2-408">
         - File</span><span class="sxs-lookup"><span data-stu-id="9d8d2-408">
         - File</span></span><br><span data-ttu-id="9d8d2-409">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-409">
         -HtmlCoercion</span></span><br><span data-ttu-id="9d8d2-410">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-410">
         -ImageCoercion</span></span><br><span data-ttu-id="9d8d2-411">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-411">
         -MatrixBindings</span></span><br><span data-ttu-id="9d8d2-412">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-412">
         -MatrixCoercion</span></span><br><span data-ttu-id="9d8d2-413">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-413">
         -OoxmlCoercion</span></span><br><span data-ttu-id="9d8d2-414">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-414">
         -PdfFile</span></span><br><span data-ttu-id="9d8d2-415">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="9d8d2-415">
         - Selection</span></span><br><span data-ttu-id="9d8d2-416">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="9d8d2-416">
         - Settings</span></span><br><span data-ttu-id="9d8d2-417">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-417">
         -TableBindings</span></span><br><span data-ttu-id="9d8d2-418">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-418">
         -TableCoercion</span></span><br><span data-ttu-id="9d8d2-419">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="9d8d2-419">
         -TextBindings</span></span><br><span data-ttu-id="9d8d2-420">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-420">
         -TextCoercion</span></span><br><span data-ttu-id="9d8d2-421">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-421">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="9d8d2-422">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="9d8d2-422">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9d8d2-423">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="9d8d2-423">Platform</span></span></th>
    <th><span data-ttu-id="9d8d2-424">拡張点</span><span class="sxs-lookup"><span data-stu-id="9d8d2-424">Extension points</span></span></th>
    <th><span data-ttu-id="9d8d2-425">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="9d8d2-425">API requirement sets</span></span></th>
    <th><span data-ttu-id="9d8d2-426"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-426"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-427">Office Online</span><span class="sxs-lookup"><span data-stu-id="9d8d2-427">Office Online</span></span></td>
    <td> <span data-ttu-id="9d8d2-428">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-428">- Content</span></span><br><span data-ttu-id="9d8d2-429">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-429">
         - Taskpane</span></span><br><span data-ttu-id="9d8d2-430">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-430">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9d8d2-431">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-431">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9d8d2-432">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9d8d2-432">-ActiveView</span></span><br><span data-ttu-id="9d8d2-433">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-433">
         -CompressedFile</span></span><br><span data-ttu-id="9d8d2-434">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-434">
         -DocumentEvents</span></span><br><span data-ttu-id="9d8d2-435">
         - File</span><span class="sxs-lookup"><span data-stu-id="9d8d2-435">
         - File</span></span><br><span data-ttu-id="9d8d2-436">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-436">
         -ImageCoercion</span></span><br><span data-ttu-id="9d8d2-437">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-437">
         -PdfFile</span></span><br><span data-ttu-id="9d8d2-438">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="9d8d2-438">
         - Selection</span></span><br><span data-ttu-id="9d8d2-439">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="9d8d2-439">
         - Settings</span></span><br><span data-ttu-id="9d8d2-440">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-440">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-441">Windows 用 Office 2013</span><span class="sxs-lookup"><span data-stu-id="9d8d2-441">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="9d8d2-442">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-442">- Content</span></span><br><span data-ttu-id="9d8d2-443">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-443">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="9d8d2-444">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="9d8d2-444">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="9d8d2-445">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9d8d2-445">-ActiveView</span></span><br><span data-ttu-id="9d8d2-446">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-446">
         -CompressedFile</span></span><br><span data-ttu-id="9d8d2-447">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-447">
         -DocumentEvents</span></span><br><span data-ttu-id="9d8d2-448">
         - File</span><span class="sxs-lookup"><span data-stu-id="9d8d2-448">
         - File</span></span><br><span data-ttu-id="9d8d2-449">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-449">
         -ImageCoercion</span></span><br><span data-ttu-id="9d8d2-450">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-450">
         -PdfFile</span></span><br><span data-ttu-id="9d8d2-451">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="9d8d2-451">
         - Selection</span></span><br><span data-ttu-id="9d8d2-452">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="9d8d2-452">
         - Settings</span></span><br><span data-ttu-id="9d8d2-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-453">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-454">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="9d8d2-454">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="9d8d2-455">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-455">- Content</span></span><br><span data-ttu-id="9d8d2-456">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-456">
         - Taskpane</span></span><br><span data-ttu-id="9d8d2-457">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-457">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9d8d2-458">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-458">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9d8d2-459">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9d8d2-459">-ActiveView</span></span><br><span data-ttu-id="9d8d2-460">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-460">
         -CompressedFile</span></span><br><span data-ttu-id="9d8d2-461">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-461">
         -DocumentEvents</span></span><br><span data-ttu-id="9d8d2-462">
         - File</span><span class="sxs-lookup"><span data-stu-id="9d8d2-462">
         - File</span></span><br><span data-ttu-id="9d8d2-463">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-463">
         -ImageCoercion</span></span><br><span data-ttu-id="9d8d2-464">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-464">
         -PdfFile</span></span><br><span data-ttu-id="9d8d2-465">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="9d8d2-465">
         - Selection</span></span><br><span data-ttu-id="9d8d2-466">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="9d8d2-466">
         - Settings</span></span><br><span data-ttu-id="9d8d2-467">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-467">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-468">iOS 用 Office</span><span class="sxs-lookup"><span data-stu-id="9d8d2-468">Office for iOS</span></span></td>
    <td> <span data-ttu-id="9d8d2-469">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-469">- Content</span></span><br><span data-ttu-id="9d8d2-470">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-470">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="9d8d2-471">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-471">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="9d8d2-472">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9d8d2-472">-ActiveView</span></span><br><span data-ttu-id="9d8d2-473">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-473">
         -CompressedFile</span></span><br><span data-ttu-id="9d8d2-474">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-474">
         -DocumentEvents</span></span><br><span data-ttu-id="9d8d2-475">
         - File</span><span class="sxs-lookup"><span data-stu-id="9d8d2-475">
         - File</span></span><br><span data-ttu-id="9d8d2-476">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-476">
         -PdfFile</span></span><br><span data-ttu-id="9d8d2-477">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="9d8d2-477">
         - Selection</span></span><br><span data-ttu-id="9d8d2-478">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="9d8d2-478">
         - Settings</span></span><br><span data-ttu-id="9d8d2-479">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-479">
         -TextCoercion</span></span><br><span data-ttu-id="9d8d2-480">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-480">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-481">Mac 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="9d8d2-481">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="9d8d2-482">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-482">- Content</span></span><br><span data-ttu-id="9d8d2-483">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-483">
         - Taskpane</span></span><br><span data-ttu-id="9d8d2-484">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-484">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9d8d2-485">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-485">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9d8d2-486">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="9d8d2-486">-ActiveView</span></span><br><span data-ttu-id="9d8d2-487">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-487">
         -CompressedFile</span></span><br><span data-ttu-id="9d8d2-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-488">
         -DocumentEvents</span></span><br><span data-ttu-id="9d8d2-489">
         - File</span><span class="sxs-lookup"><span data-stu-id="9d8d2-489">
         - File</span></span><br><span data-ttu-id="9d8d2-490">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-490">
         -ImageCoercion</span></span><br><span data-ttu-id="9d8d2-491">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="9d8d2-491">
         -PdfFile</span></span><br><span data-ttu-id="9d8d2-492">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="9d8d2-492">
         - Selection</span></span><br><span data-ttu-id="9d8d2-493">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="9d8d2-493">
         - Settings</span></span><br><span data-ttu-id="9d8d2-494">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-494">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="9d8d2-495">OneNote</span><span class="sxs-lookup"><span data-stu-id="9d8d2-495">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="9d8d2-496">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="9d8d2-496">Platform</span></span></th>
    <th><span data-ttu-id="9d8d2-497">拡張点</span><span class="sxs-lookup"><span data-stu-id="9d8d2-497">Extension points</span></span></th>
    <th><span data-ttu-id="9d8d2-498">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="9d8d2-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="9d8d2-499"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-499"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="9d8d2-500">Office Online</span><span class="sxs-lookup"><span data-stu-id="9d8d2-500">Office Online</span></span></td>
    <td> <span data-ttu-id="9d8d2-501">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-501">- Content</span></span><br><span data-ttu-id="9d8d2-502">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="9d8d2-502">
         - Taskpane</span></span><br><span data-ttu-id="9d8d2-503">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-503">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="9d8d2-504">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-504">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="9d8d2-505">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="9d8d2-505">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="9d8d2-506">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="9d8d2-506">-DocumentEvents</span></span><br><span data-ttu-id="9d8d2-507">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-507">
         -HtmlCoercion</span></span><br><span data-ttu-id="9d8d2-508">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-508">
         -ImageCoercion</span></span><br><span data-ttu-id="9d8d2-509">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="9d8d2-509">
         - Settings</span></span><br><span data-ttu-id="9d8d2-510">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="9d8d2-510">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="9d8d2-511">関連項目</span><span class="sxs-lookup"><span data-stu-id="9d8d2-511">See also</span></span>

- [<span data-ttu-id="9d8d2-512">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="9d8d2-512">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="9d8d2-513">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="9d8d2-513">Common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="9d8d2-514">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="9d8d2-514">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="9d8d2-515">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="9d8d2-515">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office)
