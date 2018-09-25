---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、Word、Outlook、PowerPoint、および OneNote のサポートされる要件セット。
ms.date: 09/24/2018
ms.openlocfilehash: b06602e35ec906866ad16d667036a4cbaff2d89e
ms.sourcegitcommit: 8ce9a8d7f41d96879c39cc5527a3007dff25bee8
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/24/2018
ms.locfileid: "24985824"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="91c68-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="91c68-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="91c68-104">Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。</span><span class="sxs-lookup"><span data-stu-id="91c68-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="91c68-105">次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API の要件セットが含まれています。</span><span class="sxs-lookup"><span data-stu-id="91c68-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="91c68-106">表のセルにアスタリスク ( \* ) が含まれる場合は、準備中です。</span><span class="sxs-lookup"><span data-stu-id="91c68-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="91c68-107">Project または Access の要件セットについては、「[Office の共有要件セット](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="91c68-107">For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="91c68-p103">MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。このバージョンには、ExcelApi 1.1、WordApi 1.1、および共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="91c68-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="91c68-110">Excel</span><span class="sxs-lookup"><span data-stu-id="91c68-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="91c68-111">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="91c68-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="91c68-112">拡張点</span><span class="sxs-lookup"><span data-stu-id="91c68-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="91c68-113">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="91c68-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="91c68-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="91c68-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="91c68-115">Office Online</span></span></td>
    <td> <span data-ttu-id="91c68-116">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-116">- Taskpane</span></span><br><span data-ttu-id="91c68-117">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="91c68-117">
        - Content</span></span><br><span data-ttu-id="91c68-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="91c68-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="91c68-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="91c68-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="91c68-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="91c68-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="91c68-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="91c68-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="91c68-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="91c68-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="91c68-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="91c68-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="91c68-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="91c68-125">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="91c68-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="91c68-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="91c68-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-127">
        -BindingEvents</span></span><br><span data-ttu-id="91c68-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="91c68-128">
        -CompressedFile</span></span><br><span data-ttu-id="91c68-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-129">
        -DocumentEvents</span></span><br><span data-ttu-id="91c68-130">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="91c68-130">
        - File</span></span><br><span data-ttu-id="91c68-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-131">
        -MatrixBindings</span></span><br><span data-ttu-id="91c68-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-132">
        -MatrixCoercion</span></span><br><span data-ttu-id="91c68-133">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-133">
        - Selection</span></span><br><span data-ttu-id="91c68-134">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-134">
        - Settings</span></span><br><span data-ttu-id="91c68-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-135">
        -TableBindings</span></span><br><span data-ttu-id="91c68-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-136">
        -TableCoercion</span></span><br><span data-ttu-id="91c68-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-137">
        -TextBindings</span></span><br><span data-ttu-id="91c68-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-138">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-139">Windows 用 Office 2013</span><span class="sxs-lookup"><span data-stu-id="91c68-139">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="91c68-140">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-140">
        - Taskpane</span></span><br><span data-ttu-id="91c68-141">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="91c68-141">
        - Content</span></span></td>
    <td>  <span data-ttu-id="91c68-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="91c68-143">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-143">
        -BindingEvents</span></span><br><span data-ttu-id="91c68-144">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="91c68-144">
        -CompressedFile</span></span><br><span data-ttu-id="91c68-145">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-145">
        -DocumentEvents</span></span><br><span data-ttu-id="91c68-146">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="91c68-146">
        - File</span></span><br><span data-ttu-id="91c68-147">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-147">
        -ImageCoercion</span></span><br><span data-ttu-id="91c68-148">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-148">
        -MatrixBindings</span></span><br><span data-ttu-id="91c68-149">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-149">
        -MatrixCoercion</span></span><br><span data-ttu-id="91c68-150">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-150">
        - Selection</span></span><br><span data-ttu-id="91c68-151">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-151">
        - Settings</span></span><br><span data-ttu-id="91c68-152">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-152">
        -TableBindings</span></span><br><span data-ttu-id="91c68-153">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-153">
        -TableCoercion</span></span><br><span data-ttu-id="91c68-154">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-154">
        -TextBindings</span></span><br><span data-ttu-id="91c68-155">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-155">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-156">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="91c68-156">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="91c68-157">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-157">- Taskpane</span></span><br><span data-ttu-id="91c68-158">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="91c68-158">
        - Content</span></span><br><span data-ttu-id="91c68-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="91c68-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="91c68-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="91c68-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="91c68-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="91c68-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="91c68-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="91c68-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="91c68-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="91c68-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="91c68-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="91c68-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="91c68-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="91c68-166">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="91c68-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="91c68-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-168">-BindingEvents</span></span><br><span data-ttu-id="91c68-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="91c68-169">
        -CompressedFile</span></span><br><span data-ttu-id="91c68-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-170">
        -DocumentEvents</span></span><br><span data-ttu-id="91c68-171">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="91c68-171">
        - File</span></span><br><span data-ttu-id="91c68-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-172">
        -ImageCoercion</span></span><br><span data-ttu-id="91c68-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-173">
        -MatrixBindings</span></span><br><span data-ttu-id="91c68-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-174">
        -MatrixCoercion</span></span><br><span data-ttu-id="91c68-175">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-175">
        - Selection</span></span><br><span data-ttu-id="91c68-176">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-176">
        - Settings</span></span><br><span data-ttu-id="91c68-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-177">
        -TableBindings</span></span><br><span data-ttu-id="91c68-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-178">
        -TableCoercion</span></span><br><span data-ttu-id="91c68-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-179">
        -TextBindings</span></span><br><span data-ttu-id="91c68-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-181">Windows 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="91c68-181">Office for Windows</span></span></td>
    <td><span data-ttu-id="91c68-182">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-182">- Taskpane</span></span><br><span data-ttu-id="91c68-183">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="91c68-183">
        - Content</span></span><br><span data-ttu-id="91c68-184">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-184">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="91c68-185">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-185">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="91c68-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="91c68-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="91c68-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="91c68-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="91c68-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="91c68-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="91c68-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="91c68-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="91c68-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="91c68-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="91c68-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="91c68-191">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="91c68-192">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-192">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="91c68-193">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-193">-BindingEvents</span></span><br><span data-ttu-id="91c68-194">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="91c68-194">
        -CompressedFile</span></span><br><span data-ttu-id="91c68-195">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-195">
        -DocumentEvents</span></span><br><span data-ttu-id="91c68-196">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="91c68-196">
        - File</span></span><br><span data-ttu-id="91c68-197">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-197">
        -ImageCoercion</span></span><br><span data-ttu-id="91c68-198">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-198">
        -MatrixBindings</span></span><br><span data-ttu-id="91c68-199">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-199">
        -MatrixCoercion</span></span><br><span data-ttu-id="91c68-200">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-200">
        - Selection</span></span><br><span data-ttu-id="91c68-201">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-201">
        - Settings</span></span><br><span data-ttu-id="91c68-202">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-202">
        -TableBindings</span></span><br><span data-ttu-id="91c68-203">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-203">
        -TableCoercion</span></span><br><span data-ttu-id="91c68-204">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-204">
        -TextBindings</span></span><br><span data-ttu-id="91c68-205">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-205">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-206">iOS 用 Office</span><span class="sxs-lookup"><span data-stu-id="91c68-206">Office for iOS</span></span></td>
    <td><span data-ttu-id="91c68-207">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-207">- Taskpane</span></span><br><span data-ttu-id="91c68-208">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="91c68-208">
        - Content</span></span></td>
    <td><span data-ttu-id="91c68-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="91c68-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="91c68-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="91c68-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="91c68-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="91c68-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="91c68-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="91c68-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="91c68-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="91c68-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="91c68-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="91c68-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="91c68-215">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="91c68-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="91c68-217">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-217">-BindingEvents</span></span><br><span data-ttu-id="91c68-218">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="91c68-218">
        -CompressedFile</span></span><br><span data-ttu-id="91c68-219">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-219">
        -DocumentEvents</span></span><br><span data-ttu-id="91c68-220">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="91c68-220">
        - File</span></span><br><span data-ttu-id="91c68-221">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-221">
        -ImageCoercion</span></span><br><span data-ttu-id="91c68-222">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-222">
        -MatrixBindings</span></span><br><span data-ttu-id="91c68-223">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-223">
        -MatrixCoercion</span></span><br><span data-ttu-id="91c68-224">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-224">
        - Selection</span></span><br><span data-ttu-id="91c68-225">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-225">
        - Settings</span></span><br><span data-ttu-id="91c68-226">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-226">
        -TableBindings</span></span><br><span data-ttu-id="91c68-227">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-227">
        -TableCoercion</span></span><br><span data-ttu-id="91c68-228">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-228">
        -TextBindings</span></span><br><span data-ttu-id="91c68-229">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-229">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-230">Mac 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="91c68-230">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="91c68-231">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-231">- Taskpane</span></span><br><span data-ttu-id="91c68-232">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="91c68-232">
        - Content</span></span><br><span data-ttu-id="91c68-233">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-233">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="91c68-234">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-234">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="91c68-235">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="91c68-235">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="91c68-236">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="91c68-236">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="91c68-237">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="91c68-237">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="91c68-238">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="91c68-238">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="91c68-239">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="91c68-239">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="91c68-240">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="91c68-240">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="91c68-241">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-241">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="91c68-242">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-242">-BindingEvents</span></span><br><span data-ttu-id="91c68-243">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="91c68-243">
        -CompressedFile</span></span><br><span data-ttu-id="91c68-244">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-244">
        -DocumentEvents</span></span><br><span data-ttu-id="91c68-245">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="91c68-245">
        - File</span></span><br><span data-ttu-id="91c68-246">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-246">
        -ImageCoercion</span></span><br><span data-ttu-id="91c68-247">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-247">
        -MatrixBindings</span></span><br><span data-ttu-id="91c68-248">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-248">
        -MatrixCoercion</span></span><br><span data-ttu-id="91c68-249">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="91c68-249">
        -PdfFile</span></span><br><span data-ttu-id="91c68-250">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-250">
        - Selection</span></span><br><span data-ttu-id="91c68-251">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-251">
        - Settings</span></span><br><span data-ttu-id="91c68-252">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-252">
        -TableBindings</span></span><br><span data-ttu-id="91c68-253">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-253">
        -TableCoercion</span></span><br><span data-ttu-id="91c68-254">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-254">
        -TextBindings</span></span><br><span data-ttu-id="91c68-255">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-255">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-256">Mac 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="91c68-256">Office for Mac</span></span></td>
    <td><span data-ttu-id="91c68-257">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-257">- Taskpane</span></span><br><span data-ttu-id="91c68-258">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="91c68-258">
        - Content</span></span><br><span data-ttu-id="91c68-259">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-259">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="91c68-260">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-260">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="91c68-261">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="91c68-261">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="91c68-262">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="91c68-262">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="91c68-263">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="91c68-263">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="91c68-264">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="91c68-264">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="91c68-265">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="91c68-265">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="91c68-266">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="91c68-266">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="91c68-267">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-267">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="91c68-268">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-268">-BindingEvents</span></span><br><span data-ttu-id="91c68-269">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="91c68-269">
        -CompressedFile</span></span><br><span data-ttu-id="91c68-270">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-270">
        -DocumentEvents</span></span><br><span data-ttu-id="91c68-271">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="91c68-271">
        - File</span></span><br><span data-ttu-id="91c68-272">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-272">
        -ImageCoercion</span></span><br><span data-ttu-id="91c68-273">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-273">
        -MatrixBindings</span></span><br><span data-ttu-id="91c68-274">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-274">
        -MatrixCoercion</span></span><br><span data-ttu-id="91c68-275">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="91c68-275">
        -PdfFile</span></span><br><span data-ttu-id="91c68-276">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-276">
        - Selection</span></span><br><span data-ttu-id="91c68-277">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-277">
        - Settings</span></span><br><span data-ttu-id="91c68-278">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-278">
        -TableBindings</span></span><br><span data-ttu-id="91c68-279">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-279">
        -TableCoercion</span></span><br><span data-ttu-id="91c68-280">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-280">
        -TextBindings</span></span><br><span data-ttu-id="91c68-281">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-281">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="91c68-282">Outlook</span><span class="sxs-lookup"><span data-stu-id="91c68-282">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="91c68-283">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="91c68-283">Platform</span></span></th>
    <th><span data-ttu-id="91c68-284">拡張点</span><span class="sxs-lookup"><span data-stu-id="91c68-284">Extension points</span></span></th>
    <th><span data-ttu-id="91c68-285">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="91c68-285">API requirement sets</span></span></th>
    <th><span data-ttu-id="91c68-286"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="91c68-286"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-287">Office Online</span><span class="sxs-lookup"><span data-stu-id="91c68-287">Office Online</span></span></td>
    <td> <span data-ttu-id="91c68-288">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="91c68-288">- Mail Read</span></span><br><span data-ttu-id="91c68-289">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="91c68-289">
      - Mail Compose</span></span><br><span data-ttu-id="91c68-290">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-290">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="91c68-291">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-291">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="91c68-292">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="91c68-292">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="91c68-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="91c68-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="91c68-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="91c68-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="91c68-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="91c68-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="91c68-296">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="91c68-296">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="91c68-297">使用不可</span><span class="sxs-lookup"><span data-stu-id="91c68-297">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-298">Windows 用 Office 2013</span><span class="sxs-lookup"><span data-stu-id="91c68-298">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="91c68-299">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="91c68-299">- Mail Read</span></span><br><span data-ttu-id="91c68-300">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="91c68-300">
      - Mail Compose</span></span><br><span data-ttu-id="91c68-301">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-301">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="91c68-302">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-302">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="91c68-303">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="91c68-303">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="91c68-304">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="91c68-304">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="91c68-305">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="91c68-305">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="91c68-306">使用不可</span><span class="sxs-lookup"><span data-stu-id="91c68-306">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-307">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="91c68-307">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="91c68-308">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="91c68-308">- Mail Read</span></span><br><span data-ttu-id="91c68-309">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="91c68-309">
      - Mail Compose</span></span><br><span data-ttu-id="91c68-310">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-310">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="91c68-311">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="91c68-311">
      - Modules</span></span></td>
    <td> <span data-ttu-id="91c68-312">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-312">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="91c68-313">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="91c68-313">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="91c68-314">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="91c68-314">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="91c68-315">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="91c68-315">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="91c68-316">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="91c68-316">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="91c68-317">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="91c68-317">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="91c68-318">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="91c68-318">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="91c68-319">使用不可</span><span class="sxs-lookup"><span data-stu-id="91c68-319">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-320">Windows 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="91c68-320">Office for Windows</span></span></td>
    <td> <span data-ttu-id="91c68-321">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="91c68-321">- Mail Read</span></span><br><span data-ttu-id="91c68-322">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="91c68-322">
      - Mail Compose</span></span><br><span data-ttu-id="91c68-323">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-323">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="91c68-324">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="91c68-324">
      - Modules</span></span></td>
    <td> <span data-ttu-id="91c68-325">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-325">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="91c68-326">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="91c68-326">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="91c68-327">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="91c68-327">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="91c68-328">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="91c68-328">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="91c68-329">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="91c68-329">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="91c68-330">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="91c68-330">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="91c68-331">使用不可</span><span class="sxs-lookup"><span data-stu-id="91c68-331">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-332">iOS 用 Office</span><span class="sxs-lookup"><span data-stu-id="91c68-332">Office for iOS</span></span></td>
    <td> <span data-ttu-id="91c68-333">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="91c68-333">- Mail Read</span></span><br><span data-ttu-id="91c68-334">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-334">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="91c68-335">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-335">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="91c68-336">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="91c68-336">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="91c68-337">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="91c68-337">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="91c68-338">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="91c68-338">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="91c68-339">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="91c68-339">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="91c68-340">使用不可</span><span class="sxs-lookup"><span data-stu-id="91c68-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-341">Mac 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="91c68-341">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="91c68-342">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="91c68-342">- Mail Read</span></span><br><span data-ttu-id="91c68-343">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="91c68-343">
      - Mail Compose</span></span><br><span data-ttu-id="91c68-344">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-344">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="91c68-345">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-345">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="91c68-346">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="91c68-346">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="91c68-347">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="91c68-347">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="91c68-348">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="91c68-348">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="91c68-349">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="91c68-349">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="91c68-350">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="91c68-350">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="91c68-351">使用不可</span><span class="sxs-lookup"><span data-stu-id="91c68-351">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-352">Mac 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="91c68-352">Office for Mac</span></span></td>
    <td> <span data-ttu-id="91c68-353">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="91c68-353">- Mail Read</span></span><br><span data-ttu-id="91c68-354">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="91c68-354">
      - Mail Compose</span></span><br><span data-ttu-id="91c68-355">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-355">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="91c68-356">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-356">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="91c68-357">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="91c68-357">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="91c68-358">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="91c68-358">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="91c68-359">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="91c68-359">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="91c68-360">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="91c68-360">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="91c68-361">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="91c68-361">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="91c68-362">使用不可</span><span class="sxs-lookup"><span data-stu-id="91c68-362">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-363">Android 用 Office</span><span class="sxs-lookup"><span data-stu-id="91c68-363">Office for Android</span></span></td>
    <td> <span data-ttu-id="91c68-364">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="91c68-364">- Mail Read</span></span><br><span data-ttu-id="91c68-365">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-365">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="91c68-366">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-366">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="91c68-367">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="91c68-367">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="91c68-368">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="91c68-368">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="91c68-369">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="91c68-369">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="91c68-370">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="91c68-370">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="91c68-371">使用不可</span><span class="sxs-lookup"><span data-stu-id="91c68-371">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="91c68-372">Word</span><span class="sxs-lookup"><span data-stu-id="91c68-372">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="91c68-373">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="91c68-373">Platform</span></span></th>
    <th><span data-ttu-id="91c68-374">拡張点</span><span class="sxs-lookup"><span data-stu-id="91c68-374">Extension points</span></span></th>
    <th><span data-ttu-id="91c68-375">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="91c68-375">API requirement sets</span></span></th>
    <th><span data-ttu-id="91c68-376"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="91c68-376"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-377">Office Online</span><span class="sxs-lookup"><span data-stu-id="91c68-377">Office Online</span></span></td>
    <td> <span data-ttu-id="91c68-378">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-378">- Taskpane</span></span><br><span data-ttu-id="91c68-379">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-379">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="91c68-380">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-380">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="91c68-381">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="91c68-381">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="91c68-382">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="91c68-382">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="91c68-383">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-383">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="91c68-384">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-384">-BindingEvents</span></span><br><span data-ttu-id="91c68-385">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="91c68-385">
         -</span></span><br><span data-ttu-id="91c68-386">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-386">
         -DocumentEvents</span></span><br><span data-ttu-id="91c68-387">
         - File</span><span class="sxs-lookup"><span data-stu-id="91c68-387">
         - File</span></span><br><span data-ttu-id="91c68-388">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-388">
         -HtmlCoercion</span></span><br><span data-ttu-id="91c68-389">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-389">
         -ImageCoercion</span></span><br><span data-ttu-id="91c68-390">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-390">
         -MatrixBindings</span></span><br><span data-ttu-id="91c68-391">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-391">
         -MatrixCoercion</span></span><br><span data-ttu-id="91c68-392">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-392">
         -OoxmlCoercion</span></span><br><span data-ttu-id="91c68-393">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="91c68-393">
         -PdfFile</span></span><br><span data-ttu-id="91c68-394">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-394">
         - Selection</span></span><br><span data-ttu-id="91c68-395">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-395">
         - Settings</span></span><br><span data-ttu-id="91c68-396">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-396">
         -TableBindings</span></span><br><span data-ttu-id="91c68-397">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-397">
         -TableCoercion</span></span><br><span data-ttu-id="91c68-398">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-398">
         -TextBindings</span></span><br><span data-ttu-id="91c68-399">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-399">
         -TextCoercion</span></span><br><span data-ttu-id="91c68-400">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="91c68-400">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-401">Windows 用 Office 2013</span><span class="sxs-lookup"><span data-stu-id="91c68-401">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="91c68-402">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-402">- Taskpane</span></span></td>
    <td> <span data-ttu-id="91c68-403">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-403">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="91c68-404">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-404">-BindingEvents</span></span><br><span data-ttu-id="91c68-405">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="91c68-405">
         -CompressedFile</span></span><br><span data-ttu-id="91c68-406">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="91c68-406">
         -</span></span><br><span data-ttu-id="91c68-407">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-407">
         -DocumentEvents</span></span><br><span data-ttu-id="91c68-408">
         - File</span><span class="sxs-lookup"><span data-stu-id="91c68-408">
         - File</span></span><br><span data-ttu-id="91c68-409">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-409">
         -HtmlCoercion</span></span><br><span data-ttu-id="91c68-410">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-410">
         -ImageCoercion</span></span><br><span data-ttu-id="91c68-411">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-411">
         -MatrixBindings</span></span><br><span data-ttu-id="91c68-412">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-412">
         -MatrixCoercion</span></span><br><span data-ttu-id="91c68-413">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-413">
         -OoxmlCoercion</span></span><br><span data-ttu-id="91c68-414">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="91c68-414">
         -PdfFile</span></span><br><span data-ttu-id="91c68-415">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-415">
         - Selection</span></span><br><span data-ttu-id="91c68-416">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-416">
         - Settings</span></span><br><span data-ttu-id="91c68-417">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-417">
         -TableBindings</span></span><br><span data-ttu-id="91c68-418">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-418">
         -TableCoercion</span></span><br><span data-ttu-id="91c68-419">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-419">
         -TextBindings</span></span><br><span data-ttu-id="91c68-420">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-420">
         -TextCoercion</span></span><br><span data-ttu-id="91c68-421">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="91c68-421">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-422">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="91c68-422">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="91c68-423">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-423">- Taskpane</span></span><br><span data-ttu-id="91c68-424">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-424">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="91c68-425">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-425">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="91c68-426">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="91c68-426">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="91c68-427">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="91c68-427">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="91c68-428">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-428">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="91c68-429">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-429">-BindingEvents</span></span><br><span data-ttu-id="91c68-430">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="91c68-430">
         -CompressedFile</span></span><br><span data-ttu-id="91c68-431">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="91c68-431">
         -</span></span><br><span data-ttu-id="91c68-432">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-432">
         -DocumentEvents</span></span><br><span data-ttu-id="91c68-433">
         - File</span><span class="sxs-lookup"><span data-stu-id="91c68-433">
         - File</span></span><br><span data-ttu-id="91c68-434">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-434">
         -HtmlCoercion</span></span><br><span data-ttu-id="91c68-435">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-435">
         -ImageCoercion</span></span><br><span data-ttu-id="91c68-436">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-436">
         -MatrixBindings</span></span><br><span data-ttu-id="91c68-437">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-437">
         -MatrixCoercion</span></span><br><span data-ttu-id="91c68-438">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-438">
         -OoxmlCoercion</span></span><br><span data-ttu-id="91c68-439">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="91c68-439">
         -PdfFile</span></span><br><span data-ttu-id="91c68-440">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-440">
         - Selection</span></span><br><span data-ttu-id="91c68-441">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-441">
         - Settings</span></span><br><span data-ttu-id="91c68-442">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-442">
         -TableBindings</span></span><br><span data-ttu-id="91c68-443">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-443">
         -TableCoercion</span></span><br><span data-ttu-id="91c68-444">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-444">
         -TextBindings</span></span><br><span data-ttu-id="91c68-445">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-445">
         -TextCoercion</span></span><br><span data-ttu-id="91c68-446">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="91c68-446">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-447">Windows 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="91c68-447">Office for Windows</span></span></td>
    <td> <span data-ttu-id="91c68-448">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-448">- Taskpane</span></span><br><span data-ttu-id="91c68-449">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-449">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="91c68-450">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-450">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="91c68-451">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="91c68-451">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="91c68-452">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="91c68-452">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="91c68-453">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-453">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="91c68-454">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-454">-BindingEvents</span></span><br><span data-ttu-id="91c68-455">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="91c68-455">
         -CompressedFile</span></span><br><span data-ttu-id="91c68-456">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="91c68-456">
         -</span></span><br><span data-ttu-id="91c68-457">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-457">
         -DocumentEvents</span></span><br><span data-ttu-id="91c68-458">
         - File</span><span class="sxs-lookup"><span data-stu-id="91c68-458">
         - File</span></span><br><span data-ttu-id="91c68-459">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-459">
         -HtmlCoercion</span></span><br><span data-ttu-id="91c68-460">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-460">
         -ImageCoercion</span></span><br><span data-ttu-id="91c68-461">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-461">
         -MatrixBindings</span></span><br><span data-ttu-id="91c68-462">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-462">
         -MatrixCoercion</span></span><br><span data-ttu-id="91c68-463">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-463">
         -OoxmlCoercion</span></span><br><span data-ttu-id="91c68-464">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="91c68-464">
         -PdfFile</span></span><br><span data-ttu-id="91c68-465">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-465">
         - Selection</span></span><br><span data-ttu-id="91c68-466">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-466">
         - Settings</span></span><br><span data-ttu-id="91c68-467">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-467">
         -TableBindings</span></span><br><span data-ttu-id="91c68-468">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-468">
         -TableCoercion</span></span><br><span data-ttu-id="91c68-469">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-469">
         -TextBindings</span></span><br><span data-ttu-id="91c68-470">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-470">
         -TextCoercion</span></span><br><span data-ttu-id="91c68-471">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="91c68-471">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-472">iOS 用 Office</span><span class="sxs-lookup"><span data-stu-id="91c68-472">Office for iOS</span></span></td>
    <td> <span data-ttu-id="91c68-473">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-473">- Taskpane</span></span></td>
    <td> <span data-ttu-id="91c68-474">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-474">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="91c68-475">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="91c68-475">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="91c68-476">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="91c68-476">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="91c68-477">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="91c68-477">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="91c68-478">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-478">-BindingEvents</span></span><br><span data-ttu-id="91c68-479">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="91c68-479">
         -CompressedFile</span></span><br><span data-ttu-id="91c68-480">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="91c68-480">
         -</span></span><br><span data-ttu-id="91c68-481">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-481">
         -DocumentEvents</span></span><br><span data-ttu-id="91c68-482">
         - File</span><span class="sxs-lookup"><span data-stu-id="91c68-482">
         - File</span></span><br><span data-ttu-id="91c68-483">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-483">
         -HtmlCoercion</span></span><br><span data-ttu-id="91c68-484">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-484">
         -ImageCoercion</span></span><br><span data-ttu-id="91c68-485">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-485">
         -MatrixBindings</span></span><br><span data-ttu-id="91c68-486">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-486">
         -MatrixCoercion</span></span><br><span data-ttu-id="91c68-487">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-487">
         -OoxmlCoercion</span></span><br><span data-ttu-id="91c68-488">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="91c68-488">
         -PdfFile</span></span><br><span data-ttu-id="91c68-489">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-489">
         - Selection</span></span><br><span data-ttu-id="91c68-490">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-490">
         - Settings</span></span><br><span data-ttu-id="91c68-491">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-491">
         -TableBindings</span></span><br><span data-ttu-id="91c68-492">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-492">
         -TableCoercion</span></span><br><span data-ttu-id="91c68-493">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-493">
         -TextBindings</span></span><br><span data-ttu-id="91c68-494">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-494">
         -TextCoercion</span></span><br><span data-ttu-id="91c68-495">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="91c68-495">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-496">Mac 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="91c68-496">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="91c68-497">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-497">- Taskpane</span></span><br><span data-ttu-id="91c68-498">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-498">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="91c68-499">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-499">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="91c68-500">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="91c68-500">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="91c68-501">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="91c68-501">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="91c68-502">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="91c68-502">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="91c68-503">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-503">-BindingEvents</span></span><br><span data-ttu-id="91c68-504">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="91c68-504">
         -CompressedFile</span></span><br><span data-ttu-id="91c68-505">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="91c68-505">
         -</span></span><br><span data-ttu-id="91c68-506">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-506">
         -DocumentEvents</span></span><br><span data-ttu-id="91c68-507">
         - File</span><span class="sxs-lookup"><span data-stu-id="91c68-507">
         - File</span></span><br><span data-ttu-id="91c68-508">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-508">
         -HtmlCoercion</span></span><br><span data-ttu-id="91c68-509">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-509">
         -ImageCoercion</span></span><br><span data-ttu-id="91c68-510">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-510">
         -MatrixBindings</span></span><br><span data-ttu-id="91c68-511">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-511">
         -MatrixCoercion</span></span><br><span data-ttu-id="91c68-512">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-512">
         -OoxmlCoercion</span></span><br><span data-ttu-id="91c68-513">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="91c68-513">
         -PdfFile</span></span><br><span data-ttu-id="91c68-514">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-514">
         - Selection</span></span><br><span data-ttu-id="91c68-515">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-515">
         - Settings</span></span><br><span data-ttu-id="91c68-516">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-516">
         -TableBindings</span></span><br><span data-ttu-id="91c68-517">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-517">
         -TableCoercion</span></span><br><span data-ttu-id="91c68-518">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-518">
         -TextBindings</span></span><br><span data-ttu-id="91c68-519">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-519">
         -TextCoercion</span></span><br><span data-ttu-id="91c68-520">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="91c68-520">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-521">Mac 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="91c68-521">Office for Mac</span></span></td>
    <td> <span data-ttu-id="91c68-522">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-522">- Taskpane</span></span><br><span data-ttu-id="91c68-523">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-523">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="91c68-524">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-524">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="91c68-525">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="91c68-525">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="91c68-526">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="91c68-526">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="91c68-527">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="91c68-527">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="91c68-528">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-528">-BindingEvents</span></span><br><span data-ttu-id="91c68-529">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="91c68-529">
         -CompressedFile</span></span><br><span data-ttu-id="91c68-530">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="91c68-530">
         -</span></span><br><span data-ttu-id="91c68-531">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-531">
         -DocumentEvents</span></span><br><span data-ttu-id="91c68-532">
         - File</span><span class="sxs-lookup"><span data-stu-id="91c68-532">
         - File</span></span><br><span data-ttu-id="91c68-533">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-533">
         -HtmlCoercion</span></span><br><span data-ttu-id="91c68-534">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-534">
         -ImageCoercion</span></span><br><span data-ttu-id="91c68-535">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-535">
         -MatrixBindings</span></span><br><span data-ttu-id="91c68-536">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-536">
         -MatrixCoercion</span></span><br><span data-ttu-id="91c68-537">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-537">
         -OoxmlCoercion</span></span><br><span data-ttu-id="91c68-538">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="91c68-538">
         -PdfFile</span></span><br><span data-ttu-id="91c68-539">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-539">
         - Selection</span></span><br><span data-ttu-id="91c68-540">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-540">
         - Settings</span></span><br><span data-ttu-id="91c68-541">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-541">
         -TableBindings</span></span><br><span data-ttu-id="91c68-542">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-542">
         -TableCoercion</span></span><br><span data-ttu-id="91c68-543">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="91c68-543">
         -TextBindings</span></span><br><span data-ttu-id="91c68-544">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-544">
         -TextCoercion</span></span><br><span data-ttu-id="91c68-545">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="91c68-545">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="91c68-546">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="91c68-546">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="91c68-547">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="91c68-547">Platform</span></span></th>
    <th><span data-ttu-id="91c68-548">拡張点</span><span class="sxs-lookup"><span data-stu-id="91c68-548">Extension points</span></span></th>
    <th><span data-ttu-id="91c68-549">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="91c68-549">API requirement sets</span></span></th>
    <th><span data-ttu-id="91c68-550"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="91c68-550"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-551">Office Online</span><span class="sxs-lookup"><span data-stu-id="91c68-551">Office Online</span></span></td>
    <td> <span data-ttu-id="91c68-552">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="91c68-552">- Content</span></span><br><span data-ttu-id="91c68-553">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-553">
         - Taskpane</span></span><br><span data-ttu-id="91c68-554">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-554">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="91c68-555">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-555">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="91c68-556">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="91c68-556">-ActiveView</span></span><br><span data-ttu-id="91c68-557">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="91c68-557">
         -CompressedFile</span></span><br><span data-ttu-id="91c68-558">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-558">
         -DocumentEvents</span></span><br><span data-ttu-id="91c68-559">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="91c68-559">
         - File</span></span><br><span data-ttu-id="91c68-560">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-560">
         -ImageCoercion</span></span><br><span data-ttu-id="91c68-561">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="91c68-561">
         -PdfFile</span></span><br><span data-ttu-id="91c68-562">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-562">
         - Selection</span></span><br><span data-ttu-id="91c68-563">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-563">
         - Settings</span></span><br><span data-ttu-id="91c68-564">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-564">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-565">Windows 用 Office 2013</span><span class="sxs-lookup"><span data-stu-id="91c68-565">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="91c68-566">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="91c68-566">- Content</span></span><br><span data-ttu-id="91c68-567">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-567">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="91c68-568">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="91c68-568">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="91c68-569">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="91c68-569">-ActiveView</span></span><br><span data-ttu-id="91c68-570">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="91c68-570">
         -CompressedFile</span></span><br><span data-ttu-id="91c68-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-571">
         -DocumentEvents</span></span><br><span data-ttu-id="91c68-572">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="91c68-572">
         - File</span></span><br><span data-ttu-id="91c68-573">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-573">
         -ImageCoercion</span></span><br><span data-ttu-id="91c68-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="91c68-574">
         -PdfFile</span></span><br><span data-ttu-id="91c68-575">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-575">
         - Selection</span></span><br><span data-ttu-id="91c68-576">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-576">
         - Settings</span></span><br><span data-ttu-id="91c68-577">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-577">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-578">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="91c68-578">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="91c68-579">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="91c68-579">- Content</span></span><br><span data-ttu-id="91c68-580">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-580">
         - Taskpane</span></span><br><span data-ttu-id="91c68-581">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-581">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="91c68-582">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-582">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="91c68-583">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="91c68-583">-ActiveView</span></span><br><span data-ttu-id="91c68-584">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="91c68-584">
         -CompressedFile</span></span><br><span data-ttu-id="91c68-585">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-585">
         -DocumentEvents</span></span><br><span data-ttu-id="91c68-586">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="91c68-586">
         - File</span></span><br><span data-ttu-id="91c68-587">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-587">
         -ImageCoercion</span></span><br><span data-ttu-id="91c68-588">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="91c68-588">
         -PdfFile</span></span><br><span data-ttu-id="91c68-589">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-589">
         - Selection</span></span><br><span data-ttu-id="91c68-590">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-590">
         - Settings</span></span><br><span data-ttu-id="91c68-591">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-591">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-592">Windows 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="91c68-592">Office for Windows</span></span></td>
    <td> <span data-ttu-id="91c68-593">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="91c68-593">- Content</span></span><br><span data-ttu-id="91c68-594">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-594">
         - Taskpane</span></span><br><span data-ttu-id="91c68-595">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-595">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="91c68-596">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-596">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="91c68-597">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="91c68-597">-ActiveView</span></span><br><span data-ttu-id="91c68-598">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="91c68-598">
         -CompressedFile</span></span><br><span data-ttu-id="91c68-599">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-599">
         -DocumentEvents</span></span><br><span data-ttu-id="91c68-600">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="91c68-600">
         - File</span></span><br><span data-ttu-id="91c68-601">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-601">
         -ImageCoercion</span></span><br><span data-ttu-id="91c68-602">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="91c68-602">
         -PdfFile</span></span><br><span data-ttu-id="91c68-603">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-603">
         - Selection</span></span><br><span data-ttu-id="91c68-604">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-604">
         - Settings</span></span><br><span data-ttu-id="91c68-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-605">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-606">iOS 用 Office</span><span class="sxs-lookup"><span data-stu-id="91c68-606">Office for iOS</span></span></td>
    <td> <span data-ttu-id="91c68-607">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="91c68-607">- Content</span></span><br><span data-ttu-id="91c68-608">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-608">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="91c68-609">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-609">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="91c68-610">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="91c68-610">-ActiveView</span></span><br><span data-ttu-id="91c68-611">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="91c68-611">
         -CompressedFile</span></span><br><span data-ttu-id="91c68-612">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-612">
         -DocumentEvents</span></span><br><span data-ttu-id="91c68-613">
         - File</span><span class="sxs-lookup"><span data-stu-id="91c68-613">
         - File</span></span><br><span data-ttu-id="91c68-614">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="91c68-614">
         -PdfFile</span></span><br><span data-ttu-id="91c68-615">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-615">
         - Selection</span></span><br><span data-ttu-id="91c68-616">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-616">
         - Settings</span></span><br><span data-ttu-id="91c68-617">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-617">
         -TextCoercion</span></span><br><span data-ttu-id="91c68-618">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-618">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-619">Mac 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="91c68-619">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="91c68-620">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="91c68-620">- Content</span></span><br><span data-ttu-id="91c68-621">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-621">
         - Taskpane</span></span><br><span data-ttu-id="91c68-622">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-622">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="91c68-623">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-623">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="91c68-624">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="91c68-624">-ActiveView</span></span><br><span data-ttu-id="91c68-625">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="91c68-625">
         -CompressedFile</span></span><br><span data-ttu-id="91c68-626">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-626">
         -DocumentEvents</span></span><br><span data-ttu-id="91c68-627">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="91c68-627">
         - File</span></span><br><span data-ttu-id="91c68-628">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-628">
         -ImageCoercion</span></span><br><span data-ttu-id="91c68-629">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="91c68-629">
         -PdfFile</span></span><br><span data-ttu-id="91c68-630">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-630">
         - Selection</span></span><br><span data-ttu-id="91c68-631">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-631">
         - Settings</span></span><br><span data-ttu-id="91c68-632">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-632">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-633">Mac 用 Office 2019</span><span class="sxs-lookup"><span data-stu-id="91c68-633">Office for Mac</span></span></td>
    <td> <span data-ttu-id="91c68-634">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="91c68-634">- Content</span></span><br><span data-ttu-id="91c68-635">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-635">
         - Taskpane</span></span><br><span data-ttu-id="91c68-636">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-636">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="91c68-637">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-637">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="91c68-638">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="91c68-638">-ActiveView</span></span><br><span data-ttu-id="91c68-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="91c68-639">
         -CompressedFile</span></span><br><span data-ttu-id="91c68-640">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-640">
         -DocumentEvents</span></span><br><span data-ttu-id="91c68-641">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="91c68-641">
         - File</span></span><br><span data-ttu-id="91c68-642">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-642">
         -ImageCoercion</span></span><br><span data-ttu-id="91c68-643">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="91c68-643">
         -PdfFile</span></span><br><span data-ttu-id="91c68-644">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="91c68-644">
         - Selection</span></span><br><span data-ttu-id="91c68-645">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-645">
         - Settings</span></span><br><span data-ttu-id="91c68-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-646">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="91c68-647">OneNote</span><span class="sxs-lookup"><span data-stu-id="91c68-647">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="91c68-648">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="91c68-648">Platform</span></span></th>
    <th><span data-ttu-id="91c68-649">拡張点</span><span class="sxs-lookup"><span data-stu-id="91c68-649">Extension points</span></span></th>
    <th><span data-ttu-id="91c68-650">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="91c68-650">API requirement sets</span></span></th>
    <th><span data-ttu-id="91c68-651"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="91c68-651"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="91c68-652">Office Online</span><span class="sxs-lookup"><span data-stu-id="91c68-652">Office Online</span></span></td>
    <td> <span data-ttu-id="91c68-653">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="91c68-653">- Content</span></span><br><span data-ttu-id="91c68-654">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="91c68-654">
         - Taskpane</span></span><br><span data-ttu-id="91c68-655">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="91c68-655">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="91c68-656">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-656">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="91c68-657">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="91c68-657">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="91c68-658">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="91c68-658">-DocumentEvents</span></span><br><span data-ttu-id="91c68-659">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-659">
         -HtmlCoercion</span></span><br><span data-ttu-id="91c68-660">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-660">
         -ImageCoercion</span></span><br><span data-ttu-id="91c68-661">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="91c68-661">
         - Settings</span></span><br><span data-ttu-id="91c68-662">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="91c68-662">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="91c68-663">関連項目</span><span class="sxs-lookup"><span data-stu-id="91c68-663">See also</span></span>

- [<span data-ttu-id="91c68-664">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="91c68-664">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="91c68-665">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="91c68-665">Common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="91c68-666">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="91c68-666">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="91c68-667">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="91c68-667">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office)
