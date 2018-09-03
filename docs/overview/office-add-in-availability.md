---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、Word、Outlook、PowerPoint、および OneNote のサポートされる要件セット。
ms.date: 08/30/2018
ms.openlocfilehash: 06fb073693bd8adca7d196f4361699ac3f54cee1
ms.sourcegitcommit: 78b28ae88d53bfef3134c09cc4336a5a8722c70b
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/01/2018
ms.locfileid: "23797302"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="e0c55-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e0c55-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="e0c55-104">Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。</span><span class="sxs-lookup"><span data-stu-id="e0c55-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="e0c55-105">次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API の要件セットが含まれています。</span><span class="sxs-lookup"><span data-stu-id="e0c55-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="e0c55-106">表のセルにアスタリスク ( \* ) が含まれる場合は、準備中です。</span><span class="sxs-lookup"><span data-stu-id="e0c55-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="e0c55-107">Project または Access の要件セットについては、「[Office の共有要件セット](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e0c55-107">For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="e0c55-p103">MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。このバージョンには、ExcelApi 1.1、WordApi 1.1、および共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="e0c55-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="e0c55-110">Excel</span><span class="sxs-lookup"><span data-stu-id="e0c55-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="e0c55-111">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e0c55-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="e0c55-112">拡張点</span><span class="sxs-lookup"><span data-stu-id="e0c55-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="e0c55-113">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="e0c55-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="e0c55-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e0c55-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="e0c55-115">Office Online</span></span></td>
    <td> <span data-ttu-id="e0c55-116">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e0c55-116">- Taskpane</span></span><br><span data-ttu-id="e0c55-117">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e0c55-117">
        - Content</span></span><br><span data-ttu-id="e0c55-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="e0c55-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="e0c55-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e0c55-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e0c55-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e0c55-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e0c55-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e0c55-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e0c55-125">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="e0c55-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e0c55-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-127">
        -BindingEvents</span></span><br><span data-ttu-id="e0c55-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-128">
        -CompressedFile</span></span><br><span data-ttu-id="e0c55-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-129">
        -DocumentEvents</span></span><br><span data-ttu-id="e0c55-130">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="e0c55-130">
        - File</span></span><br><span data-ttu-id="e0c55-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-131">
        -MatrixBindings</span></span><br><span data-ttu-id="e0c55-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-132">
        -MatrixCoercion</span></span><br><span data-ttu-id="e0c55-133">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="e0c55-133">
        - Selection</span></span><br><span data-ttu-id="e0c55-134">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="e0c55-134">
        - Settings</span></span><br><span data-ttu-id="e0c55-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-135">
        -TableBindings</span></span><br><span data-ttu-id="e0c55-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-136">
        -TableCoercion</span></span><br><span data-ttu-id="e0c55-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-137">
        -TextBindings</span></span><br><span data-ttu-id="e0c55-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-138">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-139">Windows 用 Office 2013</span><span class="sxs-lookup"><span data-stu-id="e0c55-139">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="e0c55-140">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e0c55-140">
        - Taskpane</span></span><br><span data-ttu-id="e0c55-141">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e0c55-141">
        - Content</span></span></td>
    <td>  <span data-ttu-id="e0c55-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e0c55-143">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-143">
        -BindingEvents</span></span><br><span data-ttu-id="e0c55-144">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-144">
        -CompressedFile</span></span><br><span data-ttu-id="e0c55-145">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-145">
        -DocumentEvents</span></span><br><span data-ttu-id="e0c55-146">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="e0c55-146">
        - File</span></span><br><span data-ttu-id="e0c55-147">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-147">
        -ImageCoercion</span></span><br><span data-ttu-id="e0c55-148">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-148">
        -MatrixBindings</span></span><br><span data-ttu-id="e0c55-149">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-149">
        -MatrixCoercion</span></span><br><span data-ttu-id="e0c55-150">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="e0c55-150">
        - Selection</span></span><br><span data-ttu-id="e0c55-151">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="e0c55-151">
        - Settings</span></span><br><span data-ttu-id="e0c55-152">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-152">
        -TableBindings</span></span><br><span data-ttu-id="e0c55-153">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-153">
        -TableCoercion</span></span><br><span data-ttu-id="e0c55-154">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-154">
        -TextBindings</span></span><br><span data-ttu-id="e0c55-155">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-155">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-156">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="e0c55-156">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="e0c55-157">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e0c55-157">- Taskpane</span></span><br><span data-ttu-id="e0c55-158">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e0c55-158">
        - Content</span></span><br><span data-ttu-id="e0c55-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e0c55-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e0c55-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e0c55-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e0c55-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e0c55-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e0c55-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e0c55-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-166">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="e0c55-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e0c55-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-168">-BindingEvents</span></span><br><span data-ttu-id="e0c55-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-169">
        -CompressedFile</span></span><br><span data-ttu-id="e0c55-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-170">
        -DocumentEvents</span></span><br><span data-ttu-id="e0c55-171">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="e0c55-171">
        - File</span></span><br><span data-ttu-id="e0c55-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-172">
        -ImageCoercion</span></span><br><span data-ttu-id="e0c55-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-173">
        -MatrixBindings</span></span><br><span data-ttu-id="e0c55-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-174">
        -MatrixCoercion</span></span><br><span data-ttu-id="e0c55-175">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="e0c55-175">
        - Selection</span></span><br><span data-ttu-id="e0c55-176">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="e0c55-176">
        - Settings</span></span><br><span data-ttu-id="e0c55-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-177">
        -TableBindings</span></span><br><span data-ttu-id="e0c55-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-178">
        -TableCoercion</span></span><br><span data-ttu-id="e0c55-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-179">
        -TextBindings</span></span><br><span data-ttu-id="e0c55-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-181">iOS 用 Office</span><span class="sxs-lookup"><span data-stu-id="e0c55-181">Office for iOS</span></span></td>
    <td><span data-ttu-id="e0c55-182">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e0c55-182">- Taskpane</span></span><br><span data-ttu-id="e0c55-183">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e0c55-183">
        - Content</span></span></td>
    <td><span data-ttu-id="e0c55-184">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-184">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e0c55-185">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-185">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e0c55-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e0c55-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e0c55-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e0c55-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e0c55-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-190">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="e0c55-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e0c55-192">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-192">-BindingEvents</span></span><br><span data-ttu-id="e0c55-193">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-193">
        -CompressedFile</span></span><br><span data-ttu-id="e0c55-194">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-194">
        -DocumentEvents</span></span><br><span data-ttu-id="e0c55-195">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="e0c55-195">
        - File</span></span><br><span data-ttu-id="e0c55-196">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-196">
        -ImageCoercion</span></span><br><span data-ttu-id="e0c55-197">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-197">
        -MatrixBindings</span></span><br><span data-ttu-id="e0c55-198">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-198">
        -MatrixCoercion</span></span><br><span data-ttu-id="e0c55-199">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="e0c55-199">
        - Selection</span></span><br><span data-ttu-id="e0c55-200">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="e0c55-200">
        - Settings</span></span><br><span data-ttu-id="e0c55-201">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-201">
        -TableBindings</span></span><br><span data-ttu-id="e0c55-202">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-202">
        -TableCoercion</span></span><br><span data-ttu-id="e0c55-203">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-203">
        -TextBindings</span></span><br><span data-ttu-id="e0c55-204">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-204">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-205">Mac 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="e0c55-205">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="e0c55-206">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e0c55-206">- Taskpane</span></span><br><span data-ttu-id="e0c55-207">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e0c55-207">
        - Content</span></span><br><span data-ttu-id="e0c55-208">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-208">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e0c55-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e0c55-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e0c55-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e0c55-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e0c55-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e0c55-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e0c55-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-215">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="e0c55-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e0c55-217">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-217">-BindingEvents</span></span><br><span data-ttu-id="e0c55-218">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-218">
        -CompressedFile</span></span><br><span data-ttu-id="e0c55-219">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-219">
        -DocumentEvents</span></span><br><span data-ttu-id="e0c55-220">
        - ファイル</span><span class="sxs-lookup"><span data-stu-id="e0c55-220">
        - File</span></span><br><span data-ttu-id="e0c55-221">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-221">
        -ImageCoercion</span></span><br><span data-ttu-id="e0c55-222">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-222">
        -MatrixBindings</span></span><br><span data-ttu-id="e0c55-223">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-223">
        -MatrixCoercion</span></span><br><span data-ttu-id="e0c55-224">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-224">
        -PdfFile</span></span><br><span data-ttu-id="e0c55-225">
        - 選択</span><span class="sxs-lookup"><span data-stu-id="e0c55-225">
        - Selection</span></span><br><span data-ttu-id="e0c55-226">
        - 設定</span><span class="sxs-lookup"><span data-stu-id="e0c55-226">
        - Settings</span></span><br><span data-ttu-id="e0c55-227">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-227">
        -TableBindings</span></span><br><span data-ttu-id="e0c55-228">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-228">
        -TableCoercion</span></span><br><span data-ttu-id="e0c55-229">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-229">
        -TextBindings</span></span><br><span data-ttu-id="e0c55-230">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-230">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="e0c55-231">Outlook</span><span class="sxs-lookup"><span data-stu-id="e0c55-231">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e0c55-232">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e0c55-232">Platform</span></span></th>
    <th><span data-ttu-id="e0c55-233">拡張点</span><span class="sxs-lookup"><span data-stu-id="e0c55-233">Extension points</span></span></th>
    <th><span data-ttu-id="e0c55-234">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="e0c55-234">API requirement sets</span></span></th>
    <th><span data-ttu-id="e0c55-235"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e0c55-235"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-236">Office Online</span><span class="sxs-lookup"><span data-stu-id="e0c55-236">Office Online</span></span></td>
    <td> <span data-ttu-id="e0c55-237">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e0c55-237">- Mail Read</span></span><br><span data-ttu-id="e0c55-238">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="e0c55-238">
      - Mail Compose</span></span><br><span data-ttu-id="e0c55-239">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-239">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0c55-240">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-240">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e0c55-241">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-241">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e0c55-242">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-242">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e0c55-243">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-243">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e0c55-244">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-244">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e0c55-245">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-245">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e0c55-246">使用不可</span><span class="sxs-lookup"><span data-stu-id="e0c55-246">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-247">Windows 用 Office 2013</span><span class="sxs-lookup"><span data-stu-id="e0c55-247">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e0c55-248">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e0c55-248">- Mail Read</span></span><br><span data-ttu-id="e0c55-249">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="e0c55-249">
      - Mail Compose</span></span><br><span data-ttu-id="e0c55-250">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-250">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0c55-251">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-251">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e0c55-252">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-252">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e0c55-253">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-253">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e0c55-254">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-254">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="e0c55-255">使用不可</span><span class="sxs-lookup"><span data-stu-id="e0c55-255">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-256">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="e0c55-256">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e0c55-257">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e0c55-257">- Mail Read</span></span><br><span data-ttu-id="e0c55-258">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="e0c55-258">
      - Mail Compose</span></span><br><span data-ttu-id="e0c55-259">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-259">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="e0c55-260">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="e0c55-260">
      - Modules</span></span></td>
    <td> <span data-ttu-id="e0c55-261">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-261">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e0c55-262">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-262">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e0c55-263">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-263">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e0c55-264">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-264">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e0c55-265">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-265">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e0c55-266">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-266">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e0c55-267">使用不可</span><span class="sxs-lookup"><span data-stu-id="e0c55-267">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-268">iOS 版 Office</span><span class="sxs-lookup"><span data-stu-id="e0c55-268">Office for iOS</span></span></td>
    <td> <span data-ttu-id="e0c55-269">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e0c55-269">- Mail Read</span></span><br><span data-ttu-id="e0c55-270">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-270">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0c55-271">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-271">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e0c55-272">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-272">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e0c55-273">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-273">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e0c55-274">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-274">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e0c55-275">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-275">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="e0c55-276">使用不可</span><span class="sxs-lookup"><span data-stu-id="e0c55-276">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-277">Mac 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="e0c55-277">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e0c55-278">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e0c55-278">- Mail Read</span></span><br><span data-ttu-id="e0c55-279">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="e0c55-279">
      - Mail Compose</span></span><br><span data-ttu-id="e0c55-280">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-280">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0c55-281">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-281">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e0c55-282">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-282">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e0c55-283">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-283">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e0c55-284">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-284">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e0c55-285">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-285">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e0c55-286">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-286">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e0c55-287">使用不可</span><span class="sxs-lookup"><span data-stu-id="e0c55-287">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-288">Android 版 Office</span><span class="sxs-lookup"><span data-stu-id="e0c55-288">Office for Android</span></span></td>
    <td> <span data-ttu-id="e0c55-289">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e0c55-289">- Mail Read</span></span><br><span data-ttu-id="e0c55-290">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-290">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0c55-291">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-291">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e0c55-292">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-292">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e0c55-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e0c55-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e0c55-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="e0c55-296">使用不可</span><span class="sxs-lookup"><span data-stu-id="e0c55-296">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="e0c55-297">Word</span><span class="sxs-lookup"><span data-stu-id="e0c55-297">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e0c55-298">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e0c55-298">Platform</span></span></th>
    <th><span data-ttu-id="e0c55-299">拡張点</span><span class="sxs-lookup"><span data-stu-id="e0c55-299">Extension points</span></span></th>
    <th><span data-ttu-id="e0c55-300">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="e0c55-300">API requirement sets</span></span></th>
    <th><span data-ttu-id="e0c55-301"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e0c55-301"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-302">Office Online</span><span class="sxs-lookup"><span data-stu-id="e0c55-302">Office Online</span></span></td>
    <td> <span data-ttu-id="e0c55-303">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e0c55-303">- Taskpane</span></span><br><span data-ttu-id="e0c55-304">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-304">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0c55-305">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-305">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e0c55-306">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-306">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e0c55-307">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-307">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e0c55-308">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-308">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e0c55-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-309">-BindingEvents</span></span><br><span data-ttu-id="e0c55-310">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="e0c55-310">
         -</span></span><br><span data-ttu-id="e0c55-311">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-311">
         -DocumentEvents</span></span><br><span data-ttu-id="e0c55-312">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="e0c55-312">
         - File</span></span><br><span data-ttu-id="e0c55-313">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-313">
         -HtmlCoercion</span></span><br><span data-ttu-id="e0c55-314">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-314">
         -ImageCoercion</span></span><br><span data-ttu-id="e0c55-315">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-315">
         -MatrixBindings</span></span><br><span data-ttu-id="e0c55-316">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-316">
         -MatrixCoercion</span></span><br><span data-ttu-id="e0c55-317">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-317">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e0c55-318">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-318">
         -PdfFile</span></span><br><span data-ttu-id="e0c55-319">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="e0c55-319">
         - Selection</span></span><br><span data-ttu-id="e0c55-320">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="e0c55-320">
         - Settings</span></span><br><span data-ttu-id="e0c55-321">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-321">
         -TableBindings</span></span><br><span data-ttu-id="e0c55-322">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-322">
         -TableCoercion</span></span><br><span data-ttu-id="e0c55-323">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-323">
         -TextBindings</span></span><br><span data-ttu-id="e0c55-324">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-324">
         -TextCoercion</span></span><br><span data-ttu-id="e0c55-325">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-325">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-326">Windows 用 Office 2013</span><span class="sxs-lookup"><span data-stu-id="e0c55-326">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e0c55-327">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e0c55-327">- Taskpane</span></span></td>
    <td> <span data-ttu-id="e0c55-328">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-328">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e0c55-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-329">-BindingEvents</span></span><br><span data-ttu-id="e0c55-330">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-330">
         -CompressedFile</span></span><br><span data-ttu-id="e0c55-331">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="e0c55-331">
         -</span></span><br><span data-ttu-id="e0c55-332">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-332">
         -DocumentEvents</span></span><br><span data-ttu-id="e0c55-333">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="e0c55-333">
         - File</span></span><br><span data-ttu-id="e0c55-334">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-334">
         -HtmlCoercion</span></span><br><span data-ttu-id="e0c55-335">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-335">
         -ImageCoercion</span></span><br><span data-ttu-id="e0c55-336">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-336">
         -MatrixBindings</span></span><br><span data-ttu-id="e0c55-337">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-337">
         -MatrixCoercion</span></span><br><span data-ttu-id="e0c55-338">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-338">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e0c55-339">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-339">
         -PdfFile</span></span><br><span data-ttu-id="e0c55-340">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="e0c55-340">
         - Selection</span></span><br><span data-ttu-id="e0c55-341">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="e0c55-341">
         - Settings</span></span><br><span data-ttu-id="e0c55-342">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-342">
         -TableBindings</span></span><br><span data-ttu-id="e0c55-343">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-343">
         -TableCoercion</span></span><br><span data-ttu-id="e0c55-344">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-344">
         -TextBindings</span></span><br><span data-ttu-id="e0c55-345">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-345">
         -TextCoercion</span></span><br><span data-ttu-id="e0c55-346">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-346">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-347">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="e0c55-347">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e0c55-348">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e0c55-348">- Taskpane</span></span><br><span data-ttu-id="e0c55-349">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-349">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0c55-350">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-350">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e0c55-351">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-351">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e0c55-352">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-352">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e0c55-353">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-353">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e0c55-354">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-354">-BindingEvents</span></span><br><span data-ttu-id="e0c55-355">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-355">
         -CompressedFile</span></span><br><span data-ttu-id="e0c55-356">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="e0c55-356">
         -</span></span><br><span data-ttu-id="e0c55-357">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-357">
         -DocumentEvents</span></span><br><span data-ttu-id="e0c55-358">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="e0c55-358">
         - File</span></span><br><span data-ttu-id="e0c55-359">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-359">
         -HtmlCoercion</span></span><br><span data-ttu-id="e0c55-360">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-360">
         -ImageCoercion</span></span><br><span data-ttu-id="e0c55-361">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-361">
         -MatrixBindings</span></span><br><span data-ttu-id="e0c55-362">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-362">
         -MatrixCoercion</span></span><br><span data-ttu-id="e0c55-363">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-363">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e0c55-364">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-364">
         -PdfFile</span></span><br><span data-ttu-id="e0c55-365">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="e0c55-365">
         - Selection</span></span><br><span data-ttu-id="e0c55-366">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="e0c55-366">
         - Settings</span></span><br><span data-ttu-id="e0c55-367">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-367">
         -TableBindings</span></span><br><span data-ttu-id="e0c55-368">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-368">
         -TableCoercion</span></span><br><span data-ttu-id="e0c55-369">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-369">
         -TextBindings</span></span><br><span data-ttu-id="e0c55-370">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-370">
         -TextCoercion</span></span><br><span data-ttu-id="e0c55-371">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-371">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-372">iOS 用 Office</span><span class="sxs-lookup"><span data-stu-id="e0c55-372">Office for iOS</span></span></td>
    <td> <span data-ttu-id="e0c55-373">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e0c55-373">- Taskpane</span></span></td>
    <td> <span data-ttu-id="e0c55-374">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-374">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e0c55-375">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-375">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e0c55-376">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-376">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e0c55-377">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e0c55-377">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e0c55-378">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-378">-BindingEvents</span></span><br><span data-ttu-id="e0c55-379">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-379">
         -CompressedFile</span></span><br><span data-ttu-id="e0c55-380">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="e0c55-380">
         -</span></span><br><span data-ttu-id="e0c55-381">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-381">
         -DocumentEvents</span></span><br><span data-ttu-id="e0c55-382">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="e0c55-382">
         - File</span></span><br><span data-ttu-id="e0c55-383">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-383">
         -HtmlCoercion</span></span><br><span data-ttu-id="e0c55-384">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-384">
         -ImageCoercion</span></span><br><span data-ttu-id="e0c55-385">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-385">
         -MatrixBindings</span></span><br><span data-ttu-id="e0c55-386">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-386">
         -MatrixCoercion</span></span><br><span data-ttu-id="e0c55-387">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-387">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e0c55-388">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-388">
         -PdfFile</span></span><br><span data-ttu-id="e0c55-389">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="e0c55-389">
         - Selection</span></span><br><span data-ttu-id="e0c55-390">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="e0c55-390">
         - Settings</span></span><br><span data-ttu-id="e0c55-391">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-391">
         -TableBindings</span></span><br><span data-ttu-id="e0c55-392">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-392">
         -TableCoercion</span></span><br><span data-ttu-id="e0c55-393">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-393">
         -TextBindings</span></span><br><span data-ttu-id="e0c55-394">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-394">
         -TextCoercion</span></span><br><span data-ttu-id="e0c55-395">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-395">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-396">Mac 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="e0c55-396">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e0c55-397">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e0c55-397">- Taskpane</span></span><br><span data-ttu-id="e0c55-398">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-398">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0c55-399">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-399">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e0c55-400">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-400">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e0c55-401">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-401">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e0c55-402">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e0c55-402">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e0c55-403">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-403">-BindingEvents</span></span><br><span data-ttu-id="e0c55-404">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-404">
         -CompressedFile</span></span><br><span data-ttu-id="e0c55-405">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="e0c55-405">
         -</span></span><br><span data-ttu-id="e0c55-406">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-406">
         -DocumentEvents</span></span><br><span data-ttu-id="e0c55-407">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="e0c55-407">
         - File</span></span><br><span data-ttu-id="e0c55-408">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-408">
         -HtmlCoercion</span></span><br><span data-ttu-id="e0c55-409">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-409">
         -ImageCoercion</span></span><br><span data-ttu-id="e0c55-410">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-410">
         -MatrixBindings</span></span><br><span data-ttu-id="e0c55-411">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-411">
         -MatrixCoercion</span></span><br><span data-ttu-id="e0c55-412">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-412">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e0c55-413">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-413">
         -PdfFile</span></span><br><span data-ttu-id="e0c55-414">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="e0c55-414">
         - Selection</span></span><br><span data-ttu-id="e0c55-415">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="e0c55-415">
         - Settings</span></span><br><span data-ttu-id="e0c55-416">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-416">
         -TableBindings</span></span><br><span data-ttu-id="e0c55-417">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-417">
         -TableCoercion</span></span><br><span data-ttu-id="e0c55-418">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e0c55-418">
         -TextBindings</span></span><br><span data-ttu-id="e0c55-419">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-419">
         -TextCoercion</span></span><br><span data-ttu-id="e0c55-420">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-420">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="e0c55-421">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="e0c55-421">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e0c55-422">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e0c55-422">Platform</span></span></th>
    <th><span data-ttu-id="e0c55-423">拡張点</span><span class="sxs-lookup"><span data-stu-id="e0c55-423">Extension points</span></span></th>
    <th><span data-ttu-id="e0c55-424">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="e0c55-424">API requirement sets</span></span></th>
    <th><span data-ttu-id="e0c55-425"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e0c55-425"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-426">Office Online</span><span class="sxs-lookup"><span data-stu-id="e0c55-426">Office Online</span></span></td>
    <td> <span data-ttu-id="e0c55-427">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e0c55-427">- Content</span></span><br><span data-ttu-id="e0c55-428">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e0c55-428">
         - Taskpane</span></span><br><span data-ttu-id="e0c55-429">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-429">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0c55-430">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-430">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e0c55-431">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e0c55-431">-ActiveView</span></span><br><span data-ttu-id="e0c55-432">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-432">
         -CompressedFile</span></span><br><span data-ttu-id="e0c55-433">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-433">
         -DocumentEvents</span></span><br><span data-ttu-id="e0c55-434">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="e0c55-434">
         - File</span></span><br><span data-ttu-id="e0c55-435">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-435">
         -ImageCoercion</span></span><br><span data-ttu-id="e0c55-436">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-436">
         -PdfFile</span></span><br><span data-ttu-id="e0c55-437">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="e0c55-437">
         - Selection</span></span><br><span data-ttu-id="e0c55-438">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="e0c55-438">
         - Settings</span></span><br><span data-ttu-id="e0c55-439">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-439">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-440">Windows 用 Office 2013</span><span class="sxs-lookup"><span data-stu-id="e0c55-440">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e0c55-441">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e0c55-441">- Content</span></span><br><span data-ttu-id="e0c55-442">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e0c55-442">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="e0c55-443">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e0c55-443">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e0c55-444">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e0c55-444">-ActiveView</span></span><br><span data-ttu-id="e0c55-445">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-445">
         -CompressedFile</span></span><br><span data-ttu-id="e0c55-446">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-446">
         -DocumentEvents</span></span><br><span data-ttu-id="e0c55-447">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="e0c55-447">
         - File</span></span><br><span data-ttu-id="e0c55-448">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-448">
         -ImageCoercion</span></span><br><span data-ttu-id="e0c55-449">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-449">
         -PdfFile</span></span><br><span data-ttu-id="e0c55-450">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="e0c55-450">
         - Selection</span></span><br><span data-ttu-id="e0c55-451">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="e0c55-451">
         - Settings</span></span><br><span data-ttu-id="e0c55-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-452">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-453">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="e0c55-453">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e0c55-454">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e0c55-454">- Content</span></span><br><span data-ttu-id="e0c55-455">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e0c55-455">
         - Taskpane</span></span><br><span data-ttu-id="e0c55-456">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-456">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0c55-457">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-457">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e0c55-458">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e0c55-458">-ActiveView</span></span><br><span data-ttu-id="e0c55-459">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-459">
         -CompressedFile</span></span><br><span data-ttu-id="e0c55-460">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-460">
         -DocumentEvents</span></span><br><span data-ttu-id="e0c55-461">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="e0c55-461">
         - File</span></span><br><span data-ttu-id="e0c55-462">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-462">
         -ImageCoercion</span></span><br><span data-ttu-id="e0c55-463">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-463">
         -PdfFile</span></span><br><span data-ttu-id="e0c55-464">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="e0c55-464">
         - Selection</span></span><br><span data-ttu-id="e0c55-465">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="e0c55-465">
         - Settings</span></span><br><span data-ttu-id="e0c55-466">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-466">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-467">iOS 用 Office</span><span class="sxs-lookup"><span data-stu-id="e0c55-467">Office for iOS</span></span></td>
    <td> <span data-ttu-id="e0c55-468">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e0c55-468">- Content</span></span><br><span data-ttu-id="e0c55-469">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e0c55-469">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="e0c55-470">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-470">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="e0c55-471">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e0c55-471">-ActiveView</span></span><br><span data-ttu-id="e0c55-472">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-472">
         -CompressedFile</span></span><br><span data-ttu-id="e0c55-473">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-473">
         -DocumentEvents</span></span><br><span data-ttu-id="e0c55-474">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="e0c55-474">
         - File</span></span><br><span data-ttu-id="e0c55-475">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-475">
         -PdfFile</span></span><br><span data-ttu-id="e0c55-476">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="e0c55-476">
         - Selection</span></span><br><span data-ttu-id="e0c55-477">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="e0c55-477">
         - Settings</span></span><br><span data-ttu-id="e0c55-478">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-478">
         -TextCoercion</span></span><br><span data-ttu-id="e0c55-479">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-479">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-480">Mac 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="e0c55-480">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e0c55-481">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e0c55-481">- Content</span></span><br><span data-ttu-id="e0c55-482">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e0c55-482">
         - Taskpane</span></span><br><span data-ttu-id="e0c55-483">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-483">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0c55-484">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-484">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e0c55-485">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e0c55-485">-ActiveView</span></span><br><span data-ttu-id="e0c55-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-486">
         -CompressedFile</span></span><br><span data-ttu-id="e0c55-487">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-487">
         -DocumentEvents</span></span><br><span data-ttu-id="e0c55-488">
         - ファイル</span><span class="sxs-lookup"><span data-stu-id="e0c55-488">
         - File</span></span><br><span data-ttu-id="e0c55-489">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-489">
         -ImageCoercion</span></span><br><span data-ttu-id="e0c55-490">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0c55-490">
         -PdfFile</span></span><br><span data-ttu-id="e0c55-491">
         - 選択</span><span class="sxs-lookup"><span data-stu-id="e0c55-491">
         - Selection</span></span><br><span data-ttu-id="e0c55-492">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="e0c55-492">
         - Settings</span></span><br><span data-ttu-id="e0c55-493">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-493">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="e0c55-494">OneNote</span><span class="sxs-lookup"><span data-stu-id="e0c55-494">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e0c55-495">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e0c55-495">Platform</span></span></th>
    <th><span data-ttu-id="e0c55-496">拡張点</span><span class="sxs-lookup"><span data-stu-id="e0c55-496">Extension points</span></span></th>
    <th><span data-ttu-id="e0c55-497">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="e0c55-497">API requirement sets</span></span></th>
    <th><span data-ttu-id="e0c55-498"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e0c55-498"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="e0c55-499">Office Online</span><span class="sxs-lookup"><span data-stu-id="e0c55-499">Office Online</span></span></td>
    <td> <span data-ttu-id="e0c55-500">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e0c55-500">- Content</span></span><br><span data-ttu-id="e0c55-501">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e0c55-501">
         - Taskpane</span></span><br><span data-ttu-id="e0c55-502">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-502">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0c55-503">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-503">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="e0c55-504">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0c55-504">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e0c55-505">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0c55-505">-DocumentEvents</span></span><br><span data-ttu-id="e0c55-506">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-506">
         -HtmlCoercion</span></span><br><span data-ttu-id="e0c55-507">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-507">
         -ImageCoercion</span></span><br><span data-ttu-id="e0c55-508">
         - 設定</span><span class="sxs-lookup"><span data-stu-id="e0c55-508">
         - Settings</span></span><br><span data-ttu-id="e0c55-509">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0c55-509">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="e0c55-510">関連項目</span><span class="sxs-lookup"><span data-stu-id="e0c55-510">See also</span></span>

- [<span data-ttu-id="e0c55-511">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="e0c55-511">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="e0c55-512">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="e0c55-512">Common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="e0c55-513">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="e0c55-513">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="e0c55-514">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="e0c55-514">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office)
