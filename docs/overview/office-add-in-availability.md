---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、Word、Outlook、PowerPoint、および OneNote のサポートされる要件セット。
ms.date: 07/31/2018
ms.openlocfilehash: 084029c0a5b70b73eaa0b3fcc180f4a813fb8b72
ms.sourcegitcommit: bc68b4cf811b45e8b8d1cbd7c8d2867359ab671b
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/02/2018
ms.locfileid: "21703911"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="4cec2-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4cec2-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="4cec2-104">Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。</span><span class="sxs-lookup"><span data-stu-id="4cec2-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="4cec2-105">次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API の要件セットが含まれています。</span><span class="sxs-lookup"><span data-stu-id="4cec2-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span> 

<span data-ttu-id="4cec2-106">表のセルにアスタリスク ( \* ) が含まれる場合は、準備中です。</span><span class="sxs-lookup"><span data-stu-id="4cec2-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="4cec2-107">Project または Access の要件セットについては、「[Office の共有要件セット](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4cec2-107">For requirement sets for Project or Access, see [Office common requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="4cec2-p103">MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。このバージョンには、ExcelApi 1.1、WordApi 1.1、および共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="4cec2-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="4cec2-110">Excel</span><span class="sxs-lookup"><span data-stu-id="4cec2-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="4cec2-111">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4cec2-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="4cec2-112">拡張点</span><span class="sxs-lookup"><span data-stu-id="4cec2-112">Extension points</span></span></th> 
    <th style="width:20%"><span data-ttu-id="4cec2-113">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4cec2-113">API requirement sets</span></span></th> 
    <th style="width:40%"><span data-ttu-id="4cec2-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4cec2-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="4cec2-115">Office Online</span></span></td>
    <td> <span data-ttu-id="4cec2-116">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4cec2-116">- Taskpane</span></span><br><span data-ttu-id="4cec2-117">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4cec2-117">
        - Content</span></span><br><span data-ttu-id="4cec2-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4cec2-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4cec2-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4cec2-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4cec2-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4cec2-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4cec2-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4cec2-124">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-124">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4cec2-125">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="4cec2-126">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-126">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4cec2-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-127">
        -BindingEvents</span></span><br><span data-ttu-id="4cec2-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-128">
        -DocumentEvents</span></span><br><span data-ttu-id="4cec2-129">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-129">
        -MatrixBindings</span></span><br><span data-ttu-id="4cec2-130">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-130">
        -MatrixCoercion</span></span><br><span data-ttu-id="4cec2-131">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-131">
        -TableBindings</span></span><br><span data-ttu-id="4cec2-132">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-132">
        -TableCoercion</span></span><br><span data-ttu-id="4cec2-133">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-133">
        -TextBindings</span></span><br><span data-ttu-id="4cec2-134">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4cec2-134">
        -CompressedFile</span></span><br><span data-ttu-id="4cec2-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4cec2-135">
        - Settings</span></span><br><span data-ttu-id="4cec2-136">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-136">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-137">Windows 用 Office 2013</span><span class="sxs-lookup"><span data-stu-id="4cec2-137">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="4cec2-138">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4cec2-138">
        - Taskpane</span></span><br><span data-ttu-id="4cec2-139">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4cec2-139">
        - Content</span></span></td>
    <td>  <span data-ttu-id="4cec2-140">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-140">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4cec2-141">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-141">
        -BindingEvents</span></span><br><span data-ttu-id="4cec2-142">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-142">
        -DocumentEvents</span></span><br><span data-ttu-id="4cec2-143">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-143">
        -MatrixBindings</span></span><br><span data-ttu-id="4cec2-144">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-144">
        -MatrixCoercion</span></span><br><span data-ttu-id="4cec2-145">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-145">
        -TableBindings</span></span><br><span data-ttu-id="4cec2-146">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-146">
        -TableCoercion</span></span><br><span data-ttu-id="4cec2-147">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-147">
        -TextBindings</span></span><br><span data-ttu-id="4cec2-148">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4cec2-148">
        - Settings</span></span><br><span data-ttu-id="4cec2-149">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-149">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-150">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4cec2-150">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="4cec2-151">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4cec2-151">- Taskpane</span></span><br><span data-ttu-id="4cec2-152">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4cec2-152">
        - Content</span></span><br><span data-ttu-id="4cec2-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4cec2-154">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-154">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4cec2-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4cec2-156">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-156">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4cec2-157">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-157">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4cec2-158">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-158">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4cec2-159">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-159">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4cec2-160">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-160">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="4cec2-161">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-161">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4cec2-162">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-162">-BindingEvents</span></span><br><span data-ttu-id="4cec2-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-163">
        -DocumentEvents</span></span><br><span data-ttu-id="4cec2-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-164">
        -MatrixBindings</span></span><br><span data-ttu-id="4cec2-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-165">
        -MatrixCoercion</span></span><br><span data-ttu-id="4cec2-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-166">
        -TableBindings</span></span><br><span data-ttu-id="4cec2-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-167">
        -TableCoercion</span></span><br><span data-ttu-id="4cec2-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-168">
        -TextBindings</span></span><br><span data-ttu-id="4cec2-169">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4cec2-169">
        - Settings</span></span><br><span data-ttu-id="4cec2-170">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-170">
        -TextCoercion</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-171">iOS 用 Office</span><span class="sxs-lookup"><span data-stu-id="4cec2-171">Office for iOS</span></span></td>
    <td><span data-ttu-id="4cec2-172">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4cec2-172">- Taskpane</span></span><br><span data-ttu-id="4cec2-173">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4cec2-173">
        - Content</span></span></td>
    <td><span data-ttu-id="4cec2-174">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-174">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4cec2-175">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-175">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4cec2-176">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-176">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4cec2-177">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-177">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4cec2-178">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-178">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4cec2-179">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-179">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4cec2-180">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-180">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="4cec2-181">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-181">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4cec2-182">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-182">-BindingEvents</span></span><br><span data-ttu-id="4cec2-183">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-183">
        -DocumentEvents</span></span><br><span data-ttu-id="4cec2-184">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-184">
        -MatrixBindings</span></span><br><span data-ttu-id="4cec2-185">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-185">
        -MatrixCoercion</span></span><br><span data-ttu-id="4cec2-186">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-186">
        -TableBindings</span></span><br><span data-ttu-id="4cec2-187">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-187">
        -TableCoercion</span></span><br><span data-ttu-id="4cec2-188">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-188">
        -TextBindings</span></span><br><span data-ttu-id="4cec2-189">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4cec2-189">
        - Settings</span></span><br><span data-ttu-id="4cec2-190">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-190">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-191">Mac 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4cec2-191">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="4cec2-192">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4cec2-192">- Taskpane</span></span><br><span data-ttu-id="4cec2-193">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4cec2-193">
        - Content</span></span><br><span data-ttu-id="4cec2-194">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-194">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4cec2-195">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-195">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4cec2-196">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-196">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4cec2-197">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-197">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4cec2-198">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-198">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4cec2-199">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-199">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4cec2-200">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-200">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4cec2-201">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-201">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="4cec2-202">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-202">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4cec2-203">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-203">-BindingEvents</span></span><br><span data-ttu-id="4cec2-204">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-204">
        -DocumentEvents</span></span><br><span data-ttu-id="4cec2-205">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-205">
        -MatrixBindings</span></span><br><span data-ttu-id="4cec2-206">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-206">
        -MatrixCoercion</span></span><br><span data-ttu-id="4cec2-207">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-207">
        -TableBindings</span></span><br><span data-ttu-id="4cec2-208">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-208">
        -TableCoercion</span></span><br><span data-ttu-id="4cec2-209">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-209">
        -TextBindings</span></span><br><span data-ttu-id="4cec2-210">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-210">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="4cec2-211">Outlook</span><span class="sxs-lookup"><span data-stu-id="4cec2-211">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4cec2-212">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4cec2-212">Platform</span></span></th>
    <th><span data-ttu-id="4cec2-213">拡張点</span><span class="sxs-lookup"><span data-stu-id="4cec2-213">Extension points</span></span></th> 
    <th><span data-ttu-id="4cec2-214">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4cec2-214">API requirement sets</span></span></th> 
    <th><span data-ttu-id="4cec2-215"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4cec2-215"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-216">Office Online</span><span class="sxs-lookup"><span data-stu-id="4cec2-216">Office Online</span></span></td>
    <td> <span data-ttu-id="4cec2-217">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4cec2-217">- Mail Read</span></span><br><span data-ttu-id="4cec2-218">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4cec2-218">
      - Mail Compose</span></span><br><span data-ttu-id="4cec2-219">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-219">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4cec2-220">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-220">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4cec2-221">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-221">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4cec2-222">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-222">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4cec2-223">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-223">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4cec2-224">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-224">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4cec2-225">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-225">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4cec2-226">使用不可</span><span class="sxs-lookup"><span data-stu-id="4cec2-226">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-227">Windows 用 Office 2013</span><span class="sxs-lookup"><span data-stu-id="4cec2-227">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4cec2-228">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4cec2-228">- Mail Read</span></span><br><span data-ttu-id="4cec2-229">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4cec2-229">
      - Mail Compose</span></span><br><span data-ttu-id="4cec2-230">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-230">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4cec2-231">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-231">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4cec2-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4cec2-233">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-233">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4cec2-234">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-234">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="4cec2-235">使用不可</span><span class="sxs-lookup"><span data-stu-id="4cec2-235">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-236">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4cec2-236">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4cec2-237">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4cec2-237">- Mail Read</span></span><br><span data-ttu-id="4cec2-238">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4cec2-238">
      - Mail Compose</span></span><br><span data-ttu-id="4cec2-239">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-239">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4cec2-240">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="4cec2-240">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4cec2-241">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-241">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4cec2-242">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-242">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4cec2-243">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-243">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4cec2-244">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-244">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4cec2-245">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-245">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4cec2-246">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-246">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4cec2-247">使用不可</span><span class="sxs-lookup"><span data-stu-id="4cec2-247">Not available</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-248">iOS 用 Office</span><span class="sxs-lookup"><span data-stu-id="4cec2-248">Office for iOS</span></span></td>
    <td> <span data-ttu-id="4cec2-249">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4cec2-249">- Mail Read</span></span><br><span data-ttu-id="4cec2-250">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-250">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4cec2-251">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-251">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4cec2-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4cec2-253">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-253">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4cec2-254">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-254">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4cec2-255">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-255">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>    
    <td><span data-ttu-id="4cec2-256">使用不可</span><span class="sxs-lookup"><span data-stu-id="4cec2-256">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-257">Mac 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4cec2-257">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4cec2-258">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4cec2-258">- Mail Read</span></span><br><span data-ttu-id="4cec2-259">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4cec2-259">
      - Mail Compose</span></span><br><span data-ttu-id="4cec2-260">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-260">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4cec2-261">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-261">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4cec2-262">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-262">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4cec2-263">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-263">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4cec2-264">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-264">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4cec2-265">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-265">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4cec2-266">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-266">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4cec2-267">使用不可</span><span class="sxs-lookup"><span data-stu-id="4cec2-267">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-268">Android 用 Office</span><span class="sxs-lookup"><span data-stu-id="4cec2-268">Office for Android</span></span></td>
    <td> <span data-ttu-id="4cec2-269">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4cec2-269">- Mail Read</span></span><br><span data-ttu-id="4cec2-270">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-270">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4cec2-271">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-271">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4cec2-272">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-272">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4cec2-273">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-273">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4cec2-274">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-274">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4cec2-275">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-275">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="4cec2-276">使用不可</span><span class="sxs-lookup"><span data-stu-id="4cec2-276">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="4cec2-277">Word</span><span class="sxs-lookup"><span data-stu-id="4cec2-277">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4cec2-278">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4cec2-278">Platform</span></span></th>
    <th><span data-ttu-id="4cec2-279">拡張点</span><span class="sxs-lookup"><span data-stu-id="4cec2-279">Extension points</span></span></th> 
    <th><span data-ttu-id="4cec2-280">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4cec2-280">API requirement sets</span></span></th> 
    <th><span data-ttu-id="4cec2-281"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4cec2-281"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-282">Office Online</span><span class="sxs-lookup"><span data-stu-id="4cec2-282">Office Online</span></span></td>
    <td> <span data-ttu-id="4cec2-283">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4cec2-283">- Taskpane</span></span><br><span data-ttu-id="4cec2-284">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-284">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4cec2-285">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-285">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4cec2-286">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-286">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4cec2-287">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-287">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4cec2-288">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-288">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4cec2-289">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-289">-BindingEvents</span></span><br><span data-ttu-id="4cec2-290">
         - CustomXMLParts</span><span class="sxs-lookup"><span data-stu-id="4cec2-290">
         -</span></span><br><span data-ttu-id="4cec2-291">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-291">
         -MatrixBindings</span></span><br><span data-ttu-id="4cec2-292">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-292">
         -MatrixCoercion</span></span><br><span data-ttu-id="4cec2-293">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-293">
         -TableBindings</span></span><br><span data-ttu-id="4cec2-294">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-294">
         -TableCoercion</span></span><br><span data-ttu-id="4cec2-295">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-295">
         -TextBindings</span></span><br><span data-ttu-id="4cec2-296">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-296">
         -DocumentEvents</span></span><br><span data-ttu-id="4cec2-297">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4cec2-297">
         -TextFile</span></span><br><span data-ttu-id="4cec2-298">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-298">
         -ImageCoercion</span></span><br><span data-ttu-id="4cec2-299">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4cec2-299">
         - Settings</span></span><br><span data-ttu-id="4cec2-300">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-300">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-301">Windows 用 Office 2013</span><span class="sxs-lookup"><span data-stu-id="4cec2-301">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4cec2-302">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4cec2-302">- Taskpane</span></span></td>
    <td> <span data-ttu-id="4cec2-303">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-303">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4cec2-304">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-304">-BindingEvents</span></span><br><span data-ttu-id="4cec2-305">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4cec2-305">
         -CompressedFile</span></span><br><span data-ttu-id="4cec2-306">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="4cec2-306">
         -CustomXmlPart</span></span><br><span data-ttu-id="4cec2-307">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-307">
         -DocumentEvents</span></span><br><span data-ttu-id="4cec2-308">
         - File</span><span class="sxs-lookup"><span data-stu-id="4cec2-308">
         - File</span></span><br><span data-ttu-id="4cec2-309">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-309">
         -HtmlCoercion</span></span><br><span data-ttu-id="4cec2-310">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-310">
         -ImageCoercion</span></span><br><span data-ttu-id="4cec2-311">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-311">
         -OoxmlCoercion</span></span><br><span data-ttu-id="4cec2-312">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-312">
         -TableBindings</span></span><br><span data-ttu-id="4cec2-313">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-313">
         -TableCoercion</span></span><br><span data-ttu-id="4cec2-314">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-314">
         -TextBindings</span></span><br><span data-ttu-id="4cec2-315">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4cec2-315">
         -TextFile</span></span><br><span data-ttu-id="4cec2-316">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4cec2-316">
         - Settings</span></span><br><span data-ttu-id="4cec2-317">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-317">
         -TextCoercion</span></span><br><span data-ttu-id="4cec2-318">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-318">
         -MatrixCoercion</span></span><br><span data-ttu-id="4cec2-319">
         - Matrix Bindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-319">
         - Matrix Bindings</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-320">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4cec2-320">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4cec2-321">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4cec2-321">- Taskpane</span></span><br><span data-ttu-id="4cec2-322">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-322">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4cec2-323">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-323">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4cec2-324">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-324">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4cec2-325">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-325">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4cec2-326">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-326">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4cec2-327">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-327">-BindingEvents</span></span><br><span data-ttu-id="4cec2-328">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4cec2-328">
         -CompressedFile</span></span><br><span data-ttu-id="4cec2-329">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="4cec2-329">
         -CustomXmlPart</span></span><br><span data-ttu-id="4cec2-330">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-330">
         -DocumentEvents</span></span><br><span data-ttu-id="4cec2-331">
         - File</span><span class="sxs-lookup"><span data-stu-id="4cec2-331">
         - File</span></span><br><span data-ttu-id="4cec2-332">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-332">
         -HtmlCoercion</span></span><br><span data-ttu-id="4cec2-333">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-333">
         -ImageCoercion</span></span><br><span data-ttu-id="4cec2-334">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-334">
         -OoxmlCoercion</span></span><br><span data-ttu-id="4cec2-335">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-335">
         -TableBindings</span></span><br><span data-ttu-id="4cec2-336">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-336">
         -TableCoercion</span></span><br><span data-ttu-id="4cec2-337">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-337">
         -TextBindings</span></span><br><span data-ttu-id="4cec2-338">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4cec2-338">
         -TextFile</span></span><br><span data-ttu-id="4cec2-339">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4cec2-339">
         - Settings</span></span><br><span data-ttu-id="4cec2-340">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-340">
         -TextCoercion</span></span><br><span data-ttu-id="4cec2-341">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-341">
         -MatrixCoercion</span></span><br><span data-ttu-id="4cec2-342">
         - Matrix Bindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-342">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-343">iOS 用 Office</span><span class="sxs-lookup"><span data-stu-id="4cec2-343">Office for iOS</span></span></td>
    <td> <span data-ttu-id="4cec2-344">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4cec2-344">- Taskpane</span></span></td>
    <td> <span data-ttu-id="4cec2-345">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-345">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4cec2-346">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-346">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4cec2-347">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-347">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4cec2-348">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4cec2-348">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4cec2-349">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-349">-BindingEvents</span></span><br><span data-ttu-id="4cec2-350">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4cec2-350">
         -CompressedFile</span></span><br><span data-ttu-id="4cec2-351">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="4cec2-351">
         -CustomXmlPart</span></span><br><span data-ttu-id="4cec2-352">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-352">
         -DocumentEvents</span></span><br><span data-ttu-id="4cec2-353">
         - File</span><span class="sxs-lookup"><span data-stu-id="4cec2-353">
         - File</span></span><br><span data-ttu-id="4cec2-354">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-354">
         -HtmlCoercion</span></span><br><span data-ttu-id="4cec2-355">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-355">
         -ImageCoercion</span></span><br><span data-ttu-id="4cec2-356">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-356">
         -OoxmlCoercion</span></span><br><span data-ttu-id="4cec2-357">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-357">
         -TableBindings</span></span><br><span data-ttu-id="4cec2-358">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-358">
         -TableCoercion</span></span><br><span data-ttu-id="4cec2-359">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-359">
         -TextBindings</span></span><br><span data-ttu-id="4cec2-360">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4cec2-360">
         -TextFile</span></span><br><span data-ttu-id="4cec2-361">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4cec2-361">
         - Settings</span></span><br><span data-ttu-id="4cec2-362">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-362">
         -TextCoercion</span></span><br><span data-ttu-id="4cec2-363">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-363">
         -MatrixCoercion</span></span><br><span data-ttu-id="4cec2-364">
         - Matrix Bindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-364">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-365">Mac 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4cec2-365">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4cec2-366">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4cec2-366">- Taskpane</span></span><br><span data-ttu-id="4cec2-367">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-367">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4cec2-368">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-368">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4cec2-369">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-369">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4cec2-370">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-370">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4cec2-371">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4cec2-371">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4cec2-372">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-372">-BindingEvents</span></span><br><span data-ttu-id="4cec2-373">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4cec2-373">
         -CompressedFile</span></span><br><span data-ttu-id="4cec2-374">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="4cec2-374">
         -CustomXmlPart</span></span><br><span data-ttu-id="4cec2-375">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-375">
         -DocumentEvents</span></span><br><span data-ttu-id="4cec2-376">
         - File</span><span class="sxs-lookup"><span data-stu-id="4cec2-376">
         - File</span></span><br><span data-ttu-id="4cec2-377">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-377">
         -HtmlCoercion</span></span><br><span data-ttu-id="4cec2-378">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-378">
         -ImageCoercion</span></span><br><span data-ttu-id="4cec2-379">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-379">
         -OoxmlCoercion</span></span><br><span data-ttu-id="4cec2-380">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-380">
         -TableBindings</span></span><br><span data-ttu-id="4cec2-381">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-381">
         -TableCoercion</span></span><br><span data-ttu-id="4cec2-382">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-382">
         -TextBindings</span></span><br><span data-ttu-id="4cec2-383">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4cec2-383">
         -TextFile</span></span><br><span data-ttu-id="4cec2-384">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4cec2-384">
         - Settings</span></span><br><span data-ttu-id="4cec2-385">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-385">
         -TextCoercion</span></span><br><span data-ttu-id="4cec2-386">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-386">
         -MatrixCoercion</span></span><br><span data-ttu-id="4cec2-387">
         - Matrix Bindings</span><span class="sxs-lookup"><span data-stu-id="4cec2-387">
         - Matrix Bindings</span></span> </td> 
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="4cec2-388">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="4cec2-388">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4cec2-389">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4cec2-389">Platform</span></span></th>
    <th><span data-ttu-id="4cec2-390">拡張点</span><span class="sxs-lookup"><span data-stu-id="4cec2-390">Extension points</span></span></th> 
    <th><span data-ttu-id="4cec2-391">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4cec2-391">API requirement sets</span></span></th> 
    <th><span data-ttu-id="4cec2-392"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4cec2-392"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-393">Office Online</span><span class="sxs-lookup"><span data-stu-id="4cec2-393">Office Online</span></span></td>
    <td> <span data-ttu-id="4cec2-394">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4cec2-394">- Content</span></span><br><span data-ttu-id="4cec2-395">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4cec2-395">
         - Taskpane</span></span><br><span data-ttu-id="4cec2-396">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-396">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4cec2-397">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-397">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4cec2-398">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4cec2-398">-ActiveView</span></span><br><span data-ttu-id="4cec2-399">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4cec2-399">
         -CompressedFile</span></span><br><span data-ttu-id="4cec2-400">
         - File</span><span class="sxs-lookup"><span data-stu-id="4cec2-400">
         - File</span></span><br><span data-ttu-id="4cec2-401">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4cec2-401">
         - Selection</span></span><br><span data-ttu-id="4cec2-402">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4cec2-402">
         - Settings</span></span><br><span data-ttu-id="4cec2-403">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-403">
         -TextCoercion</span></span><br><span data-ttu-id="4cec2-404">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-404">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-405">Windows 用 Office 2013</span><span class="sxs-lookup"><span data-stu-id="4cec2-405">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4cec2-406">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4cec2-406">- Content</span></span><br><span data-ttu-id="4cec2-407">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4cec2-407">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="4cec2-408">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4cec2-408">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4cec2-409">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4cec2-409">-ActiveView</span></span><br><span data-ttu-id="4cec2-410">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4cec2-410">
         -CompressedFile</span></span><br><span data-ttu-id="4cec2-411">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-411">
         -DocumentEvents</span></span><br><span data-ttu-id="4cec2-412">
         - File</span><span class="sxs-lookup"><span data-stu-id="4cec2-412">
         - File</span></span><br><span data-ttu-id="4cec2-413">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4cec2-413">
         - Selection</span></span><br><span data-ttu-id="4cec2-414">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4cec2-414">
         - Settings</span></span><br><span data-ttu-id="4cec2-415">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-415">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-416">Windows 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4cec2-416">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4cec2-417">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4cec2-417">- Content</span></span><br><span data-ttu-id="4cec2-418">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4cec2-418">
         - Taskpane</span></span><br><span data-ttu-id="4cec2-419">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-419">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4cec2-420">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-420">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4cec2-421">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4cec2-421">-ActiveView</span></span><br><span data-ttu-id="4cec2-422">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4cec2-422">
         -CompressedFile</span></span><br><span data-ttu-id="4cec2-423">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-423">
         -DocumentEvents</span></span><br><span data-ttu-id="4cec2-424">
         - DocuentEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-424">
         - File</span></span><br><span data-ttu-id="4cec2-425">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4cec2-425">
         - Selection</span></span><br><span data-ttu-id="4cec2-426">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4cec2-426">
         - Settings</span></span><br><span data-ttu-id="4cec2-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-427">
         -TextCoercion</span></span><br><span data-ttu-id="4cec2-428">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-428">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-429">iOS 用 Office</span><span class="sxs-lookup"><span data-stu-id="4cec2-429">Office for iOS</span></span></td>
    <td> <span data-ttu-id="4cec2-430">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4cec2-430">- Content</span></span><br><span data-ttu-id="4cec2-431">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4cec2-431">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="4cec2-432">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-432">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="4cec2-433">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4cec2-433">-ActiveView</span></span><br><span data-ttu-id="4cec2-434">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4cec2-434">
         -CompressedFile</span></span><br><span data-ttu-id="4cec2-435">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-435">
         -DocumentEvents</span></span><br><span data-ttu-id="4cec2-436">
         - File</span><span class="sxs-lookup"><span data-stu-id="4cec2-436">
         - File</span></span><br><span data-ttu-id="4cec2-437">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4cec2-437">
         - Selection</span></span><br><span data-ttu-id="4cec2-438">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4cec2-438">
         - Settings</span></span><br><span data-ttu-id="4cec2-439">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-439">
         -TextCoercion</span></span><br><span data-ttu-id="4cec2-440">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-440">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-441">Mac 用 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4cec2-441">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4cec2-442">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4cec2-442">- Content</span></span><br><span data-ttu-id="4cec2-443">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4cec2-443">
         - Taskpane</span></span><br><span data-ttu-id="4cec2-444">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-444">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4cec2-445">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-445">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4cec2-446">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4cec2-446">-ActiveView</span></span><br><span data-ttu-id="4cec2-447">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4cec2-447">
         -CompressedFile</span></span><br><span data-ttu-id="4cec2-448">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-448">
         -DocumentEvents</span></span><br><span data-ttu-id="4cec2-449">
         - File</span><span class="sxs-lookup"><span data-stu-id="4cec2-449">
         - File</span></span><br><span data-ttu-id="4cec2-450">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4cec2-450">
         - Selection</span></span><br><span data-ttu-id="4cec2-451">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4cec2-451">
         - Settings</span></span><br><span data-ttu-id="4cec2-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-452">
         -TextCoercion</span></span><br><span data-ttu-id="4cec2-453">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-453">
         -ImageCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="4cec2-454">OneNote</span><span class="sxs-lookup"><span data-stu-id="4cec2-454">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4cec2-455">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4cec2-455">Platform</span></span></th>
    <th><span data-ttu-id="4cec2-456">拡張点</span><span class="sxs-lookup"><span data-stu-id="4cec2-456">Extension points</span></span></th> 
    <th><span data-ttu-id="4cec2-457">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4cec2-457">API requirement sets</span></span></th> 
    <th><span data-ttu-id="4cec2-458"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4cec2-458"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="4cec2-459">Office Online</span><span class="sxs-lookup"><span data-stu-id="4cec2-459">Office Online</span></span></td>
    <td> <span data-ttu-id="4cec2-460">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4cec2-460">- Content</span></span><br><span data-ttu-id="4cec2-461">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4cec2-461">
         - Taskpane</span></span><br><span data-ttu-id="4cec2-462">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-462">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4cec2-463">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-463">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="4cec2-464">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4cec2-464">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4cec2-465">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4cec2-465">-DocumentEvents</span></span><br><span data-ttu-id="4cec2-466">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4cec2-466">
         - Settings</span></span><br><span data-ttu-id="4cec2-467">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-467">
         -TextCoercion</span></span><br><span data-ttu-id="4cec2-468">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-468">
         -HtmlCoercion</span></span><br><span data-ttu-id="4cec2-469">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4cec2-469">
         -ImageCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="4cec2-470">関連項目</span><span class="sxs-lookup"><span data-stu-id="4cec2-470">See also</span></span>

- [<span data-ttu-id="4cec2-471">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="4cec2-471">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="4cec2-472">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="4cec2-472">Common API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="4cec2-473">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="4cec2-473">Add-in Commands requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="4cec2-474">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="4cec2-474">JavaScript API for Office reference</span></span>](https://dev.office.com/reference/add-ins/javascript-api-for-office)

