---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、Word、Outlook、PowerPoint、OneNote、および Project のサポートされる要件セット。
ms.date: 11/07/2018
ms.openlocfilehash: c3da40be21c0e569028dd10e93e33760ba2bd39d
ms.sourcegitcommit: 3e84d616e69f39eeeeea773f2431e7d674c4a9f5
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/22/2018
ms.locfileid: "26644754"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="911d4-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="911d4-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="911d4-104">Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。</span><span class="sxs-lookup"><span data-stu-id="911d4-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="911d4-105">次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API の要件セットが含まれています。</span><span class="sxs-lookup"><span data-stu-id="911d4-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="911d4-p102">MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。このバージョンには、ExcelApi 1.1、WordApi 1.1、および共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="911d4-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="911d4-108">Excel</span><span class="sxs-lookup"><span data-stu-id="911d4-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="911d4-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="911d4-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="911d4-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="911d4-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="911d4-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="911d4-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="911d4-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="911d4-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="911d4-113">Office Online</span></span></td>
    <td> <span data-ttu-id="911d4-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-114">- Taskpane</span></span><br><span data-ttu-id="911d4-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="911d4-115">
        - Content</span></span><br><span data-ttu-id="911d4-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="911d4-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="911d4-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="911d4-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="911d4-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="911d4-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="911d4-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="911d4-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="911d4-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="911d4-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="911d4-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="911d4-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="911d4-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="911d4-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="911d4-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="911d4-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="911d4-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="911d4-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="911d4-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-126">
        -BindingEvents</span></span><br><span data-ttu-id="911d4-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="911d4-127">
        -CompressedFile</span></span><br><span data-ttu-id="911d4-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-128">
        -DocumentEvents</span></span><br><span data-ttu-id="911d4-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="911d4-129">
        - File</span></span><br><span data-ttu-id="911d4-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-130">
        -MatrixBindings</span></span><br><span data-ttu-id="911d4-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-131">
        -MatrixCoercion</span></span><br><span data-ttu-id="911d4-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-132">
        - Selection</span></span><br><span data-ttu-id="911d4-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-133">
        - Settings</span></span><br><span data-ttu-id="911d4-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-134">
        -TableBindings</span></span><br><span data-ttu-id="911d4-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-135">
        -TableCoercion</span></span><br><span data-ttu-id="911d4-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-136">
        -TextBindings</span></span><br><span data-ttu-id="911d4-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-137">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-138">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="911d4-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="911d4-139">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-139">
        - Taskpane</span></span><br><span data-ttu-id="911d4-140">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="911d4-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="911d4-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="911d4-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-142">
        -BindingEvents</span></span><br><span data-ttu-id="911d4-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="911d4-143">
        -CompressedFile</span></span><br><span data-ttu-id="911d4-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-144">
        -DocumentEvents</span></span><br><span data-ttu-id="911d4-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="911d4-145">
        - File</span></span><br><span data-ttu-id="911d4-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-146">
        -ImageCoercion</span></span><br><span data-ttu-id="911d4-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-147">
        -MatrixBindings</span></span><br><span data-ttu-id="911d4-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-148">
        -MatrixCoercion</span></span><br><span data-ttu-id="911d4-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-149">
        - Selection</span></span><br><span data-ttu-id="911d4-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-150">
        - Settings</span></span><br><span data-ttu-id="911d4-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-151">
        -TableBindings</span></span><br><span data-ttu-id="911d4-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-152">
        -TableCoercion</span></span><br><span data-ttu-id="911d4-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-153">
        -TextBindings</span></span><br><span data-ttu-id="911d4-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-154">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-155">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="911d4-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="911d4-156">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-156">- Taskpane</span></span><br><span data-ttu-id="911d4-157">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="911d4-157">
        - Content</span></span><br><span data-ttu-id="911d4-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="911d4-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="911d4-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="911d4-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="911d4-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="911d4-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="911d4-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="911d4-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="911d4-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="911d4-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="911d4-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="911d4-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="911d4-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="911d4-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="911d4-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="911d4-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="911d4-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="911d4-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-168">-BindingEvents</span></span><br><span data-ttu-id="911d4-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="911d4-169">
        -CompressedFile</span></span><br><span data-ttu-id="911d4-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-170">
        -DocumentEvents</span></span><br><span data-ttu-id="911d4-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="911d4-171">
        - File</span></span><br><span data-ttu-id="911d4-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-172">
        -ImageCoercion</span></span><br><span data-ttu-id="911d4-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-173">
        -MatrixBindings</span></span><br><span data-ttu-id="911d4-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-174">
        -MatrixCoercion</span></span><br><span data-ttu-id="911d4-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-175">
        - Selection</span></span><br><span data-ttu-id="911d4-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-176">
        - Settings</span></span><br><span data-ttu-id="911d4-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-177">
        -TableBindings</span></span><br><span data-ttu-id="911d4-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-178">
        -TableCoercion</span></span><br><span data-ttu-id="911d4-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-179">
        -TextBindings</span></span><br><span data-ttu-id="911d4-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-181">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="911d4-181">Outlook 2019 for Windows</span></span></td>
    <td><span data-ttu-id="911d4-182">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-182">- Taskpane</span></span><br><span data-ttu-id="911d4-183">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="911d4-183">
        - Content</span></span><br><span data-ttu-id="911d4-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="911d4-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="911d4-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="911d4-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="911d4-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="911d4-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="911d4-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="911d4-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="911d4-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="911d4-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="911d4-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="911d4-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="911d4-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="911d4-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="911d4-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="911d4-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="911d4-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="911d4-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-194">-BindingEvents</span></span><br><span data-ttu-id="911d4-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="911d4-195">
        -CompressedFile</span></span><br><span data-ttu-id="911d4-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-196">
        -DocumentEvents</span></span><br><span data-ttu-id="911d4-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="911d4-197">
        - File</span></span><br><span data-ttu-id="911d4-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-198">
        -ImageCoercion</span></span><br><span data-ttu-id="911d4-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-199">
        -MatrixBindings</span></span><br><span data-ttu-id="911d4-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-200">
        -MatrixCoercion</span></span><br><span data-ttu-id="911d4-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-201">
        - Selection</span></span><br><span data-ttu-id="911d4-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-202">
        - Settings</span></span><br><span data-ttu-id="911d4-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-203">
        -TableBindings</span></span><br><span data-ttu-id="911d4-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-204">
        -TableCoercion</span></span><br><span data-ttu-id="911d4-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-205">
        -TextBindings</span></span><br><span data-ttu-id="911d4-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-206">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-207">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="911d4-207">Office for iPad</span></span></td>
    <td><span data-ttu-id="911d4-208">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-208">- Taskpane</span></span><br><span data-ttu-id="911d4-209">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="911d4-209">
        - Content</span></span></td>
    <td><span data-ttu-id="911d4-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="911d4-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="911d4-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="911d4-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="911d4-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="911d4-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="911d4-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="911d4-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="911d4-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="911d4-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="911d4-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="911d4-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="911d4-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="911d4-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="911d4-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="911d4-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="911d4-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-219">-BindingEvents</span></span><br><span data-ttu-id="911d4-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="911d4-220">
        -CompressedFile</span></span><br><span data-ttu-id="911d4-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-221">
        -DocumentEvents</span></span><br><span data-ttu-id="911d4-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="911d4-222">
        - File</span></span><br><span data-ttu-id="911d4-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-223">
        -ImageCoercion</span></span><br><span data-ttu-id="911d4-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-224">
        -MatrixBindings</span></span><br><span data-ttu-id="911d4-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-225">
        -MatrixCoercion</span></span><br><span data-ttu-id="911d4-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-226">
        - Selection</span></span><br><span data-ttu-id="911d4-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-227">
        - Settings</span></span><br><span data-ttu-id="911d4-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-228">
        -TableBindings</span></span><br><span data-ttu-id="911d4-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-229">
        -TableCoercion</span></span><br><span data-ttu-id="911d4-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-230">
        -TextBindings</span></span><br><span data-ttu-id="911d4-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-231">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-232">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="911d4-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="911d4-233">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-233">- Taskpane</span></span><br><span data-ttu-id="911d4-234">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="911d4-234">
        - Content</span></span><br><span data-ttu-id="911d4-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="911d4-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="911d4-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="911d4-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="911d4-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="911d4-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="911d4-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="911d4-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="911d4-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="911d4-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="911d4-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="911d4-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="911d4-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="911d4-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="911d4-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="911d4-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="911d4-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="911d4-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-245">-BindingEvents</span></span><br><span data-ttu-id="911d4-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="911d4-246">
        -CompressedFile</span></span><br><span data-ttu-id="911d4-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-247">
        -DocumentEvents</span></span><br><span data-ttu-id="911d4-248">
        - File</span><span class="sxs-lookup"><span data-stu-id="911d4-248">
        - File</span></span><br><span data-ttu-id="911d4-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-249">
        -ImageCoercion</span></span><br><span data-ttu-id="911d4-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-250">
        -MatrixBindings</span></span><br><span data-ttu-id="911d4-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-251">
        -MatrixCoercion</span></span><br><span data-ttu-id="911d4-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="911d4-252">
        -PdfFile</span></span><br><span data-ttu-id="911d4-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-253">
        - Selection</span></span><br><span data-ttu-id="911d4-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-254">
        - Settings</span></span><br><span data-ttu-id="911d4-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-255">
        -TableBindings</span></span><br><span data-ttu-id="911d4-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-256">
        -TableCoercion</span></span><br><span data-ttu-id="911d4-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-257">
        -TextBindings</span></span><br><span data-ttu-id="911d4-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-258">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-259">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="911d4-259">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="911d4-260">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-260">- Taskpane</span></span><br><span data-ttu-id="911d4-261">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="911d4-261">
        - Content</span></span><br><span data-ttu-id="911d4-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="911d4-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="911d4-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="911d4-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="911d4-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="911d4-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="911d4-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="911d4-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="911d4-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="911d4-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="911d4-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="911d4-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="911d4-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="911d4-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="911d4-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="911d4-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="911d4-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="911d4-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-272">-BindingEvents</span></span><br><span data-ttu-id="911d4-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="911d4-273">
        -CompressedFile</span></span><br><span data-ttu-id="911d4-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-274">
        -DocumentEvents</span></span><br><span data-ttu-id="911d4-275">
        - File</span><span class="sxs-lookup"><span data-stu-id="911d4-275">
        - File</span></span><br><span data-ttu-id="911d4-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-276">
        -ImageCoercion</span></span><br><span data-ttu-id="911d4-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-277">
        -MatrixBindings</span></span><br><span data-ttu-id="911d4-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-278">
        -MatrixCoercion</span></span><br><span data-ttu-id="911d4-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="911d4-279">
        -PdfFile</span></span><br><span data-ttu-id="911d4-280">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-280">
        - Selection</span></span><br><span data-ttu-id="911d4-281">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-281">
        - Settings</span></span><br><span data-ttu-id="911d4-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-282">
        -TableBindings</span></span><br><span data-ttu-id="911d4-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-283">
        -TableCoercion</span></span><br><span data-ttu-id="911d4-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-284">
        -TextBindings</span></span><br><span data-ttu-id="911d4-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-285">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="911d4-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="911d4-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="911d4-287">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="911d4-287">Platform</span></span></th>
    <th><span data-ttu-id="911d4-288">拡張点</span><span class="sxs-lookup"><span data-stu-id="911d4-288">Extension points</span></span></th>
    <th><span data-ttu-id="911d4-289">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="911d4-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="911d4-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="911d4-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="911d4-291">Office Online</span></span></td>
    <td> <span data-ttu-id="911d4-292">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="911d4-292">- Mail Read</span></span><br><span data-ttu-id="911d4-293">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="911d4-293">
      - Mail Compose</span></span><br><span data-ttu-id="911d4-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="911d4-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="911d4-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="911d4-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="911d4-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="911d4-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="911d4-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="911d4-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="911d4-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="911d4-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="911d4-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="911d4-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="911d4-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="911d4-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="911d4-302">利用不可</span><span class="sxs-lookup"><span data-stu-id="911d4-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-303">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="911d4-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="911d4-304">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="911d4-304">- Mail Read</span></span><br><span data-ttu-id="911d4-305">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="911d4-305">
      - Mail Compose</span></span><br><span data-ttu-id="911d4-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="911d4-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="911d4-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="911d4-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="911d4-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="911d4-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="911d4-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="911d4-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="911d4-311">利用不可</span><span class="sxs-lookup"><span data-stu-id="911d4-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-312">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="911d4-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="911d4-313">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="911d4-313">- Mail Read</span></span><br><span data-ttu-id="911d4-314">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="911d4-314">
      - Mail Compose</span></span><br><span data-ttu-id="911d4-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="911d4-316">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="911d4-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="911d4-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="911d4-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="911d4-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="911d4-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="911d4-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="911d4-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="911d4-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="911d4-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="911d4-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="911d4-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="911d4-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="911d4-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="911d4-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="911d4-324">利用不可</span><span class="sxs-lookup"><span data-stu-id="911d4-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-325">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="911d4-325">Outlook 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="911d4-326">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="911d4-326">- Mail Read</span></span><br><span data-ttu-id="911d4-327">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="911d4-327">
      - Mail Compose</span></span><br><span data-ttu-id="911d4-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="911d4-329">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="911d4-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="911d4-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="911d4-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="911d4-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="911d4-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="911d4-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="911d4-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="911d4-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="911d4-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="911d4-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="911d4-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="911d4-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="911d4-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="911d4-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="911d4-337">利用不可</span><span class="sxs-lookup"><span data-stu-id="911d4-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-338">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="911d4-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="911d4-339">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="911d4-339">- Mail Read</span></span><br><span data-ttu-id="911d4-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="911d4-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="911d4-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="911d4-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="911d4-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="911d4-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="911d4-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="911d4-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="911d4-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="911d4-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="911d4-346">利用不可</span><span class="sxs-lookup"><span data-stu-id="911d4-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-347">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="911d4-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="911d4-348">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="911d4-348">- Mail Read</span></span><br><span data-ttu-id="911d4-349">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="911d4-349">
      - Mail Compose</span></span><br><span data-ttu-id="911d4-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="911d4-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="911d4-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="911d4-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="911d4-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="911d4-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="911d4-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="911d4-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="911d4-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="911d4-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="911d4-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="911d4-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="911d4-357">利用不可</span><span class="sxs-lookup"><span data-stu-id="911d4-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-358">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="911d4-358">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="911d4-359">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="911d4-359">- Mail Read</span></span><br><span data-ttu-id="911d4-360">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="911d4-360">
      - Mail Compose</span></span><br><span data-ttu-id="911d4-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="911d4-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="911d4-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="911d4-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="911d4-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="911d4-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="911d4-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="911d4-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="911d4-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="911d4-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="911d4-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="911d4-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="911d4-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="911d4-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="911d4-369">利用不可</span><span class="sxs-lookup"><span data-stu-id="911d4-369">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-370">Office for Android</span><span class="sxs-lookup"><span data-stu-id="911d4-370">Office for Android</span></span></td>
    <td> <span data-ttu-id="911d4-371">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="911d4-371">- Mail Read</span></span><br><span data-ttu-id="911d4-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="911d4-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="911d4-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="911d4-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="911d4-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="911d4-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="911d4-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="911d4-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="911d4-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="911d4-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="911d4-378">利用不可</span><span class="sxs-lookup"><span data-stu-id="911d4-378">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="911d4-379">Word</span><span class="sxs-lookup"><span data-stu-id="911d4-379">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="911d4-380">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="911d4-380">Platform</span></span></th>
    <th><span data-ttu-id="911d4-381">拡張点</span><span class="sxs-lookup"><span data-stu-id="911d4-381">Extension points</span></span></th>
    <th><span data-ttu-id="911d4-382">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="911d4-382">API requirement sets</span></span></th>
    <th><span data-ttu-id="911d4-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="911d4-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-384">Office Online</span><span class="sxs-lookup"><span data-stu-id="911d4-384">Office Online</span></span></td>
    <td> <span data-ttu-id="911d4-385">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-385">- Taskpane</span></span><br><span data-ttu-id="911d4-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="911d4-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="911d4-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="911d4-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="911d4-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="911d4-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="911d4-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="911d4-391">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-391">-BindingEvents</span></span><br><span data-ttu-id="911d4-392">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="911d4-392">
         -</span></span><br><span data-ttu-id="911d4-393">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-393">
         -DocumentEvents</span></span><br><span data-ttu-id="911d4-394">
         - File</span><span class="sxs-lookup"><span data-stu-id="911d4-394">
         - File</span></span><br><span data-ttu-id="911d4-395">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-395">
         -HtmlCoercion</span></span><br><span data-ttu-id="911d4-396">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-396">
         -ImageCoercion</span></span><br><span data-ttu-id="911d4-397">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-397">
         -MatrixBindings</span></span><br><span data-ttu-id="911d4-398">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-398">
         -MatrixCoercion</span></span><br><span data-ttu-id="911d4-399">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-399">
         -OoxmlCoercion</span></span><br><span data-ttu-id="911d4-400">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="911d4-400">
         -PdfFile</span></span><br><span data-ttu-id="911d4-401">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-401">
         - Selection</span></span><br><span data-ttu-id="911d4-402">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-402">
         - Settings</span></span><br><span data-ttu-id="911d4-403">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-403">
         -TableBindings</span></span><br><span data-ttu-id="911d4-404">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-404">
         -TableCoercion</span></span><br><span data-ttu-id="911d4-405">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-405">
         -TextBindings</span></span><br><span data-ttu-id="911d4-406">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-406">
         -TextCoercion</span></span><br><span data-ttu-id="911d4-407">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="911d4-407">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-408">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="911d4-408">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="911d4-409">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-409">- Taskpane</span></span></td>
    <td> <span data-ttu-id="911d4-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="911d4-411">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-411">-BindingEvents</span></span><br><span data-ttu-id="911d4-412">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="911d4-412">
         -CompressedFile</span></span><br><span data-ttu-id="911d4-413">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="911d4-413">
         -</span></span><br><span data-ttu-id="911d4-414">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-414">
         -DocumentEvents</span></span><br><span data-ttu-id="911d4-415">
         - File</span><span class="sxs-lookup"><span data-stu-id="911d4-415">
         - File</span></span><br><span data-ttu-id="911d4-416">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-416">
         -HtmlCoercion</span></span><br><span data-ttu-id="911d4-417">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-417">
         -ImageCoercion</span></span><br><span data-ttu-id="911d4-418">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-418">
         -MatrixBindings</span></span><br><span data-ttu-id="911d4-419">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-419">
         -MatrixCoercion</span></span><br><span data-ttu-id="911d4-420">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-420">
         -OoxmlCoercion</span></span><br><span data-ttu-id="911d4-421">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="911d4-421">
         -PdfFile</span></span><br><span data-ttu-id="911d4-422">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-422">
         - Selection</span></span><br><span data-ttu-id="911d4-423">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-423">
         - Settings</span></span><br><span data-ttu-id="911d4-424">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-424">
         -TableBindings</span></span><br><span data-ttu-id="911d4-425">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-425">
         -TableCoercion</span></span><br><span data-ttu-id="911d4-426">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-426">
         -TextBindings</span></span><br><span data-ttu-id="911d4-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-427">
         -TextCoercion</span></span><br><span data-ttu-id="911d4-428">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="911d4-428">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-429">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="911d4-429">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="911d4-430">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-430">- Taskpane</span></span><br><span data-ttu-id="911d4-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="911d4-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="911d4-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="911d4-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="911d4-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="911d4-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="911d4-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="911d4-436">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-436">-BindingEvents</span></span><br><span data-ttu-id="911d4-437">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="911d4-437">
         -CompressedFile</span></span><br><span data-ttu-id="911d4-438">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="911d4-438">
         -</span></span><br><span data-ttu-id="911d4-439">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-439">
         -DocumentEvents</span></span><br><span data-ttu-id="911d4-440">
         - File</span><span class="sxs-lookup"><span data-stu-id="911d4-440">
         - File</span></span><br><span data-ttu-id="911d4-441">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-441">
         -HtmlCoercion</span></span><br><span data-ttu-id="911d4-442">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-442">
         -ImageCoercion</span></span><br><span data-ttu-id="911d4-443">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-443">
         -MatrixBindings</span></span><br><span data-ttu-id="911d4-444">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-444">
         -MatrixCoercion</span></span><br><span data-ttu-id="911d4-445">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-445">
         -OoxmlCoercion</span></span><br><span data-ttu-id="911d4-446">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="911d4-446">
         -PdfFile</span></span><br><span data-ttu-id="911d4-447">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-447">
         - Selection</span></span><br><span data-ttu-id="911d4-448">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-448">
         - Settings</span></span><br><span data-ttu-id="911d4-449">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-449">
         -TableBindings</span></span><br><span data-ttu-id="911d4-450">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-450">
         -TableCoercion</span></span><br><span data-ttu-id="911d4-451">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-451">
         -TextBindings</span></span><br><span data-ttu-id="911d4-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-452">
         -TextCoercion</span></span><br><span data-ttu-id="911d4-453">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="911d4-453">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-454">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="911d4-454">Outlook 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="911d4-455">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-455">- Taskpane</span></span><br><span data-ttu-id="911d4-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="911d4-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="911d4-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="911d4-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="911d4-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="911d4-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="911d4-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="911d4-461">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-461">-BindingEvents</span></span><br><span data-ttu-id="911d4-462">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="911d4-462">
         -CompressedFile</span></span><br><span data-ttu-id="911d4-463">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="911d4-463">
         -</span></span><br><span data-ttu-id="911d4-464">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-464">
         -DocumentEvents</span></span><br><span data-ttu-id="911d4-465">
         - File</span><span class="sxs-lookup"><span data-stu-id="911d4-465">
         - File</span></span><br><span data-ttu-id="911d4-466">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-466">
         -HtmlCoercion</span></span><br><span data-ttu-id="911d4-467">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-467">
         -ImageCoercion</span></span><br><span data-ttu-id="911d4-468">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-468">
         -MatrixBindings</span></span><br><span data-ttu-id="911d4-469">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-469">
         -MatrixCoercion</span></span><br><span data-ttu-id="911d4-470">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-470">
         -OoxmlCoercion</span></span><br><span data-ttu-id="911d4-471">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="911d4-471">
         -PdfFile</span></span><br><span data-ttu-id="911d4-472">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-472">
         - Selection</span></span><br><span data-ttu-id="911d4-473">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-473">
         - Settings</span></span><br><span data-ttu-id="911d4-474">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-474">
         -TableBindings</span></span><br><span data-ttu-id="911d4-475">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-475">
         -TableCoercion</span></span><br><span data-ttu-id="911d4-476">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-476">
         -TextBindings</span></span><br><span data-ttu-id="911d4-477">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-477">
         -TextCoercion</span></span><br><span data-ttu-id="911d4-478">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="911d4-478">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-479">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="911d4-479">Office for iPad</span></span></td>
    <td> <span data-ttu-id="911d4-480">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-480">- Taskpane</span></span></td>
    <td> <span data-ttu-id="911d4-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="911d4-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="911d4-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="911d4-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="911d4-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="911d4-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="911d4-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="911d4-485">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-485">-BindingEvents</span></span><br><span data-ttu-id="911d4-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="911d4-486">
         -CompressedFile</span></span><br><span data-ttu-id="911d4-487">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="911d4-487">
         -</span></span><br><span data-ttu-id="911d4-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-488">
         -DocumentEvents</span></span><br><span data-ttu-id="911d4-489">
         - File</span><span class="sxs-lookup"><span data-stu-id="911d4-489">
         - File</span></span><br><span data-ttu-id="911d4-490">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-490">
         -HtmlCoercion</span></span><br><span data-ttu-id="911d4-491">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-491">
         -ImageCoercion</span></span><br><span data-ttu-id="911d4-492">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-492">
         -MatrixBindings</span></span><br><span data-ttu-id="911d4-493">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-493">
         -MatrixCoercion</span></span><br><span data-ttu-id="911d4-494">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-494">
         -OoxmlCoercion</span></span><br><span data-ttu-id="911d4-495">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="911d4-495">
         -PdfFile</span></span><br><span data-ttu-id="911d4-496">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-496">
         - Selection</span></span><br><span data-ttu-id="911d4-497">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-497">
         - Settings</span></span><br><span data-ttu-id="911d4-498">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-498">
         -TableBindings</span></span><br><span data-ttu-id="911d4-499">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-499">
         -TableCoercion</span></span><br><span data-ttu-id="911d4-500">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-500">
         -TextBindings</span></span><br><span data-ttu-id="911d4-501">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-501">
         -TextCoercion</span></span><br><span data-ttu-id="911d4-502">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="911d4-502">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-503">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="911d4-503">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="911d4-504">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-504">- Taskpane</span></span><br><span data-ttu-id="911d4-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="911d4-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="911d4-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="911d4-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="911d4-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="911d4-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="911d4-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="911d4-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="911d4-510">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-510">-BindingEvents</span></span><br><span data-ttu-id="911d4-511">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="911d4-511">
         -CompressedFile</span></span><br><span data-ttu-id="911d4-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="911d4-512">
         -</span></span><br><span data-ttu-id="911d4-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-513">
         -DocumentEvents</span></span><br><span data-ttu-id="911d4-514">
         - File</span><span class="sxs-lookup"><span data-stu-id="911d4-514">
         - File</span></span><br><span data-ttu-id="911d4-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-515">
         -HtmlCoercion</span></span><br><span data-ttu-id="911d4-516">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-516">
         -ImageCoercion</span></span><br><span data-ttu-id="911d4-517">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-517">
         -MatrixBindings</span></span><br><span data-ttu-id="911d4-518">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-518">
         -MatrixCoercion</span></span><br><span data-ttu-id="911d4-519">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-519">
         -OoxmlCoercion</span></span><br><span data-ttu-id="911d4-520">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="911d4-520">
         -PdfFile</span></span><br><span data-ttu-id="911d4-521">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-521">
         - Selection</span></span><br><span data-ttu-id="911d4-522">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-522">
         - Settings</span></span><br><span data-ttu-id="911d4-523">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-523">
         -TableBindings</span></span><br><span data-ttu-id="911d4-524">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-524">
         -TableCoercion</span></span><br><span data-ttu-id="911d4-525">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-525">
         -TextBindings</span></span><br><span data-ttu-id="911d4-526">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-526">
         -TextCoercion</span></span><br><span data-ttu-id="911d4-527">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="911d4-527">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-528">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="911d4-528">Outlook 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="911d4-529">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-529">- Taskpane</span></span><br><span data-ttu-id="911d4-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="911d4-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="911d4-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="911d4-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="911d4-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="911d4-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="911d4-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="911d4-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="911d4-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-535">-BindingEvents</span></span><br><span data-ttu-id="911d4-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="911d4-536">
         -CompressedFile</span></span><br><span data-ttu-id="911d4-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="911d4-537">
         -</span></span><br><span data-ttu-id="911d4-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-538">
         -DocumentEvents</span></span><br><span data-ttu-id="911d4-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="911d4-539">
         - File</span></span><br><span data-ttu-id="911d4-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-540">
         -HtmlCoercion</span></span><br><span data-ttu-id="911d4-541">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-541">
         -ImageCoercion</span></span><br><span data-ttu-id="911d4-542">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-542">
         -MatrixBindings</span></span><br><span data-ttu-id="911d4-543">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-543">
         -MatrixCoercion</span></span><br><span data-ttu-id="911d4-544">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-544">
         -OoxmlCoercion</span></span><br><span data-ttu-id="911d4-545">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="911d4-545">
         -PdfFile</span></span><br><span data-ttu-id="911d4-546">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-546">
         - Selection</span></span><br><span data-ttu-id="911d4-547">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-547">
         - Settings</span></span><br><span data-ttu-id="911d4-548">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-548">
         -TableBindings</span></span><br><span data-ttu-id="911d4-549">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-549">
         -TableCoercion</span></span><br><span data-ttu-id="911d4-550">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="911d4-550">
         -TextBindings</span></span><br><span data-ttu-id="911d4-551">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-551">
         -TextCoercion</span></span><br><span data-ttu-id="911d4-552">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="911d4-552">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="911d4-553">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="911d4-553">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="911d4-554">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="911d4-554">Platform</span></span></th>
    <th><span data-ttu-id="911d4-555">拡張点</span><span class="sxs-lookup"><span data-stu-id="911d4-555">Extension points</span></span></th>
    <th><span data-ttu-id="911d4-556">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="911d4-556">API requirement sets</span></span></th>
    <th><span data-ttu-id="911d4-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="911d4-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-558">Office Online</span><span class="sxs-lookup"><span data-stu-id="911d4-558">Office Online</span></span></td>
    <td> <span data-ttu-id="911d4-559">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="911d4-559">- Content</span></span><br><span data-ttu-id="911d4-560">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-560">
         - Taskpane</span></span><br><span data-ttu-id="911d4-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="911d4-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="911d4-563">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="911d4-563">-</span></span><br><span data-ttu-id="911d4-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="911d4-564">
         -CompressedFile</span></span><br><span data-ttu-id="911d4-565">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-565">
         -DocumentEvents</span></span><br><span data-ttu-id="911d4-566">
         - File</span><span class="sxs-lookup"><span data-stu-id="911d4-566">
         - File</span></span><br><span data-ttu-id="911d4-567">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-567">
         -ImageCoercion</span></span><br><span data-ttu-id="911d4-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="911d4-568">
         -PdfFile</span></span><br><span data-ttu-id="911d4-569">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-569">
         - Selection</span></span><br><span data-ttu-id="911d4-570">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-570">
         - Settings</span></span><br><span data-ttu-id="911d4-571">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-571">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-572">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="911d4-572">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="911d4-573">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="911d4-573">- Content</span></span><br><span data-ttu-id="911d4-574">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-574">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="911d4-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="911d4-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="911d4-576">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="911d4-576">-</span></span><br><span data-ttu-id="911d4-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="911d4-577">
         -CompressedFile</span></span><br><span data-ttu-id="911d4-578">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-578">
         -DocumentEvents</span></span><br><span data-ttu-id="911d4-579">
         - File</span><span class="sxs-lookup"><span data-stu-id="911d4-579">
         - File</span></span><br><span data-ttu-id="911d4-580">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-580">
         -ImageCoercion</span></span><br><span data-ttu-id="911d4-581">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="911d4-581">
         -PdfFile</span></span><br><span data-ttu-id="911d4-582">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-582">
         - Selection</span></span><br><span data-ttu-id="911d4-583">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-583">
         - Settings</span></span><br><span data-ttu-id="911d4-584">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-584">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-585">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="911d4-585">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="911d4-586">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="911d4-586">- Content</span></span><br><span data-ttu-id="911d4-587">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-587">
         - Taskpane</span></span><br><span data-ttu-id="911d4-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="911d4-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="911d4-590">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="911d4-590">-</span></span><br><span data-ttu-id="911d4-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="911d4-591">
         -CompressedFile</span></span><br><span data-ttu-id="911d4-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-592">
         -DocumentEvents</span></span><br><span data-ttu-id="911d4-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="911d4-593">
         - File</span></span><br><span data-ttu-id="911d4-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-594">
         -ImageCoercion</span></span><br><span data-ttu-id="911d4-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="911d4-595">
         -PdfFile</span></span><br><span data-ttu-id="911d4-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-596">
         - Selection</span></span><br><span data-ttu-id="911d4-597">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-597">
         - Settings</span></span><br><span data-ttu-id="911d4-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-598">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-599">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="911d4-599">Outlook 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="911d4-600">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="911d4-600">- Content</span></span><br><span data-ttu-id="911d4-601">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-601">
         - Taskpane</span></span><br><span data-ttu-id="911d4-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="911d4-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="911d4-604">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="911d4-604">-</span></span><br><span data-ttu-id="911d4-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="911d4-605">
         -CompressedFile</span></span><br><span data-ttu-id="911d4-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-606">
         -DocumentEvents</span></span><br><span data-ttu-id="911d4-607">
         - File</span><span class="sxs-lookup"><span data-stu-id="911d4-607">
         - File</span></span><br><span data-ttu-id="911d4-608">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-608">
         -ImageCoercion</span></span><br><span data-ttu-id="911d4-609">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="911d4-609">
         -PdfFile</span></span><br><span data-ttu-id="911d4-610">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-610">
         - Selection</span></span><br><span data-ttu-id="911d4-611">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-611">
         - Settings</span></span><br><span data-ttu-id="911d4-612">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-612">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-613">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="911d4-613">Office for iPad</span></span></td>
    <td> <span data-ttu-id="911d4-614">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="911d4-614">- Content</span></span><br><span data-ttu-id="911d4-615">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-615">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="911d4-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="911d4-617">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="911d4-617">-</span></span><br><span data-ttu-id="911d4-618">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="911d4-618">
         -CompressedFile</span></span><br><span data-ttu-id="911d4-619">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-619">
         -DocumentEvents</span></span><br><span data-ttu-id="911d4-620">
         - File</span><span class="sxs-lookup"><span data-stu-id="911d4-620">
         - File</span></span><br><span data-ttu-id="911d4-621">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="911d4-621">
         -PdfFile</span></span><br><span data-ttu-id="911d4-622">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-622">
         - Selection</span></span><br><span data-ttu-id="911d4-623">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-623">
         - Settings</span></span><br><span data-ttu-id="911d4-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-624">
         -TextCoercion</span></span><br><span data-ttu-id="911d4-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-625">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-626">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="911d4-626">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="911d4-627">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="911d4-627">- Content</span></span><br><span data-ttu-id="911d4-628">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-628">
         - Taskpane</span></span><br><span data-ttu-id="911d4-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="911d4-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="911d4-631">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="911d4-631">-</span></span><br><span data-ttu-id="911d4-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="911d4-632">
         -CompressedFile</span></span><br><span data-ttu-id="911d4-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-633">
         -DocumentEvents</span></span><br><span data-ttu-id="911d4-634">
         - File</span><span class="sxs-lookup"><span data-stu-id="911d4-634">
         - File</span></span><br><span data-ttu-id="911d4-635">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-635">
         -ImageCoercion</span></span><br><span data-ttu-id="911d4-636">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="911d4-636">
         -PdfFile</span></span><br><span data-ttu-id="911d4-637">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-637">
         - Selection</span></span><br><span data-ttu-id="911d4-638">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-638">
         - Settings</span></span><br><span data-ttu-id="911d4-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-639">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-640">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="911d4-640">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="911d4-641">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="911d4-641">- Content</span></span><br><span data-ttu-id="911d4-642">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-642">
         - Taskpane</span></span><br><span data-ttu-id="911d4-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="911d4-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="911d4-645">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="911d4-645">-</span></span><br><span data-ttu-id="911d4-646">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="911d4-646">
         -CompressedFile</span></span><br><span data-ttu-id="911d4-647">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-647">
         -DocumentEvents</span></span><br><span data-ttu-id="911d4-648">
         - File</span><span class="sxs-lookup"><span data-stu-id="911d4-648">
         - File</span></span><br><span data-ttu-id="911d4-649">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-649">
         -ImageCoercion</span></span><br><span data-ttu-id="911d4-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="911d4-650">
         -PdfFile</span></span><br><span data-ttu-id="911d4-651">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-651">
         - Selection</span></span><br><span data-ttu-id="911d4-652">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-652">
         - Settings</span></span><br><span data-ttu-id="911d4-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-653">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="911d4-654">OneNote</span><span class="sxs-lookup"><span data-stu-id="911d4-654">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="911d4-655">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="911d4-655">Platform</span></span></th>
    <th><span data-ttu-id="911d4-656">拡張点</span><span class="sxs-lookup"><span data-stu-id="911d4-656">Extension points</span></span></th>
    <th><span data-ttu-id="911d4-657">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="911d4-657">API requirement sets</span></span></th>
    <th><span data-ttu-id="911d4-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="911d4-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-659">Office Online</span><span class="sxs-lookup"><span data-stu-id="911d4-659">Office Online</span></span></td>
    <td> <span data-ttu-id="911d4-660">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="911d4-660">- Content</span></span><br><span data-ttu-id="911d4-661">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-661">
         - Taskpane</span></span><br><span data-ttu-id="911d4-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="911d4-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="911d4-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="911d4-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="911d4-665">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="911d4-665">-DocumentEvents</span></span><br><span data-ttu-id="911d4-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-666">
         -HtmlCoercion</span></span><br><span data-ttu-id="911d4-667">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-667">
         -ImageCoercion</span></span><br><span data-ttu-id="911d4-668">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="911d4-668">
         - Settings</span></span><br><span data-ttu-id="911d4-669">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-669">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="911d4-670">Project</span><span class="sxs-lookup"><span data-stu-id="911d4-670">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="911d4-671">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="911d4-671">Platform</span></span></th>
    <th><span data-ttu-id="911d4-672">拡張点</span><span class="sxs-lookup"><span data-stu-id="911d4-672">Extension points</span></span></th>
    <th><span data-ttu-id="911d4-673">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="911d4-673">API requirement sets</span></span></th>
    <th><span data-ttu-id="911d4-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="911d4-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-675">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="911d4-675">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="911d4-676">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-676">- Taskpane</span></span></td>
    <td> <span data-ttu-id="911d4-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="911d4-678">- Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-678">- Selection</span></span><br><span data-ttu-id="911d4-679">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-679">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-680">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="911d4-680">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="911d4-681">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-681">- Taskpane</span></span></td>
    <td> <span data-ttu-id="911d4-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="911d4-683">- Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-683">- Selection</span></span><br><span data-ttu-id="911d4-684">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-684">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="911d4-685">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="911d4-685">Outlook 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="911d4-686">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="911d4-686">- Taskpane</span></span></td>
    <td> <span data-ttu-id="911d4-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="911d4-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="911d4-688">- Selection</span><span class="sxs-lookup"><span data-stu-id="911d4-688">- Selection</span></span><br><span data-ttu-id="911d4-689">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="911d4-689">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="911d4-690">関連項目</span><span class="sxs-lookup"><span data-stu-id="911d4-690">See also</span></span>

- [<span data-ttu-id="911d4-691">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="911d4-691">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="911d4-692">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="911d4-692">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="911d4-693">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="911d4-693">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="911d4-694">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="911d4-694">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
