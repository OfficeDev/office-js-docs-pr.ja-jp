---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、Word、Outlook、PowerPoint、OneNote、および Project のサポートされる要件セット。
ms.date: 03/07/2019
localization_priority: Priority
ms.openlocfilehash: 636c6290d8c67901beb195990593727485467460
ms.sourcegitcommit: 8e7b7b0cfb68b91a3a95585d094cf5f5ffd00178
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/09/2019
ms.locfileid: "30512882"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="94e3d-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="94e3d-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="94e3d-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="94e3d-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="94e3d-p102">MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。このバージョンには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="94e3d-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="94e3d-108">Excel</span><span class="sxs-lookup"><span data-stu-id="94e3d-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="94e3d-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="94e3d-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="94e3d-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="94e3d-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="94e3d-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="94e3d-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="94e3d-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="94e3d-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="94e3d-113">Office Online</span></span></td>
    <td> <span data-ttu-id="94e3d-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-114">- TaskPane</span></span><br><span data-ttu-id="94e3d-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="94e3d-115">
        - Content</span></span><br><span data-ttu-id="94e3d-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="94e3d-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="94e3d-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="94e3d-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="94e3d-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="94e3d-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="94e3d-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="94e3d-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="94e3d-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="94e3d-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="94e3d-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="94e3d-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-126">
        - BindingEvents</span></span><br><span data-ttu-id="94e3d-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-127">
        - CompressedFile</span></span><br><span data-ttu-id="94e3d-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-128">
        - DocumentEvents</span></span><br><span data-ttu-id="94e3d-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-129">
        - File</span></span><br><span data-ttu-id="94e3d-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-130">
        - MatrixBindings</span></span><br><span data-ttu-id="94e3d-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="94e3d-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-132">
        - Selection</span></span><br><span data-ttu-id="94e3d-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-133">
        - Settings</span></span><br><span data-ttu-id="94e3d-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-134">
        - TableBindings</span></span><br><span data-ttu-id="94e3d-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-135">
        - TableCoercion</span></span><br><span data-ttu-id="94e3d-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-136">
        - TextBindings</span></span><br><span data-ttu-id="94e3d-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-138">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="94e3d-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="94e3d-139">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-139">
        - TaskPane</span></span><br><span data-ttu-id="94e3d-140">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="94e3d-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="94e3d-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="94e3d-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="94e3d-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-142">
        - BindingEvents</span></span><br><span data-ttu-id="94e3d-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-143">
        - CompressedFile</span></span><br><span data-ttu-id="94e3d-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-144">
        - DocumentEvents</span></span><br><span data-ttu-id="94e3d-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-145">
        - File</span></span><br><span data-ttu-id="94e3d-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-146">
        - ImageCoercion</span></span><br><span data-ttu-id="94e3d-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-147">
        - MatrixBindings</span></span><br><span data-ttu-id="94e3d-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="94e3d-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-149">
        - Selection</span></span><br><span data-ttu-id="94e3d-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-150">
        - Settings</span></span><br><span data-ttu-id="94e3d-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-151">
        - TableBindings</span></span><br><span data-ttu-id="94e3d-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-152">
        - TableCoercion</span></span><br><span data-ttu-id="94e3d-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-153">
        - TextBindings</span></span><br><span data-ttu-id="94e3d-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-155">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="94e3d-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="94e3d-156">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-156">- TaskPane</span></span><br><span data-ttu-id="94e3d-157">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="94e3d-157">
        - Content</span></span></td>
    <td><span data-ttu-id="94e3d-158">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-158">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="94e3d-159">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="94e3d-159">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="94e3d-160">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-160">- BindingEvents</span></span><br><span data-ttu-id="94e3d-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-161">
        - CompressedFile</span></span><br><span data-ttu-id="94e3d-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-162">
        - DocumentEvents</span></span><br><span data-ttu-id="94e3d-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-163">
        - File</span></span><br><span data-ttu-id="94e3d-164">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-164">
        - ImageCoercion</span></span><br><span data-ttu-id="94e3d-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-165">
        - MatrixBindings</span></span><br><span data-ttu-id="94e3d-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="94e3d-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-167">
        - Selection</span></span><br><span data-ttu-id="94e3d-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-168">
        - Settings</span></span><br><span data-ttu-id="94e3d-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-169">
        - TableBindings</span></span><br><span data-ttu-id="94e3d-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-170">
        - TableCoercion</span></span><br><span data-ttu-id="94e3d-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-171">
        - TextBindings</span></span><br><span data-ttu-id="94e3d-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-173">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="94e3d-173">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="94e3d-174">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-174">- TaskPane</span></span><br><span data-ttu-id="94e3d-175">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="94e3d-175">
        - Content</span></span><br><span data-ttu-id="94e3d-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="94e3d-177">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-177">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="94e3d-178">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-178">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="94e3d-179">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-179">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="94e3d-180">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-180">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="94e3d-181">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-181">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="94e3d-182">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-182">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="94e3d-183">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-183">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="94e3d-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="94e3d-185">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-185">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="94e3d-186">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-186">- BindingEvents</span></span><br><span data-ttu-id="94e3d-187">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-187">
        - CompressedFile</span></span><br><span data-ttu-id="94e3d-188">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-188">
        - DocumentEvents</span></span><br><span data-ttu-id="94e3d-189">
        - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-189">
        - File</span></span><br><span data-ttu-id="94e3d-190">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-190">
        - ImageCoercion</span></span><br><span data-ttu-id="94e3d-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-191">
        - MatrixBindings</span></span><br><span data-ttu-id="94e3d-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="94e3d-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-193">
        - Selection</span></span><br><span data-ttu-id="94e3d-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-194">
        - Settings</span></span><br><span data-ttu-id="94e3d-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-195">
        - TableBindings</span></span><br><span data-ttu-id="94e3d-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-196">
        - TableCoercion</span></span><br><span data-ttu-id="94e3d-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-197">
        - TextBindings</span></span><br><span data-ttu-id="94e3d-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-199">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="94e3d-199">Office for iPad</span></span></td>
    <td><span data-ttu-id="94e3d-200">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-200">- TaskPane</span></span><br><span data-ttu-id="94e3d-201">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="94e3d-201">
        - Content</span></span></td>
    <td><span data-ttu-id="94e3d-202">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-202">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="94e3d-203">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-203">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="94e3d-204">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-204">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="94e3d-205">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-205">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="94e3d-206">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-206">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="94e3d-207">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-207">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="94e3d-208">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-208">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="94e3d-209">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-209">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="94e3d-210">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-210">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="94e3d-211">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-211">- BindingEvents</span></span><br><span data-ttu-id="94e3d-212">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-212">
        - CompressedFile</span></span><br><span data-ttu-id="94e3d-213">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-213">
        - DocumentEvents</span></span><br><span data-ttu-id="94e3d-214">
        - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-214">
        - File</span></span><br><span data-ttu-id="94e3d-215">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-215">
        - ImageCoercion</span></span><br><span data-ttu-id="94e3d-216">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-216">
        - MatrixBindings</span></span><br><span data-ttu-id="94e3d-217">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-217">
        - MatrixCoercion</span></span><br><span data-ttu-id="94e3d-218">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-218">
        - Selection</span></span><br><span data-ttu-id="94e3d-219">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-219">
        - Settings</span></span><br><span data-ttu-id="94e3d-220">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-220">
        - TableBindings</span></span><br><span data-ttu-id="94e3d-221">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-221">
        - TableCoercion</span></span><br><span data-ttu-id="94e3d-222">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-222">
        - TextBindings</span></span><br><span data-ttu-id="94e3d-223">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-223">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-224">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="94e3d-224">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="94e3d-225">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-225">- TaskPane</span></span><br><span data-ttu-id="94e3d-226">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="94e3d-226">
        - Content</span></span></td>
    <td><span data-ttu-id="94e3d-227">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-227">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="94e3d-228">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="94e3d-228">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="94e3d-229">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-229">- BindingEvents</span></span><br><span data-ttu-id="94e3d-230">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-230">
        - CompressedFile</span></span><br><span data-ttu-id="94e3d-231">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-231">
        - DocumentEvents</span></span><br><span data-ttu-id="94e3d-232">
        - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-232">
        - File</span></span><br><span data-ttu-id="94e3d-233">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-233">
        - ImageCoercion</span></span><br><span data-ttu-id="94e3d-234">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-234">
        - MatrixBindings</span></span><br><span data-ttu-id="94e3d-235">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-235">
        - MatrixCoercion</span></span><br><span data-ttu-id="94e3d-236">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-236">
        - PdfFile</span></span><br><span data-ttu-id="94e3d-237">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-237">
        - Selection</span></span><br><span data-ttu-id="94e3d-238">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-238">
        - Settings</span></span><br><span data-ttu-id="94e3d-239">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-239">
        - TableBindings</span></span><br><span data-ttu-id="94e3d-240">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-240">
        - TableCoercion</span></span><br><span data-ttu-id="94e3d-241">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-241">
        - TextBindings</span></span><br><span data-ttu-id="94e3d-242">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-242">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-243">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="94e3d-243">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="94e3d-244">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-244">- TaskPane</span></span><br><span data-ttu-id="94e3d-245">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="94e3d-245">
        - Content</span></span><br><span data-ttu-id="94e3d-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="94e3d-247">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-247">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="94e3d-248">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-248">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="94e3d-249">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-249">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="94e3d-250">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-250">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="94e3d-251">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-251">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="94e3d-252">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-252">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="94e3d-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="94e3d-254">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-254">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="94e3d-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="94e3d-256">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-256">- BindingEvents</span></span><br><span data-ttu-id="94e3d-257">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-257">
        - CompressedFile</span></span><br><span data-ttu-id="94e3d-258">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-258">
        - DocumentEvents</span></span><br><span data-ttu-id="94e3d-259">
        - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-259">
        - File</span></span><br><span data-ttu-id="94e3d-260">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-260">
        - ImageCoercion</span></span><br><span data-ttu-id="94e3d-261">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-261">
        - MatrixBindings</span></span><br><span data-ttu-id="94e3d-262">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-262">
        - MatrixCoercion</span></span><br><span data-ttu-id="94e3d-263">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-263">
        - PdfFile</span></span><br><span data-ttu-id="94e3d-264">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-264">
        - Selection</span></span><br><span data-ttu-id="94e3d-265">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-265">
        - Settings</span></span><br><span data-ttu-id="94e3d-266">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-266">
        - TableBindings</span></span><br><span data-ttu-id="94e3d-267">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-267">
        - TableCoercion</span></span><br><span data-ttu-id="94e3d-268">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-268">
        - TextBindings</span></span><br><span data-ttu-id="94e3d-269">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-269">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="94e3d-270">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="94e3d-270">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="94e3d-271">Outlook</span><span class="sxs-lookup"><span data-stu-id="94e3d-271">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="94e3d-272">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="94e3d-272">Platform</span></span></th>
    <th><span data-ttu-id="94e3d-273">拡張点</span><span class="sxs-lookup"><span data-stu-id="94e3d-273">Extension points</span></span></th>
    <th><span data-ttu-id="94e3d-274">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="94e3d-274">API requirement sets</span></span></th>
    <th><span data-ttu-id="94e3d-275"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="94e3d-275"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-276">Office Online</span><span class="sxs-lookup"><span data-stu-id="94e3d-276">Office Online</span></span></td>
    <td> <span data-ttu-id="94e3d-277">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="94e3d-277">- Mail Read</span></span><br><span data-ttu-id="94e3d-278">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="94e3d-278">
      - Mail Compose</span></span><br><span data-ttu-id="94e3d-279">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-279">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="94e3d-280">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-280">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="94e3d-281">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-281">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="94e3d-282">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-282">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="94e3d-283">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-283">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="94e3d-284">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-284">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="94e3d-285">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-285">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="94e3d-286">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-286">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="94e3d-287">利用不可</span><span class="sxs-lookup"><span data-stu-id="94e3d-287">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-288">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="94e3d-288">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="94e3d-289">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="94e3d-289">- Mail Read</span></span><br><span data-ttu-id="94e3d-290">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="94e3d-290">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="94e3d-291">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-291">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="94e3d-292">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-292">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="94e3d-293">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-293">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="94e3d-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="94e3d-295">利用不可</span><span class="sxs-lookup"><span data-stu-id="94e3d-295">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-296">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="94e3d-296">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="94e3d-297">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="94e3d-297">- Mail Read</span></span><br><span data-ttu-id="94e3d-298">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="94e3d-298">
      - Mail Compose</span></span><br><span data-ttu-id="94e3d-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="94e3d-300">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="94e3d-300">
      - Modules</span></span></td>
    <td> <span data-ttu-id="94e3d-301">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-301">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="94e3d-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="94e3d-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="94e3d-304">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-304">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="94e3d-305">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-305">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="94e3d-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="94e3d-307">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-307">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="94e3d-308">利用不可</span><span class="sxs-lookup"><span data-stu-id="94e3d-308">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-309">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="94e3d-309">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="94e3d-310">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="94e3d-310">- Mail Read</span></span><br><span data-ttu-id="94e3d-311">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="94e3d-311">
      - Mail Compose</span></span><br><span data-ttu-id="94e3d-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="94e3d-313">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="94e3d-313">
      - Modules</span></span></td>
    <td> <span data-ttu-id="94e3d-314">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-314">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="94e3d-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="94e3d-316">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-316">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="94e3d-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="94e3d-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="94e3d-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="94e3d-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="94e3d-321">利用不可</span><span class="sxs-lookup"><span data-stu-id="94e3d-321">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-322">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="94e3d-322">Office for iOS</span></span></td>
    <td> <span data-ttu-id="94e3d-323">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="94e3d-323">- Mail Read</span></span><br><span data-ttu-id="94e3d-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="94e3d-325">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-325">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="94e3d-326">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-326">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="94e3d-327">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-327">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="94e3d-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="94e3d-329">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-329">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="94e3d-330">利用不可</span><span class="sxs-lookup"><span data-stu-id="94e3d-330">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-331">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="94e3d-331">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="94e3d-332">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="94e3d-332">- Mail Read</span></span><br><span data-ttu-id="94e3d-333">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="94e3d-333">
      - Mail Compose</span></span><br><span data-ttu-id="94e3d-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="94e3d-335">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-335">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="94e3d-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="94e3d-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="94e3d-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="94e3d-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="94e3d-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="94e3d-341">利用不可</span><span class="sxs-lookup"><span data-stu-id="94e3d-341">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-342">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="94e3d-342">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="94e3d-343">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="94e3d-343">- Mail Read</span></span><br><span data-ttu-id="94e3d-344">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="94e3d-344">
      - Mail Compose</span></span><br><span data-ttu-id="94e3d-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="94e3d-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="94e3d-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="94e3d-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="94e3d-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="94e3d-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="94e3d-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="94e3d-352">利用不可</span><span class="sxs-lookup"><span data-stu-id="94e3d-352">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-353">Office for Android</span><span class="sxs-lookup"><span data-stu-id="94e3d-353">Office for Android</span></span></td>
    <td> <span data-ttu-id="94e3d-354">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="94e3d-354">- Mail Read</span></span><br><span data-ttu-id="94e3d-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="94e3d-356">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-356">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="94e3d-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="94e3d-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="94e3d-359">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-359">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="94e3d-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="94e3d-361">利用不可</span><span class="sxs-lookup"><span data-stu-id="94e3d-361">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="94e3d-362">Word</span><span class="sxs-lookup"><span data-stu-id="94e3d-362">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="94e3d-363">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="94e3d-363">Platform</span></span></th>
    <th><span data-ttu-id="94e3d-364">拡張点</span><span class="sxs-lookup"><span data-stu-id="94e3d-364">Extension points</span></span></th>
    <th><span data-ttu-id="94e3d-365">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="94e3d-365">API requirement sets</span></span></th>
    <th><span data-ttu-id="94e3d-366"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="94e3d-366"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-367">Office Online</span><span class="sxs-lookup"><span data-stu-id="94e3d-367">Office Online</span></span></td>
    <td> <span data-ttu-id="94e3d-368">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-368">- TaskPane</span></span><br><span data-ttu-id="94e3d-369">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-369">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="94e3d-370">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-370">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="94e3d-371">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-371">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="94e3d-372">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-372">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="94e3d-373">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-373">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="94e3d-374">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-374">- BindingEvents</span></span><br><span data-ttu-id="94e3d-375">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="94e3d-375">
         - CustomXmlParts</span></span><br><span data-ttu-id="94e3d-376">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-376">
         - DocumentEvents</span></span><br><span data-ttu-id="94e3d-377">
         - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-377">
         - File</span></span><br><span data-ttu-id="94e3d-378">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-378">
         - HtmlCoercion</span></span><br><span data-ttu-id="94e3d-379">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-379">
         - ImageCoercion</span></span><br><span data-ttu-id="94e3d-380">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-380">
         - MatrixBindings</span></span><br><span data-ttu-id="94e3d-381">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-381">
         - MatrixCoercion</span></span><br><span data-ttu-id="94e3d-382">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-382">
         - OoxmlCoercion</span></span><br><span data-ttu-id="94e3d-383">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-383">
         - PdfFile</span></span><br><span data-ttu-id="94e3d-384">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-384">
         - Selection</span></span><br><span data-ttu-id="94e3d-385">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-385">
         - Settings</span></span><br><span data-ttu-id="94e3d-386">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-386">
         - TableBindings</span></span><br><span data-ttu-id="94e3d-387">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-387">
         - TableCoercion</span></span><br><span data-ttu-id="94e3d-388">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-388">
         - TextBindings</span></span><br><span data-ttu-id="94e3d-389">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-389">
         - TextCoercion</span></span><br><span data-ttu-id="94e3d-390">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-390">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-391">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="94e3d-391">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="94e3d-392">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-392">- TaskPane</span></span></td>
    <td> <span data-ttu-id="94e3d-393">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="94e3d-393">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="94e3d-394">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-394">- BindingEvents</span></span><br><span data-ttu-id="94e3d-395">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-395">
         - CompressedFile</span></span><br><span data-ttu-id="94e3d-396">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="94e3d-396">
         - CustomXmlParts</span></span><br><span data-ttu-id="94e3d-397">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-397">
         - DocumentEvents</span></span><br><span data-ttu-id="94e3d-398">
         - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-398">
         - File</span></span><br><span data-ttu-id="94e3d-399">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-399">
         - HtmlCoercion</span></span><br><span data-ttu-id="94e3d-400">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-400">
         - ImageCoercion</span></span><br><span data-ttu-id="94e3d-401">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-401">
         - MatrixBindings</span></span><br><span data-ttu-id="94e3d-402">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-402">
         - MatrixCoercion</span></span><br><span data-ttu-id="94e3d-403">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-403">
         - OoxmlCoercion</span></span><br><span data-ttu-id="94e3d-404">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-404">
         - PdfFile</span></span><br><span data-ttu-id="94e3d-405">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-405">
         - Selection</span></span><br><span data-ttu-id="94e3d-406">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-406">
         - Settings</span></span><br><span data-ttu-id="94e3d-407">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-407">
         - TableBindings</span></span><br><span data-ttu-id="94e3d-408">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-408">
         - TableCoercion</span></span><br><span data-ttu-id="94e3d-409">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-409">
         - TextBindings</span></span><br><span data-ttu-id="94e3d-410">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-410">
         - TextCoercion</span></span><br><span data-ttu-id="94e3d-411">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-411">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-412">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="94e3d-412">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="94e3d-413">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-413">- TaskPane</span></span></td>
    <td> <span data-ttu-id="94e3d-414">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-414">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="94e3d-415">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="94e3d-415">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="94e3d-416">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-416">- BindingEvents</span></span><br><span data-ttu-id="94e3d-417">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-417">
         - CompressedFile</span></span><br><span data-ttu-id="94e3d-418">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="94e3d-418">
         - CustomXmlParts</span></span><br><span data-ttu-id="94e3d-419">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-419">
         - DocumentEvents</span></span><br><span data-ttu-id="94e3d-420">
         - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-420">
         - File</span></span><br><span data-ttu-id="94e3d-421">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-421">
         - HtmlCoercion</span></span><br><span data-ttu-id="94e3d-422">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-422">
         - ImageCoercion</span></span><br><span data-ttu-id="94e3d-423">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-423">
         - MatrixBindings</span></span><br><span data-ttu-id="94e3d-424">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-424">
         - MatrixCoercion</span></span><br><span data-ttu-id="94e3d-425">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-425">
         - OoxmlCoercion</span></span><br><span data-ttu-id="94e3d-426">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-426">
         - PdfFile</span></span><br><span data-ttu-id="94e3d-427">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-427">
         - Selection</span></span><br><span data-ttu-id="94e3d-428">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-428">
         - Settings</span></span><br><span data-ttu-id="94e3d-429">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-429">
         - TableBindings</span></span><br><span data-ttu-id="94e3d-430">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-430">
         - TableCoercion</span></span><br><span data-ttu-id="94e3d-431">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-431">
         - TextBindings</span></span><br><span data-ttu-id="94e3d-432">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-432">
         - TextCoercion</span></span><br><span data-ttu-id="94e3d-433">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-433">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-434">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="94e3d-434">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="94e3d-435">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-435">- TaskPane</span></span><br><span data-ttu-id="94e3d-436">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-436">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="94e3d-437">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-437">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="94e3d-438">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-438">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="94e3d-439">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-439">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="94e3d-440">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-440">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="94e3d-441">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-441">- BindingEvents</span></span><br><span data-ttu-id="94e3d-442">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-442">
         - CompressedFile</span></span><br><span data-ttu-id="94e3d-443">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="94e3d-443">
         - CustomXmlParts</span></span><br><span data-ttu-id="94e3d-444">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-444">
         - DocumentEvents</span></span><br><span data-ttu-id="94e3d-445">
         - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-445">
         - File</span></span><br><span data-ttu-id="94e3d-446">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-446">
         - HtmlCoercion</span></span><br><span data-ttu-id="94e3d-447">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-447">
         - ImageCoercion</span></span><br><span data-ttu-id="94e3d-448">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-448">
         - MatrixBindings</span></span><br><span data-ttu-id="94e3d-449">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-449">
         - MatrixCoercion</span></span><br><span data-ttu-id="94e3d-450">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-450">
         - OoxmlCoercion</span></span><br><span data-ttu-id="94e3d-451">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-451">
         - PdfFile</span></span><br><span data-ttu-id="94e3d-452">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-452">
         - Selection</span></span><br><span data-ttu-id="94e3d-453">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-453">
         - Settings</span></span><br><span data-ttu-id="94e3d-454">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-454">
         - TableBindings</span></span><br><span data-ttu-id="94e3d-455">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-455">
         - TableCoercion</span></span><br><span data-ttu-id="94e3d-456">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-456">
         - TextBindings</span></span><br><span data-ttu-id="94e3d-457">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-457">
         - TextCoercion</span></span><br><span data-ttu-id="94e3d-458">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-458">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-459">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="94e3d-459">Office for iPad</span></span></td>
    <td> <span data-ttu-id="94e3d-460">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-460">- TaskPane</span></span></td>
    <td> <span data-ttu-id="94e3d-461">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-461">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="94e3d-462">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-462">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="94e3d-463">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-463">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="94e3d-464">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="94e3d-464">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="94e3d-465">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-465">- BindingEvents</span></span><br><span data-ttu-id="94e3d-466">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-466">
         - CompressedFile</span></span><br><span data-ttu-id="94e3d-467">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="94e3d-467">
         - CustomXmlParts</span></span><br><span data-ttu-id="94e3d-468">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-468">
         - DocumentEvents</span></span><br><span data-ttu-id="94e3d-469">
         - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-469">
         - File</span></span><br><span data-ttu-id="94e3d-470">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-470">
         - HtmlCoercion</span></span><br><span data-ttu-id="94e3d-471">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-471">
         - ImageCoercion</span></span><br><span data-ttu-id="94e3d-472">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-472">
         - MatrixBindings</span></span><br><span data-ttu-id="94e3d-473">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-473">
         - MatrixCoercion</span></span><br><span data-ttu-id="94e3d-474">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-474">
         - OoxmlCoercion</span></span><br><span data-ttu-id="94e3d-475">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-475">
         - PdfFile</span></span><br><span data-ttu-id="94e3d-476">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-476">
         - Selection</span></span><br><span data-ttu-id="94e3d-477">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-477">
         - Settings</span></span><br><span data-ttu-id="94e3d-478">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-478">
         - TableBindings</span></span><br><span data-ttu-id="94e3d-479">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-479">
         - TableCoercion</span></span><br><span data-ttu-id="94e3d-480">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-480">
         - TextBindings</span></span><br><span data-ttu-id="94e3d-481">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-481">
         - TextCoercion</span></span><br><span data-ttu-id="94e3d-482">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-482">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-483">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="94e3d-483">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="94e3d-484">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-484">- TaskPane</span></span></td>
    <td> <span data-ttu-id="94e3d-485">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-485">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="94e3d-486">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="94e3d-486">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="94e3d-487">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-487">- BindingEvents</span></span><br><span data-ttu-id="94e3d-488">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-488">
         - CompressedFile</span></span><br><span data-ttu-id="94e3d-489">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="94e3d-489">
         - CustomXmlParts</span></span><br><span data-ttu-id="94e3d-490">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-490">
         - DocumentEvents</span></span><br><span data-ttu-id="94e3d-491">
         - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-491">
         - File</span></span><br><span data-ttu-id="94e3d-492">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-492">
         - HtmlCoercion</span></span><br><span data-ttu-id="94e3d-493">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-493">
         - ImageCoercion</span></span><br><span data-ttu-id="94e3d-494">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-494">
         - MatrixBindings</span></span><br><span data-ttu-id="94e3d-495">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-495">
         - MatrixCoercion</span></span><br><span data-ttu-id="94e3d-496">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-496">
         - OoxmlCoercion</span></span><br><span data-ttu-id="94e3d-497">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-497">
         - PdfFile</span></span><br><span data-ttu-id="94e3d-498">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-498">
         - Selection</span></span><br><span data-ttu-id="94e3d-499">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-499">
         - Settings</span></span><br><span data-ttu-id="94e3d-500">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-500">
         - TableBindings</span></span><br><span data-ttu-id="94e3d-501">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-501">
         - TableCoercion</span></span><br><span data-ttu-id="94e3d-502">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-502">
         - TextBindings</span></span><br><span data-ttu-id="94e3d-503">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-503">
         - TextCoercion</span></span><br><span data-ttu-id="94e3d-504">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-504">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-505">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="94e3d-505">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="94e3d-506">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-506">- TaskPane</span></span><br><span data-ttu-id="94e3d-507">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-507">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="94e3d-508">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-508">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="94e3d-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="94e3d-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="94e3d-511">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="94e3d-511">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="94e3d-512">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-512">- BindingEvents</span></span><br><span data-ttu-id="94e3d-513">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-513">
         - CompressedFile</span></span><br><span data-ttu-id="94e3d-514">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="94e3d-514">
         - CustomXmlParts</span></span><br><span data-ttu-id="94e3d-515">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-515">
         - DocumentEvents</span></span><br><span data-ttu-id="94e3d-516">
         - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-516">
         - File</span></span><br><span data-ttu-id="94e3d-517">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-517">
         - HtmlCoercion</span></span><br><span data-ttu-id="94e3d-518">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-518">
         - ImageCoercion</span></span><br><span data-ttu-id="94e3d-519">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-519">
         - MatrixBindings</span></span><br><span data-ttu-id="94e3d-520">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-520">
         - MatrixCoercion</span></span><br><span data-ttu-id="94e3d-521">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-521">
         - OoxmlCoercion</span></span><br><span data-ttu-id="94e3d-522">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-522">
         - PdfFile</span></span><br><span data-ttu-id="94e3d-523">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-523">
         - Selection</span></span><br><span data-ttu-id="94e3d-524">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-524">
         - Settings</span></span><br><span data-ttu-id="94e3d-525">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-525">
         - TableBindings</span></span><br><span data-ttu-id="94e3d-526">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-526">
         - TableCoercion</span></span><br><span data-ttu-id="94e3d-527">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="94e3d-527">
         - TextBindings</span></span><br><span data-ttu-id="94e3d-528">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-528">
         - TextCoercion</span></span><br><span data-ttu-id="94e3d-529">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-529">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="94e3d-530">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="94e3d-530">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="94e3d-531">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="94e3d-531">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="94e3d-532">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="94e3d-532">Platform</span></span></th>
    <th><span data-ttu-id="94e3d-533">拡張点</span><span class="sxs-lookup"><span data-stu-id="94e3d-533">Extension points</span></span></th>
    <th><span data-ttu-id="94e3d-534">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="94e3d-534">API requirement sets</span></span></th>
    <th><span data-ttu-id="94e3d-535"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="94e3d-535"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-536">Office Online</span><span class="sxs-lookup"><span data-stu-id="94e3d-536">Office Online</span></span></td>
    <td> <span data-ttu-id="94e3d-537">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="94e3d-537">- Content</span></span><br><span data-ttu-id="94e3d-538">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-538">
         - TaskPane</span></span><br><span data-ttu-id="94e3d-539">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-539">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="94e3d-540">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-540">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="94e3d-541">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="94e3d-541">- ActiveView</span></span><br><span data-ttu-id="94e3d-542">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-542">
         - CompressedFile</span></span><br><span data-ttu-id="94e3d-543">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-543">
         - DocumentEvents</span></span><br><span data-ttu-id="94e3d-544">
         - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-544">
         - File</span></span><br><span data-ttu-id="94e3d-545">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-545">
         - ImageCoercion</span></span><br><span data-ttu-id="94e3d-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-546">
         - PdfFile</span></span><br><span data-ttu-id="94e3d-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-547">
         - Selection</span></span><br><span data-ttu-id="94e3d-548">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-548">
         - Settings</span></span><br><span data-ttu-id="94e3d-549">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-549">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-550">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="94e3d-550">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="94e3d-551">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="94e3d-551">- Content</span></span><br><span data-ttu-id="94e3d-552">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-552">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="94e3d-553">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="94e3d-553">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="94e3d-554">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="94e3d-554">- ActiveView</span></span><br><span data-ttu-id="94e3d-555">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-555">
         - CompressedFile</span></span><br><span data-ttu-id="94e3d-556">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-556">
         - DocumentEvents</span></span><br><span data-ttu-id="94e3d-557">
         - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-557">
         - File</span></span><br><span data-ttu-id="94e3d-558">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-558">
         - ImageCoercion</span></span><br><span data-ttu-id="94e3d-559">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-559">
         - PdfFile</span></span><br><span data-ttu-id="94e3d-560">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-560">
         - Selection</span></span><br><span data-ttu-id="94e3d-561">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-561">
         - Settings</span></span><br><span data-ttu-id="94e3d-562">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-562">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-563">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="94e3d-563">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="94e3d-564">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="94e3d-564">- Content</span></span><br><span data-ttu-id="94e3d-565">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-565">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="94e3d-566">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="94e3d-566">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="94e3d-567">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="94e3d-567">- ActiveView</span></span><br><span data-ttu-id="94e3d-568">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-568">
         - CompressedFile</span></span><br><span data-ttu-id="94e3d-569">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-569">
         - DocumentEvents</span></span><br><span data-ttu-id="94e3d-570">
         - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-570">
         - File</span></span><br><span data-ttu-id="94e3d-571">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-571">
         - ImageCoercion</span></span><br><span data-ttu-id="94e3d-572">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-572">
         - PdfFile</span></span><br><span data-ttu-id="94e3d-573">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-573">
         - Selection</span></span><br><span data-ttu-id="94e3d-574">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-574">
         - Settings</span></span><br><span data-ttu-id="94e3d-575">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-575">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-576">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="94e3d-576">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="94e3d-577">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="94e3d-577">- Content</span></span><br><span data-ttu-id="94e3d-578">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-578">
         - TaskPane</span></span><br><span data-ttu-id="94e3d-579">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-579">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="94e3d-580">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-580">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="94e3d-581">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="94e3d-581">- ActiveView</span></span><br><span data-ttu-id="94e3d-582">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-582">
         - CompressedFile</span></span><br><span data-ttu-id="94e3d-583">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-583">
         - DocumentEvents</span></span><br><span data-ttu-id="94e3d-584">
         - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-584">
         - File</span></span><br><span data-ttu-id="94e3d-585">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-585">
         - ImageCoercion</span></span><br><span data-ttu-id="94e3d-586">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-586">
         - PdfFile</span></span><br><span data-ttu-id="94e3d-587">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-587">
         - Selection</span></span><br><span data-ttu-id="94e3d-588">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-588">
         - Settings</span></span><br><span data-ttu-id="94e3d-589">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-589">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-590">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="94e3d-590">Office for iPad</span></span></td>
    <td> <span data-ttu-id="94e3d-591">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="94e3d-591">- Content</span></span><br><span data-ttu-id="94e3d-592">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-592">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="94e3d-593">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-593">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="94e3d-594">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="94e3d-594">- ActiveView</span></span><br><span data-ttu-id="94e3d-595">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-595">
         - CompressedFile</span></span><br><span data-ttu-id="94e3d-596">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-596">
         - DocumentEvents</span></span><br><span data-ttu-id="94e3d-597">
         - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-597">
         - File</span></span><br><span data-ttu-id="94e3d-598">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-598">
         - PdfFile</span></span><br><span data-ttu-id="94e3d-599">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-599">
         - Selection</span></span><br><span data-ttu-id="94e3d-600">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-600">
         - Settings</span></span><br><span data-ttu-id="94e3d-601">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-601">
         - TextCoercion</span></span><br><span data-ttu-id="94e3d-602">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-602">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-603">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="94e3d-603">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="94e3d-604">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="94e3d-604">- Content</span></span><br><span data-ttu-id="94e3d-605">
         - 作業ウィンドウ/td></span><span class="sxs-lookup"><span data-stu-id="94e3d-605">
         - TaskPane/td></span></span> <td> <span data-ttu-id="94e3d-606">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="94e3d-606">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="94e3d-607">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="94e3d-607">- ActiveView</span></span><br><span data-ttu-id="94e3d-608">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-608">
         - CompressedFile</span></span><br><span data-ttu-id="94e3d-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-609">
         - DocumentEvents</span></span><br><span data-ttu-id="94e3d-610">
         - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-610">
         - File</span></span><br><span data-ttu-id="94e3d-611">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-611">
         - ImageCoercion</span></span><br><span data-ttu-id="94e3d-612">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-612">
         - PdfFile</span></span><br><span data-ttu-id="94e3d-613">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-613">
         - Selection</span></span><br><span data-ttu-id="94e3d-614">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-614">
         - Settings</span></span><br><span data-ttu-id="94e3d-615">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-615">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-616">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="94e3d-616">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="94e3d-617">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="94e3d-617">- Content</span></span><br><span data-ttu-id="94e3d-618">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-618">
         - TaskPane</span></span><br><span data-ttu-id="94e3d-619">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-619">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="94e3d-620">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-620">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="94e3d-621">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="94e3d-621">- ActiveView</span></span><br><span data-ttu-id="94e3d-622">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-622">
         - CompressedFile</span></span><br><span data-ttu-id="94e3d-623">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-623">
         - DocumentEvents</span></span><br><span data-ttu-id="94e3d-624">
         - File</span><span class="sxs-lookup"><span data-stu-id="94e3d-624">
         - File</span></span><br><span data-ttu-id="94e3d-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-625">
         - ImageCoercion</span></span><br><span data-ttu-id="94e3d-626">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="94e3d-626">
         - PdfFile</span></span><br><span data-ttu-id="94e3d-627">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-627">
         - Selection</span></span><br><span data-ttu-id="94e3d-628">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-628">
         - Settings</span></span><br><span data-ttu-id="94e3d-629">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-629">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="94e3d-630">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="94e3d-630">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="94e3d-631">OneNote</span><span class="sxs-lookup"><span data-stu-id="94e3d-631">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="94e3d-632">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="94e3d-632">Platform</span></span></th>
    <th><span data-ttu-id="94e3d-633">拡張点</span><span class="sxs-lookup"><span data-stu-id="94e3d-633">Extension points</span></span></th>
    <th><span data-ttu-id="94e3d-634">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="94e3d-634">API requirement sets</span></span></th>
    <th><span data-ttu-id="94e3d-635"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="94e3d-635"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-636">Office Online</span><span class="sxs-lookup"><span data-stu-id="94e3d-636">Office Online</span></span></td>
    <td> <span data-ttu-id="94e3d-637">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="94e3d-637">- Content</span></span><br><span data-ttu-id="94e3d-638">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-638">
         - TaskPane</span></span><br><span data-ttu-id="94e3d-639">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-639">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="94e3d-640">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-640">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="94e3d-641">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-641">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="94e3d-642">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="94e3d-642">- DocumentEvents</span></span><br><span data-ttu-id="94e3d-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="94e3d-644">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-644">
         - ImageCoercion</span></span><br><span data-ttu-id="94e3d-645">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="94e3d-645">
         - Settings</span></span><br><span data-ttu-id="94e3d-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-646">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="94e3d-647">Project</span><span class="sxs-lookup"><span data-stu-id="94e3d-647">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="94e3d-648">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="94e3d-648">Platform</span></span></th>
    <th><span data-ttu-id="94e3d-649">拡張点</span><span class="sxs-lookup"><span data-stu-id="94e3d-649">Extension points</span></span></th>
    <th><span data-ttu-id="94e3d-650">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="94e3d-650">API requirement sets</span></span></th>
    <th><span data-ttu-id="94e3d-651"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="94e3d-651"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-652">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="94e3d-652">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="94e3d-653">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-653">- TaskPane</span></span></td>
    <td> <span data-ttu-id="94e3d-654">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-654">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="94e3d-655">- Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-655">- Selection</span></span><br><span data-ttu-id="94e3d-656">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-656">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-657">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="94e3d-657">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="94e3d-658">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-658">- TaskPane</span></span></td>
    <td> <span data-ttu-id="94e3d-659">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-659">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="94e3d-660">- Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-660">- Selection</span></span><br><span data-ttu-id="94e3d-661">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-661">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="94e3d-662">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="94e3d-662">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="94e3d-663">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="94e3d-663">- TaskPane</span></span></td>
    <td> <span data-ttu-id="94e3d-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="94e3d-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="94e3d-665">- Selection</span><span class="sxs-lookup"><span data-stu-id="94e3d-665">- Selection</span></span><br><span data-ttu-id="94e3d-666">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="94e3d-666">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="94e3d-667">関連項目</span><span class="sxs-lookup"><span data-stu-id="94e3d-667">See also</span></span>

- [<span data-ttu-id="94e3d-668">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="94e3d-668">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="94e3d-669">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="94e3d-669">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="94e3d-670">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="94e3d-670">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="94e3d-671">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="94e3d-671">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
