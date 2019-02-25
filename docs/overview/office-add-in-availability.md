---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、Word、Outlook、PowerPoint、OneNote、および Project のサポートされる要件セット。
ms.date: 02/20/2019
localization_priority: Priority
ms.openlocfilehash: a3e9c508a5bae0e7eb660458835b9242d0602818
ms.sourcegitcommit: 8e20e7663be2aaa0f7a5436a965324d171bc667d
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/22/2019
ms.locfileid: "30199614"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="80a41-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="80a41-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="80a41-104">Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。</span><span class="sxs-lookup"><span data-stu-id="80a41-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="80a41-105">次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、各 Office アプリケーションで現在サポートされている共通 API が記載されています。</span><span class="sxs-lookup"><span data-stu-id="80a41-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="80a41-p102">MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。このバージョンには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="80a41-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="80a41-108">Excel</span><span class="sxs-lookup"><span data-stu-id="80a41-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="80a41-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="80a41-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="80a41-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="80a41-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="80a41-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="80a41-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="80a41-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="80a41-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="80a41-113">Office Online</span></span></td>
    <td> <span data-ttu-id="80a41-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-114">- TaskPane</span></span><br><span data-ttu-id="80a41-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="80a41-115">
        - Content</span></span><br><span data-ttu-id="80a41-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="80a41-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="80a41-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="80a41-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="80a41-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="80a41-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="80a41-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="80a41-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="80a41-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="80a41-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="80a41-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="80a41-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="80a41-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="80a41-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="80a41-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="80a41-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="80a41-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="80a41-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="80a41-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-126">
        - BindingEvents</span></span><br><span data-ttu-id="80a41-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="80a41-127">
        - CompressedFile</span></span><br><span data-ttu-id="80a41-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-128">
        - DocumentEvents</span></span><br><span data-ttu-id="80a41-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="80a41-129">
        - File</span></span><br><span data-ttu-id="80a41-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-130">
        - MatrixBindings</span></span><br><span data-ttu-id="80a41-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="80a41-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-132">
        - Selection</span></span><br><span data-ttu-id="80a41-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-133">
        - Settings</span></span><br><span data-ttu-id="80a41-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-134">
        - TableBindings</span></span><br><span data-ttu-id="80a41-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-135">
        - TableCoercion</span></span><br><span data-ttu-id="80a41-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-136">
        - TextBindings</span></span><br><span data-ttu-id="80a41-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-138">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="80a41-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="80a41-139">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-139">
        - TaskPane</span></span><br><span data-ttu-id="80a41-140">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="80a41-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="80a41-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="80a41-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="80a41-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-142">
        - BindingEvents</span></span><br><span data-ttu-id="80a41-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="80a41-143">
        - CompressedFile</span></span><br><span data-ttu-id="80a41-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-144">
        - DocumentEvents</span></span><br><span data-ttu-id="80a41-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="80a41-145">
        - File</span></span><br><span data-ttu-id="80a41-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-146">
        - ImageCoercion</span></span><br><span data-ttu-id="80a41-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-147">
        - MatrixBindings</span></span><br><span data-ttu-id="80a41-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="80a41-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-149">
        - Selection</span></span><br><span data-ttu-id="80a41-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-150">
        - Settings</span></span><br><span data-ttu-id="80a41-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-151">
        - TableBindings</span></span><br><span data-ttu-id="80a41-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-152">
        - TableCoercion</span></span><br><span data-ttu-id="80a41-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-153">
        - TextBindings</span></span><br><span data-ttu-id="80a41-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-155">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="80a41-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="80a41-156">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-156">- TaskPane</span></span><br><span data-ttu-id="80a41-157">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="80a41-157">
        - Content</span></span></td>
    <td><span data-ttu-id="80a41-158">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-158">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="80a41-159">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="80a41-159">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="80a41-160">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-160">- BindingEvents</span></span><br><span data-ttu-id="80a41-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="80a41-161">
        - CompressedFile</span></span><br><span data-ttu-id="80a41-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-162">
        - DocumentEvents</span></span><br><span data-ttu-id="80a41-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="80a41-163">
        - File</span></span><br><span data-ttu-id="80a41-164">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-164">
        - ImageCoercion</span></span><br><span data-ttu-id="80a41-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-165">
        - MatrixBindings</span></span><br><span data-ttu-id="80a41-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="80a41-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-167">
        - Selection</span></span><br><span data-ttu-id="80a41-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-168">
        - Settings</span></span><br><span data-ttu-id="80a41-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-169">
        - TableBindings</span></span><br><span data-ttu-id="80a41-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-170">
        - TableCoercion</span></span><br><span data-ttu-id="80a41-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-171">
        - TextBindings</span></span><br><span data-ttu-id="80a41-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-173">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="80a41-173">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="80a41-174">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-174">- TaskPane</span></span><br><span data-ttu-id="80a41-175">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="80a41-175">
        - Content</span></span><br><span data-ttu-id="80a41-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="80a41-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="80a41-177">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-177">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="80a41-178">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="80a41-178">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="80a41-179">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="80a41-179">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="80a41-180">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="80a41-180">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="80a41-181">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="80a41-181">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="80a41-182">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="80a41-182">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="80a41-183">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="80a41-183">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="80a41-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="80a41-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="80a41-185">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-185">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="80a41-186">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-186">- BindingEvents</span></span><br><span data-ttu-id="80a41-187">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="80a41-187">
        - CompressedFile</span></span><br><span data-ttu-id="80a41-188">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-188">
        - DocumentEvents</span></span><br><span data-ttu-id="80a41-189">
        - File</span><span class="sxs-lookup"><span data-stu-id="80a41-189">
        - File</span></span><br><span data-ttu-id="80a41-190">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-190">
        - ImageCoercion</span></span><br><span data-ttu-id="80a41-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-191">
        - MatrixBindings</span></span><br><span data-ttu-id="80a41-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="80a41-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-193">
        - Selection</span></span><br><span data-ttu-id="80a41-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-194">
        - Settings</span></span><br><span data-ttu-id="80a41-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-195">
        - TableBindings</span></span><br><span data-ttu-id="80a41-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-196">
        - TableCoercion</span></span><br><span data-ttu-id="80a41-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-197">
        - TextBindings</span></span><br><span data-ttu-id="80a41-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-199">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="80a41-199">Office for iPad</span></span></td>
    <td><span data-ttu-id="80a41-200">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-200">- TaskPane</span></span><br><span data-ttu-id="80a41-201">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="80a41-201">
        - Content</span></span></td>
    <td><span data-ttu-id="80a41-202">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-202">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="80a41-203">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="80a41-203">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="80a41-204">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="80a41-204">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="80a41-205">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="80a41-205">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="80a41-206">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="80a41-206">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="80a41-207">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="80a41-207">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="80a41-208">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="80a41-208">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="80a41-209">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="80a41-209">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="80a41-210">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-210">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="80a41-211">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-211">- BindingEvents</span></span><br><span data-ttu-id="80a41-212">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="80a41-212">
        - CompressedFile</span></span><br><span data-ttu-id="80a41-213">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-213">
        - DocumentEvents</span></span><br><span data-ttu-id="80a41-214">
        - File</span><span class="sxs-lookup"><span data-stu-id="80a41-214">
        - File</span></span><br><span data-ttu-id="80a41-215">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-215">
        - ImageCoercion</span></span><br><span data-ttu-id="80a41-216">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-216">
        - MatrixBindings</span></span><br><span data-ttu-id="80a41-217">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-217">
        - MatrixCoercion</span></span><br><span data-ttu-id="80a41-218">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-218">
        - Selection</span></span><br><span data-ttu-id="80a41-219">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-219">
        - Settings</span></span><br><span data-ttu-id="80a41-220">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-220">
        - TableBindings</span></span><br><span data-ttu-id="80a41-221">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-221">
        - TableCoercion</span></span><br><span data-ttu-id="80a41-222">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-222">
        - TextBindings</span></span><br><span data-ttu-id="80a41-223">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-223">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-224">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="80a41-224">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="80a41-225">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-225">- TaskPane</span></span><br><span data-ttu-id="80a41-226">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="80a41-226">
        - Content</span></span></td>
    <td><span data-ttu-id="80a41-227">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-227">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="80a41-228">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="80a41-228">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="80a41-229">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-229">- BindingEvents</span></span><br><span data-ttu-id="80a41-230">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="80a41-230">
        - CompressedFile</span></span><br><span data-ttu-id="80a41-231">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-231">
        - DocumentEvents</span></span><br><span data-ttu-id="80a41-232">
        - File</span><span class="sxs-lookup"><span data-stu-id="80a41-232">
        - File</span></span><br><span data-ttu-id="80a41-233">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-233">
        - ImageCoercion</span></span><br><span data-ttu-id="80a41-234">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-234">
        - MatrixBindings</span></span><br><span data-ttu-id="80a41-235">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-235">
        - MatrixCoercion</span></span><br><span data-ttu-id="80a41-236">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="80a41-236">
        - PdfFile</span></span><br><span data-ttu-id="80a41-237">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-237">
        - Selection</span></span><br><span data-ttu-id="80a41-238">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-238">
        - Settings</span></span><br><span data-ttu-id="80a41-239">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-239">
        - TableBindings</span></span><br><span data-ttu-id="80a41-240">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-240">
        - TableCoercion</span></span><br><span data-ttu-id="80a41-241">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-241">
        - TextBindings</span></span><br><span data-ttu-id="80a41-242">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-242">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-243">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="80a41-243">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="80a41-244">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-244">- TaskPane</span></span><br><span data-ttu-id="80a41-245">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="80a41-245">
        - Content</span></span><br><span data-ttu-id="80a41-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="80a41-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="80a41-247">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-247">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="80a41-248">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="80a41-248">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="80a41-249">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="80a41-249">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="80a41-250">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="80a41-250">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="80a41-251">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="80a41-251">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="80a41-252">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="80a41-252">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="80a41-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="80a41-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="80a41-254">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="80a41-254">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="80a41-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="80a41-256">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-256">- BindingEvents</span></span><br><span data-ttu-id="80a41-257">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="80a41-257">
        - CompressedFile</span></span><br><span data-ttu-id="80a41-258">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-258">
        - DocumentEvents</span></span><br><span data-ttu-id="80a41-259">
        - File</span><span class="sxs-lookup"><span data-stu-id="80a41-259">
        - File</span></span><br><span data-ttu-id="80a41-260">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-260">
        - ImageCoercion</span></span><br><span data-ttu-id="80a41-261">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-261">
        - MatrixBindings</span></span><br><span data-ttu-id="80a41-262">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-262">
        - MatrixCoercion</span></span><br><span data-ttu-id="80a41-263">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="80a41-263">
        - PdfFile</span></span><br><span data-ttu-id="80a41-264">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-264">
        - Selection</span></span><br><span data-ttu-id="80a41-265">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-265">
        - Settings</span></span><br><span data-ttu-id="80a41-266">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-266">
        - TableBindings</span></span><br><span data-ttu-id="80a41-267">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-267">
        - TableCoercion</span></span><br><span data-ttu-id="80a41-268">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-268">
        - TextBindings</span></span><br><span data-ttu-id="80a41-269">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-269">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="80a41-270">Outlook</span><span class="sxs-lookup"><span data-stu-id="80a41-270">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="80a41-271">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="80a41-271">Platform</span></span></th>
    <th><span data-ttu-id="80a41-272">拡張点</span><span class="sxs-lookup"><span data-stu-id="80a41-272">Extension points</span></span></th>
    <th><span data-ttu-id="80a41-273">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="80a41-273">API requirement sets</span></span></th>
    <th><span data-ttu-id="80a41-274"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="80a41-274"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-275">Office Online</span><span class="sxs-lookup"><span data-stu-id="80a41-275">Office Online</span></span></td>
    <td> <span data-ttu-id="80a41-276">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="80a41-276">- Mail Read</span></span><br><span data-ttu-id="80a41-277">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="80a41-277">
      - Mail Compose</span></span><br><span data-ttu-id="80a41-278">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="80a41-278">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="80a41-279">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-279">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="80a41-280">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="80a41-280">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="80a41-281">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="80a41-281">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="80a41-282">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="80a41-282">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="80a41-283">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="80a41-283">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="80a41-284">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="80a41-284">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="80a41-285">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="80a41-285">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="80a41-286">利用不可</span><span class="sxs-lookup"><span data-stu-id="80a41-286">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-287">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="80a41-287">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="80a41-288">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="80a41-288">- Mail Read</span></span><br><span data-ttu-id="80a41-289">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="80a41-289">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="80a41-290">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-290">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="80a41-291">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="80a41-291">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="80a41-292">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="80a41-292">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="80a41-293">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="80a41-293">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="80a41-294">利用不可</span><span class="sxs-lookup"><span data-stu-id="80a41-294">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-295">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="80a41-295">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="80a41-296">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="80a41-296">- Mail Read</span></span><br><span data-ttu-id="80a41-297">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="80a41-297">
      - Mail Compose</span></span><br><span data-ttu-id="80a41-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="80a41-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="80a41-299">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="80a41-299">
      - Modules</span></span></td>
    <td> <span data-ttu-id="80a41-300">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-300">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="80a41-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="80a41-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="80a41-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="80a41-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="80a41-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="80a41-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="80a41-304">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="80a41-304">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="80a41-305">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="80a41-305">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="80a41-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="80a41-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="80a41-307">利用不可</span><span class="sxs-lookup"><span data-stu-id="80a41-307">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-308">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="80a41-308">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="80a41-309">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="80a41-309">- Mail Read</span></span><br><span data-ttu-id="80a41-310">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="80a41-310">
      - Mail Compose</span></span><br><span data-ttu-id="80a41-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="80a41-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="80a41-312">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="80a41-312">
      - Modules</span></span></td>
    <td> <span data-ttu-id="80a41-313">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-313">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="80a41-314">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="80a41-314">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="80a41-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="80a41-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="80a41-316">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="80a41-316">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="80a41-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="80a41-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="80a41-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="80a41-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="80a41-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="80a41-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="80a41-320">利用不可</span><span class="sxs-lookup"><span data-stu-id="80a41-320">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-321">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="80a41-321">Office for iOS</span></span></td>
    <td> <span data-ttu-id="80a41-322">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="80a41-322">- Mail Read</span></span><br><span data-ttu-id="80a41-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="80a41-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="80a41-324">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-324">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="80a41-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="80a41-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="80a41-326">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="80a41-326">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="80a41-327">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="80a41-327">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="80a41-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="80a41-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="80a41-329">利用不可</span><span class="sxs-lookup"><span data-stu-id="80a41-329">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-330">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="80a41-330">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="80a41-331">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="80a41-331">- Mail Read</span></span><br><span data-ttu-id="80a41-332">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="80a41-332">
      - Mail Compose</span></span><br><span data-ttu-id="80a41-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="80a41-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="80a41-334">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-334">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="80a41-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="80a41-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="80a41-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="80a41-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="80a41-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="80a41-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="80a41-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="80a41-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="80a41-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="80a41-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="80a41-340">利用不可</span><span class="sxs-lookup"><span data-stu-id="80a41-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-341">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="80a41-341">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="80a41-342">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="80a41-342">- Mail Read</span></span><br><span data-ttu-id="80a41-343">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="80a41-343">
      - Mail Compose</span></span><br><span data-ttu-id="80a41-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="80a41-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="80a41-345">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-345">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="80a41-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="80a41-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="80a41-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="80a41-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="80a41-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="80a41-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="80a41-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="80a41-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="80a41-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="80a41-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="80a41-351">利用不可</span><span class="sxs-lookup"><span data-stu-id="80a41-351">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-352">Office for Android</span><span class="sxs-lookup"><span data-stu-id="80a41-352">Office for Android</span></span></td>
    <td> <span data-ttu-id="80a41-353">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="80a41-353">- Mail Read</span></span><br><span data-ttu-id="80a41-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="80a41-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="80a41-355">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-355">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="80a41-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="80a41-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="80a41-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="80a41-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="80a41-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="80a41-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="80a41-359">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="80a41-359">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="80a41-360">利用不可</span><span class="sxs-lookup"><span data-stu-id="80a41-360">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="80a41-361">Word</span><span class="sxs-lookup"><span data-stu-id="80a41-361">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="80a41-362">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="80a41-362">Platform</span></span></th>
    <th><span data-ttu-id="80a41-363">拡張点</span><span class="sxs-lookup"><span data-stu-id="80a41-363">Extension points</span></span></th>
    <th><span data-ttu-id="80a41-364">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="80a41-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="80a41-365"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="80a41-365"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-366">Office Online</span><span class="sxs-lookup"><span data-stu-id="80a41-366">Office Online</span></span></td>
    <td> <span data-ttu-id="80a41-367">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-367">- TaskPane</span></span><br><span data-ttu-id="80a41-368">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="80a41-368">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="80a41-369">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-369">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="80a41-370">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="80a41-370">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="80a41-371">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="80a41-371">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="80a41-372">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-372">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="80a41-373">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-373">- BindingEvents</span></span><br><span data-ttu-id="80a41-374">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="80a41-374">
         - CustomXmlParts</span></span><br><span data-ttu-id="80a41-375">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-375">
         - DocumentEvents</span></span><br><span data-ttu-id="80a41-376">
         - File</span><span class="sxs-lookup"><span data-stu-id="80a41-376">
         - File</span></span><br><span data-ttu-id="80a41-377">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-377">
         - HtmlCoercion</span></span><br><span data-ttu-id="80a41-378">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-378">
         - ImageCoercion</span></span><br><span data-ttu-id="80a41-379">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-379">
         - MatrixBindings</span></span><br><span data-ttu-id="80a41-380">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-380">
         - MatrixCoercion</span></span><br><span data-ttu-id="80a41-381">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-381">
         - OoxmlCoercion</span></span><br><span data-ttu-id="80a41-382">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="80a41-382">
         - PdfFile</span></span><br><span data-ttu-id="80a41-383">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-383">
         - Selection</span></span><br><span data-ttu-id="80a41-384">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-384">
         - Settings</span></span><br><span data-ttu-id="80a41-385">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-385">
         - TableBindings</span></span><br><span data-ttu-id="80a41-386">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-386">
         - TableCoercion</span></span><br><span data-ttu-id="80a41-387">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-387">
         - TextBindings</span></span><br><span data-ttu-id="80a41-388">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-388">
         - TextCoercion</span></span><br><span data-ttu-id="80a41-389">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="80a41-389">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-390">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="80a41-390">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="80a41-391">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-391">- TaskPane</span></span></td>
    <td> <span data-ttu-id="80a41-392">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="80a41-392">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="80a41-393">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-393">- BindingEvents</span></span><br><span data-ttu-id="80a41-394">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="80a41-394">
         - CompressedFile</span></span><br><span data-ttu-id="80a41-395">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="80a41-395">
         - CustomXmlParts</span></span><br><span data-ttu-id="80a41-396">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-396">
         - DocumentEvents</span></span><br><span data-ttu-id="80a41-397">
         - File</span><span class="sxs-lookup"><span data-stu-id="80a41-397">
         - File</span></span><br><span data-ttu-id="80a41-398">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-398">
         - HtmlCoercion</span></span><br><span data-ttu-id="80a41-399">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-399">
         - ImageCoercion</span></span><br><span data-ttu-id="80a41-400">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-400">
         - MatrixBindings</span></span><br><span data-ttu-id="80a41-401">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-401">
         - MatrixCoercion</span></span><br><span data-ttu-id="80a41-402">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-402">
         - OoxmlCoercion</span></span><br><span data-ttu-id="80a41-403">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="80a41-403">
         - PdfFile</span></span><br><span data-ttu-id="80a41-404">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-404">
         - Selection</span></span><br><span data-ttu-id="80a41-405">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-405">
         - Settings</span></span><br><span data-ttu-id="80a41-406">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-406">
         - TableBindings</span></span><br><span data-ttu-id="80a41-407">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-407">
         - TableCoercion</span></span><br><span data-ttu-id="80a41-408">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-408">
         - TextBindings</span></span><br><span data-ttu-id="80a41-409">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-409">
         - TextCoercion</span></span><br><span data-ttu-id="80a41-410">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="80a41-410">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-411">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="80a41-411">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="80a41-412">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-412">- TaskPane</span></span></td>
    <td> <span data-ttu-id="80a41-413">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-413">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="80a41-414">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="80a41-414">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="80a41-415">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-415">- BindingEvents</span></span><br><span data-ttu-id="80a41-416">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="80a41-416">
         - CompressedFile</span></span><br><span data-ttu-id="80a41-417">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="80a41-417">
         - CustomXmlParts</span></span><br><span data-ttu-id="80a41-418">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-418">
         - DocumentEvents</span></span><br><span data-ttu-id="80a41-419">
         - File</span><span class="sxs-lookup"><span data-stu-id="80a41-419">
         - File</span></span><br><span data-ttu-id="80a41-420">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-420">
         - HtmlCoercion</span></span><br><span data-ttu-id="80a41-421">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-421">
         - ImageCoercion</span></span><br><span data-ttu-id="80a41-422">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-422">
         - MatrixBindings</span></span><br><span data-ttu-id="80a41-423">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-423">
         - MatrixCoercion</span></span><br><span data-ttu-id="80a41-424">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-424">
         - OoxmlCoercion</span></span><br><span data-ttu-id="80a41-425">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="80a41-425">
         - PdfFile</span></span><br><span data-ttu-id="80a41-426">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-426">
         - Selection</span></span><br><span data-ttu-id="80a41-427">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-427">
         - Settings</span></span><br><span data-ttu-id="80a41-428">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-428">
         - TableBindings</span></span><br><span data-ttu-id="80a41-429">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-429">
         - TableCoercion</span></span><br><span data-ttu-id="80a41-430">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-430">
         - TextBindings</span></span><br><span data-ttu-id="80a41-431">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-431">
         - TextCoercion</span></span><br><span data-ttu-id="80a41-432">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="80a41-432">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-433">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="80a41-433">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="80a41-434">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-434">- TaskPane</span></span><br><span data-ttu-id="80a41-435">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="80a41-435">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="80a41-436">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-436">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="80a41-437">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="80a41-437">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="80a41-438">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="80a41-438">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="80a41-439">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-439">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="80a41-440">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-440">- BindingEvents</span></span><br><span data-ttu-id="80a41-441">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="80a41-441">
         - CompressedFile</span></span><br><span data-ttu-id="80a41-442">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="80a41-442">
         - CustomXmlParts</span></span><br><span data-ttu-id="80a41-443">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-443">
         - DocumentEvents</span></span><br><span data-ttu-id="80a41-444">
         - File</span><span class="sxs-lookup"><span data-stu-id="80a41-444">
         - File</span></span><br><span data-ttu-id="80a41-445">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-445">
         - HtmlCoercion</span></span><br><span data-ttu-id="80a41-446">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-446">
         - ImageCoercion</span></span><br><span data-ttu-id="80a41-447">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-447">
         - MatrixBindings</span></span><br><span data-ttu-id="80a41-448">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-448">
         - MatrixCoercion</span></span><br><span data-ttu-id="80a41-449">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-449">
         - OoxmlCoercion</span></span><br><span data-ttu-id="80a41-450">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="80a41-450">
         - PdfFile</span></span><br><span data-ttu-id="80a41-451">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-451">
         - Selection</span></span><br><span data-ttu-id="80a41-452">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-452">
         - Settings</span></span><br><span data-ttu-id="80a41-453">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-453">
         - TableBindings</span></span><br><span data-ttu-id="80a41-454">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-454">
         - TableCoercion</span></span><br><span data-ttu-id="80a41-455">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-455">
         - TextBindings</span></span><br><span data-ttu-id="80a41-456">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-456">
         - TextCoercion</span></span><br><span data-ttu-id="80a41-457">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="80a41-457">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-458">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="80a41-458">Office for iPad</span></span></td>
    <td> <span data-ttu-id="80a41-459">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-459">- TaskPane</span></span></td>
    <td> <span data-ttu-id="80a41-460">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-460">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="80a41-461">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="80a41-461">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="80a41-462">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="80a41-462">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="80a41-463">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="80a41-463">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="80a41-464">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-464">- BindingEvents</span></span><br><span data-ttu-id="80a41-465">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="80a41-465">
         - CompressedFile</span></span><br><span data-ttu-id="80a41-466">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="80a41-466">
         - CustomXmlParts</span></span><br><span data-ttu-id="80a41-467">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-467">
         - DocumentEvents</span></span><br><span data-ttu-id="80a41-468">
         - File</span><span class="sxs-lookup"><span data-stu-id="80a41-468">
         - File</span></span><br><span data-ttu-id="80a41-469">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-469">
         - HtmlCoercion</span></span><br><span data-ttu-id="80a41-470">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-470">
         - ImageCoercion</span></span><br><span data-ttu-id="80a41-471">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-471">
         - MatrixBindings</span></span><br><span data-ttu-id="80a41-472">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-472">
         - MatrixCoercion</span></span><br><span data-ttu-id="80a41-473">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-473">
         - OoxmlCoercion</span></span><br><span data-ttu-id="80a41-474">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="80a41-474">
         - PdfFile</span></span><br><span data-ttu-id="80a41-475">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-475">
         - Selection</span></span><br><span data-ttu-id="80a41-476">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-476">
         - Settings</span></span><br><span data-ttu-id="80a41-477">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-477">
         - TableBindings</span></span><br><span data-ttu-id="80a41-478">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-478">
         - TableCoercion</span></span><br><span data-ttu-id="80a41-479">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-479">
         - TextBindings</span></span><br><span data-ttu-id="80a41-480">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-480">
         - TextCoercion</span></span><br><span data-ttu-id="80a41-481">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="80a41-481">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-482">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="80a41-482">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="80a41-483">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-483">- TaskPane</span></span></td>
    <td> <span data-ttu-id="80a41-484">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-484">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="80a41-485">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="80a41-485">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="80a41-486">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-486">- BindingEvents</span></span><br><span data-ttu-id="80a41-487">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="80a41-487">
         - CompressedFile</span></span><br><span data-ttu-id="80a41-488">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="80a41-488">
         - CustomXmlParts</span></span><br><span data-ttu-id="80a41-489">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-489">
         - DocumentEvents</span></span><br><span data-ttu-id="80a41-490">
         - File</span><span class="sxs-lookup"><span data-stu-id="80a41-490">
         - File</span></span><br><span data-ttu-id="80a41-491">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-491">
         - HtmlCoercion</span></span><br><span data-ttu-id="80a41-492">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-492">
         - ImageCoercion</span></span><br><span data-ttu-id="80a41-493">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-493">
         - MatrixBindings</span></span><br><span data-ttu-id="80a41-494">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-494">
         - MatrixCoercion</span></span><br><span data-ttu-id="80a41-495">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-495">
         - OoxmlCoercion</span></span><br><span data-ttu-id="80a41-496">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="80a41-496">
         - PdfFile</span></span><br><span data-ttu-id="80a41-497">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-497">
         - Selection</span></span><br><span data-ttu-id="80a41-498">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-498">
         - Settings</span></span><br><span data-ttu-id="80a41-499">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-499">
         - TableBindings</span></span><br><span data-ttu-id="80a41-500">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-500">
         - TableCoercion</span></span><br><span data-ttu-id="80a41-501">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-501">
         - TextBindings</span></span><br><span data-ttu-id="80a41-502">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-502">
         - TextCoercion</span></span><br><span data-ttu-id="80a41-503">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="80a41-503">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-504">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="80a41-504">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="80a41-505">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-505">- TaskPane</span></span><br><span data-ttu-id="80a41-506">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="80a41-506">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="80a41-507">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-507">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="80a41-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="80a41-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="80a41-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="80a41-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="80a41-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="80a41-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="80a41-511">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-511">- BindingEvents</span></span><br><span data-ttu-id="80a41-512">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="80a41-512">
         - CompressedFile</span></span><br><span data-ttu-id="80a41-513">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="80a41-513">
         - CustomXmlParts</span></span><br><span data-ttu-id="80a41-514">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-514">
         - DocumentEvents</span></span><br><span data-ttu-id="80a41-515">
         - File</span><span class="sxs-lookup"><span data-stu-id="80a41-515">
         - File</span></span><br><span data-ttu-id="80a41-516">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-516">
         - HtmlCoercion</span></span><br><span data-ttu-id="80a41-517">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-517">
         - ImageCoercion</span></span><br><span data-ttu-id="80a41-518">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-518">
         - MatrixBindings</span></span><br><span data-ttu-id="80a41-519">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-519">
         - MatrixCoercion</span></span><br><span data-ttu-id="80a41-520">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-520">
         - OoxmlCoercion</span></span><br><span data-ttu-id="80a41-521">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="80a41-521">
         - PdfFile</span></span><br><span data-ttu-id="80a41-522">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-522">
         - Selection</span></span><br><span data-ttu-id="80a41-523">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-523">
         - Settings</span></span><br><span data-ttu-id="80a41-524">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-524">
         - TableBindings</span></span><br><span data-ttu-id="80a41-525">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-525">
         - TableCoercion</span></span><br><span data-ttu-id="80a41-526">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="80a41-526">
         - TextBindings</span></span><br><span data-ttu-id="80a41-527">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-527">
         - TextCoercion</span></span><br><span data-ttu-id="80a41-528">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="80a41-528">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="80a41-529">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="80a41-529">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="80a41-530">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="80a41-530">Platform</span></span></th>
    <th><span data-ttu-id="80a41-531">拡張点</span><span class="sxs-lookup"><span data-stu-id="80a41-531">Extension points</span></span></th>
    <th><span data-ttu-id="80a41-532">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="80a41-532">API requirement sets</span></span></th>
    <th><span data-ttu-id="80a41-533"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="80a41-533"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-534">Office Online</span><span class="sxs-lookup"><span data-stu-id="80a41-534">Office Online</span></span></td>
    <td> <span data-ttu-id="80a41-535">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="80a41-535">- Content</span></span><br><span data-ttu-id="80a41-536">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-536">
         - TaskPane</span></span><br><span data-ttu-id="80a41-537">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="80a41-537">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="80a41-538">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-538">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="80a41-539">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="80a41-539">- ActiveView</span></span><br><span data-ttu-id="80a41-540">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="80a41-540">
         - CompressedFile</span></span><br><span data-ttu-id="80a41-541">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-541">
         - DocumentEvents</span></span><br><span data-ttu-id="80a41-542">
         - File</span><span class="sxs-lookup"><span data-stu-id="80a41-542">
         - File</span></span><br><span data-ttu-id="80a41-543">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-543">
         - ImageCoercion</span></span><br><span data-ttu-id="80a41-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="80a41-544">
         - PdfFile</span></span><br><span data-ttu-id="80a41-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-545">
         - Selection</span></span><br><span data-ttu-id="80a41-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-546">
         - Settings</span></span><br><span data-ttu-id="80a41-547">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-547">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-548">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="80a41-548">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="80a41-549">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="80a41-549">- Content</span></span><br><span data-ttu-id="80a41-550">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-550">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="80a41-551">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="80a41-551">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="80a41-552">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="80a41-552">- ActiveView</span></span><br><span data-ttu-id="80a41-553">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="80a41-553">
         - CompressedFile</span></span><br><span data-ttu-id="80a41-554">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-554">
         - DocumentEvents</span></span><br><span data-ttu-id="80a41-555">
         - File</span><span class="sxs-lookup"><span data-stu-id="80a41-555">
         - File</span></span><br><span data-ttu-id="80a41-556">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-556">
         - ImageCoercion</span></span><br><span data-ttu-id="80a41-557">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="80a41-557">
         - PdfFile</span></span><br><span data-ttu-id="80a41-558">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-558">
         - Selection</span></span><br><span data-ttu-id="80a41-559">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-559">
         - Settings</span></span><br><span data-ttu-id="80a41-560">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-560">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-561">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="80a41-561">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="80a41-562">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="80a41-562">- Content</span></span><br><span data-ttu-id="80a41-563">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-563">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="80a41-564">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="80a41-564">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="80a41-565">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="80a41-565">- ActiveView</span></span><br><span data-ttu-id="80a41-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="80a41-566">
         - CompressedFile</span></span><br><span data-ttu-id="80a41-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-567">
         - DocumentEvents</span></span><br><span data-ttu-id="80a41-568">
         - File</span><span class="sxs-lookup"><span data-stu-id="80a41-568">
         - File</span></span><br><span data-ttu-id="80a41-569">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-569">
         - ImageCoercion</span></span><br><span data-ttu-id="80a41-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="80a41-570">
         - PdfFile</span></span><br><span data-ttu-id="80a41-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-571">
         - Selection</span></span><br><span data-ttu-id="80a41-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-572">
         - Settings</span></span><br><span data-ttu-id="80a41-573">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-573">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-574">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="80a41-574">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="80a41-575">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="80a41-575">- Content</span></span><br><span data-ttu-id="80a41-576">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-576">
         - TaskPane</span></span><br><span data-ttu-id="80a41-577">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="80a41-577">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="80a41-578">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-578">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="80a41-579">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="80a41-579">- ActiveView</span></span><br><span data-ttu-id="80a41-580">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="80a41-580">
         - CompressedFile</span></span><br><span data-ttu-id="80a41-581">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-581">
         - DocumentEvents</span></span><br><span data-ttu-id="80a41-582">
         - File</span><span class="sxs-lookup"><span data-stu-id="80a41-582">
         - File</span></span><br><span data-ttu-id="80a41-583">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-583">
         - ImageCoercion</span></span><br><span data-ttu-id="80a41-584">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="80a41-584">
         - PdfFile</span></span><br><span data-ttu-id="80a41-585">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-585">
         - Selection</span></span><br><span data-ttu-id="80a41-586">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-586">
         - Settings</span></span><br><span data-ttu-id="80a41-587">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-587">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-588">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="80a41-588">Office for iPad</span></span></td>
    <td> <span data-ttu-id="80a41-589">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="80a41-589">- Content</span></span><br><span data-ttu-id="80a41-590">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-590">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="80a41-591">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-591">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="80a41-592">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="80a41-592">- ActiveView</span></span><br><span data-ttu-id="80a41-593">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="80a41-593">
         - CompressedFile</span></span><br><span data-ttu-id="80a41-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-594">
         - DocumentEvents</span></span><br><span data-ttu-id="80a41-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="80a41-595">
         - File</span></span><br><span data-ttu-id="80a41-596">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="80a41-596">
         - PdfFile</span></span><br><span data-ttu-id="80a41-597">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-597">
         - Selection</span></span><br><span data-ttu-id="80a41-598">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-598">
         - Settings</span></span><br><span data-ttu-id="80a41-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-599">
         - TextCoercion</span></span><br><span data-ttu-id="80a41-600">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-600">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-601">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="80a41-601">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="80a41-602">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="80a41-602">- Content</span></span><br><span data-ttu-id="80a41-603">
         - 作業ウィンドウ/td></span><span class="sxs-lookup"><span data-stu-id="80a41-603">
         - TaskPane/td></span></span> <td> <span data-ttu-id="80a41-604">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="80a41-604">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="80a41-605">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="80a41-605">- ActiveView</span></span><br><span data-ttu-id="80a41-606">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="80a41-606">
         - CompressedFile</span></span><br><span data-ttu-id="80a41-607">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-607">
         - DocumentEvents</span></span><br><span data-ttu-id="80a41-608">
         - File</span><span class="sxs-lookup"><span data-stu-id="80a41-608">
         - File</span></span><br><span data-ttu-id="80a41-609">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-609">
         - ImageCoercion</span></span><br><span data-ttu-id="80a41-610">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="80a41-610">
         - PdfFile</span></span><br><span data-ttu-id="80a41-611">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-611">
         - Selection</span></span><br><span data-ttu-id="80a41-612">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-612">
         - Settings</span></span><br><span data-ttu-id="80a41-613">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-613">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-614">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="80a41-614">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="80a41-615">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="80a41-615">- Content</span></span><br><span data-ttu-id="80a41-616">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-616">
         - TaskPane</span></span><br><span data-ttu-id="80a41-617">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="80a41-617">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="80a41-618">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-618">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="80a41-619">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="80a41-619">- ActiveView</span></span><br><span data-ttu-id="80a41-620">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="80a41-620">
         - CompressedFile</span></span><br><span data-ttu-id="80a41-621">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-621">
         - DocumentEvents</span></span><br><span data-ttu-id="80a41-622">
         - File</span><span class="sxs-lookup"><span data-stu-id="80a41-622">
         - File</span></span><br><span data-ttu-id="80a41-623">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-623">
         - ImageCoercion</span></span><br><span data-ttu-id="80a41-624">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="80a41-624">
         - PdfFile</span></span><br><span data-ttu-id="80a41-625">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-625">
         - Selection</span></span><br><span data-ttu-id="80a41-626">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-626">
         - Settings</span></span><br><span data-ttu-id="80a41-627">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-627">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="80a41-628">OneNote</span><span class="sxs-lookup"><span data-stu-id="80a41-628">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="80a41-629">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="80a41-629">Platform</span></span></th>
    <th><span data-ttu-id="80a41-630">拡張点</span><span class="sxs-lookup"><span data-stu-id="80a41-630">Extension points</span></span></th>
    <th><span data-ttu-id="80a41-631">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="80a41-631">API requirement sets</span></span></th>
    <th><span data-ttu-id="80a41-632"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="80a41-632"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-633">Office Online</span><span class="sxs-lookup"><span data-stu-id="80a41-633">Office Online</span></span></td>
    <td> <span data-ttu-id="80a41-634">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="80a41-634">- Content</span></span><br><span data-ttu-id="80a41-635">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-635">
         - TaskPane</span></span><br><span data-ttu-id="80a41-636">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="80a41-636">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="80a41-637">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-637">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="80a41-638">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-638">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="80a41-639">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="80a41-639">- DocumentEvents</span></span><br><span data-ttu-id="80a41-640">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-640">
         - HtmlCoercion</span></span><br><span data-ttu-id="80a41-641">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-641">
         - ImageCoercion</span></span><br><span data-ttu-id="80a41-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="80a41-642">
         - Settings</span></span><br><span data-ttu-id="80a41-643">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-643">
         - TextCoercion</span></span></td>
  </tr>
</table><span data-ttu-id="80a41-644">
\*&ast; - リリース後の更新プログラムで追加されました。*

</span><span class="sxs-lookup"><span data-stu-id="80a41-644">
\*&ast; - Added with post-release updates.*

</span></span><br/>

## <a name="project"></a><span data-ttu-id="80a41-645">Project</span><span class="sxs-lookup"><span data-stu-id="80a41-645">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="80a41-646">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="80a41-646">Platform</span></span></th>
    <th><span data-ttu-id="80a41-647">拡張点</span><span class="sxs-lookup"><span data-stu-id="80a41-647">Extension points</span></span></th>
    <th><span data-ttu-id="80a41-648">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="80a41-648">API requirement sets</span></span></th>
    <th><span data-ttu-id="80a41-649"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="80a41-649"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-650">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="80a41-650">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="80a41-651">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-651">- TaskPane</span></span></td>
    <td> <span data-ttu-id="80a41-652">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-652">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="80a41-653">- Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-653">- Selection</span></span><br><span data-ttu-id="80a41-654">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-654">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-655">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="80a41-655">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="80a41-656">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-656">- TaskPane</span></span></td>
    <td> <span data-ttu-id="80a41-657">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-657">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="80a41-658">- Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-658">- Selection</span></span><br><span data-ttu-id="80a41-659">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-659">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="80a41-660">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="80a41-660">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="80a41-661">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="80a41-661">- TaskPane</span></span></td>
    <td> <span data-ttu-id="80a41-662">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="80a41-662">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="80a41-663">- Selection</span><span class="sxs-lookup"><span data-stu-id="80a41-663">- Selection</span></span><br><span data-ttu-id="80a41-664">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="80a41-664">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="80a41-665">関連項目</span><span class="sxs-lookup"><span data-stu-id="80a41-665">See also</span></span>

- [<span data-ttu-id="80a41-666">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="80a41-666">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="80a41-667">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="80a41-667">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="80a41-668">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="80a41-668">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="80a41-669">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="80a41-669">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
