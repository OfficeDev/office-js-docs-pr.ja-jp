---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、Word、Outlook、PowerPoint、OneNote、および Project のサポートされる要件セット。
ms.date: 11/07/2018
localization_priority: Priority
ms.openlocfilehash: 9f8b94483d22f24dcb0a6a2ad99df6167533133f
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388340"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="670d9-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="670d9-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="670d9-104">Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。</span><span class="sxs-lookup"><span data-stu-id="670d9-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="670d9-105">次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、各 Office アプリケーションで現在サポートされている共通 API が記載されています。</span><span class="sxs-lookup"><span data-stu-id="670d9-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="670d9-p102">MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。このバージョンには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="670d9-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="670d9-108">Excel</span><span class="sxs-lookup"><span data-stu-id="670d9-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="670d9-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="670d9-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="670d9-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="670d9-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="670d9-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="670d9-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="670d9-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="670d9-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="670d9-113">Office Online</span></span></td>
    <td> <span data-ttu-id="670d9-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-114">- TaskPane</span></span><br><span data-ttu-id="670d9-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="670d9-115">
        - Content</span></span><br><span data-ttu-id="670d9-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="670d9-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="670d9-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="670d9-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="670d9-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="670d9-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="670d9-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="670d9-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="670d9-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="670d9-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="670d9-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="670d9-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="670d9-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="670d9-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="670d9-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="670d9-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="670d9-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="670d9-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="670d9-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-126">
        - BindingEvents</span></span><br><span data-ttu-id="670d9-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="670d9-127">
        - CompressedFile</span></span><br><span data-ttu-id="670d9-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-128">
        - DocumentEvents</span></span><br><span data-ttu-id="670d9-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="670d9-129">
        - File</span></span><br><span data-ttu-id="670d9-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-130">
        - MatrixBindings</span></span><br><span data-ttu-id="670d9-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="670d9-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-132">
        - Selection</span></span><br><span data-ttu-id="670d9-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-133">
        - Settings</span></span><br><span data-ttu-id="670d9-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-134">
        - TableBindings</span></span><br><span data-ttu-id="670d9-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-135">
        - TableCoercion</span></span><br><span data-ttu-id="670d9-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-136">
        - TextBindings</span></span><br><span data-ttu-id="670d9-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-138">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="670d9-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="670d9-139">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-139">
        - TaskPane</span></span><br><span data-ttu-id="670d9-140">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="670d9-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="670d9-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="670d9-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-142">
        - BindingEvents</span></span><br><span data-ttu-id="670d9-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="670d9-143">
        - CompressedFile</span></span><br><span data-ttu-id="670d9-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-144">
        - DocumentEvents</span></span><br><span data-ttu-id="670d9-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="670d9-145">
        - File</span></span><br><span data-ttu-id="670d9-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-146">
        - ImageCoercion</span></span><br><span data-ttu-id="670d9-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-147">
        - MatrixBindings</span></span><br><span data-ttu-id="670d9-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="670d9-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-149">
        - Selection</span></span><br><span data-ttu-id="670d9-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-150">
        - Settings</span></span><br><span data-ttu-id="670d9-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-151">
        - TableBindings</span></span><br><span data-ttu-id="670d9-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-152">
        - TableCoercion</span></span><br><span data-ttu-id="670d9-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-153">
        - TextBindings</span></span><br><span data-ttu-id="670d9-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-155">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="670d9-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="670d9-156">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-156">- TaskPane</span></span><br><span data-ttu-id="670d9-157">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="670d9-157">
        - Content</span></span><br><span data-ttu-id="670d9-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="670d9-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="670d9-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="670d9-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="670d9-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="670d9-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="670d9-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="670d9-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="670d9-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="670d9-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="670d9-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="670d9-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="670d9-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="670d9-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="670d9-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="670d9-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="670d9-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="670d9-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-168">- BindingEvents</span></span><br><span data-ttu-id="670d9-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="670d9-169">
        - CompressedFile</span></span><br><span data-ttu-id="670d9-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-170">
        - DocumentEvents</span></span><br><span data-ttu-id="670d9-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="670d9-171">
        - File</span></span><br><span data-ttu-id="670d9-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-172">
        - ImageCoercion</span></span><br><span data-ttu-id="670d9-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-173">
        - MatrixBindings</span></span><br><span data-ttu-id="670d9-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-174">
        - MatrixCoercion</span></span><br><span data-ttu-id="670d9-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-175">
        - Selection</span></span><br><span data-ttu-id="670d9-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-176">
        - Settings</span></span><br><span data-ttu-id="670d9-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-177">
        - TableBindings</span></span><br><span data-ttu-id="670d9-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-178">
        - TableCoercion</span></span><br><span data-ttu-id="670d9-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-179">
        - TextBindings</span></span><br><span data-ttu-id="670d9-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-180">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-181">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="670d9-181">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="670d9-182">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-182">- TaskPane</span></span><br><span data-ttu-id="670d9-183">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="670d9-183">
        - Content</span></span><br><span data-ttu-id="670d9-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="670d9-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="670d9-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="670d9-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="670d9-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="670d9-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="670d9-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="670d9-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="670d9-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="670d9-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="670d9-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="670d9-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="670d9-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="670d9-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="670d9-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="670d9-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="670d9-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="670d9-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-194">- BindingEvents</span></span><br><span data-ttu-id="670d9-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="670d9-195">
        - CompressedFile</span></span><br><span data-ttu-id="670d9-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-196">
        - DocumentEvents</span></span><br><span data-ttu-id="670d9-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="670d9-197">
        - File</span></span><br><span data-ttu-id="670d9-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-198">
        - ImageCoercion</span></span><br><span data-ttu-id="670d9-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-199">
        - MatrixBindings</span></span><br><span data-ttu-id="670d9-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="670d9-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-201">
        - Selection</span></span><br><span data-ttu-id="670d9-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-202">
        - Settings</span></span><br><span data-ttu-id="670d9-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-203">
        - TableBindings</span></span><br><span data-ttu-id="670d9-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-204">
        - TableCoercion</span></span><br><span data-ttu-id="670d9-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-205">
        - TextBindings</span></span><br><span data-ttu-id="670d9-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-207">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="670d9-207">Office for iPad</span></span></td>
    <td><span data-ttu-id="670d9-208">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-208">- TaskPane</span></span><br><span data-ttu-id="670d9-209">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="670d9-209">
        - Content</span></span></td>
    <td><span data-ttu-id="670d9-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="670d9-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="670d9-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="670d9-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="670d9-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="670d9-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="670d9-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="670d9-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="670d9-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="670d9-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="670d9-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="670d9-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="670d9-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="670d9-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="670d9-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="670d9-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="670d9-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-219">- BindingEvents</span></span><br><span data-ttu-id="670d9-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="670d9-220">
        - CompressedFile</span></span><br><span data-ttu-id="670d9-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-221">
        - DocumentEvents</span></span><br><span data-ttu-id="670d9-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="670d9-222">
        - File</span></span><br><span data-ttu-id="670d9-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-223">
        - ImageCoercion</span></span><br><span data-ttu-id="670d9-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-224">
        - MatrixBindings</span></span><br><span data-ttu-id="670d9-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="670d9-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-226">
        - Selection</span></span><br><span data-ttu-id="670d9-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-227">
        - Settings</span></span><br><span data-ttu-id="670d9-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-228">
        - TableBindings</span></span><br><span data-ttu-id="670d9-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-229">
        - TableCoercion</span></span><br><span data-ttu-id="670d9-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-230">
        - TextBindings</span></span><br><span data-ttu-id="670d9-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-232">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="670d9-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="670d9-233">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-233">- TaskPane</span></span><br><span data-ttu-id="670d9-234">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="670d9-234">
        - Content</span></span><br><span data-ttu-id="670d9-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="670d9-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="670d9-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="670d9-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="670d9-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="670d9-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="670d9-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="670d9-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="670d9-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="670d9-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="670d9-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="670d9-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="670d9-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="670d9-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="670d9-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="670d9-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="670d9-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="670d9-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-245">- BindingEvents</span></span><br><span data-ttu-id="670d9-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="670d9-246">
        - CompressedFile</span></span><br><span data-ttu-id="670d9-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-247">
        - DocumentEvents</span></span><br><span data-ttu-id="670d9-248">
        - File</span><span class="sxs-lookup"><span data-stu-id="670d9-248">
        - File</span></span><br><span data-ttu-id="670d9-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-249">
        - ImageCoercion</span></span><br><span data-ttu-id="670d9-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-250">
        - MatrixBindings</span></span><br><span data-ttu-id="670d9-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-251">
        - MatrixCoercion</span></span><br><span data-ttu-id="670d9-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="670d9-252">
        - PdfFile</span></span><br><span data-ttu-id="670d9-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-253">
        - Selection</span></span><br><span data-ttu-id="670d9-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-254">
        - Settings</span></span><br><span data-ttu-id="670d9-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-255">
        - TableBindings</span></span><br><span data-ttu-id="670d9-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-256">
        - TableCoercion</span></span><br><span data-ttu-id="670d9-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-257">
        - TextBindings</span></span><br><span data-ttu-id="670d9-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-259">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="670d9-259">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="670d9-260">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-260">- TaskPane</span></span><br><span data-ttu-id="670d9-261">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="670d9-261">
        - Content</span></span><br><span data-ttu-id="670d9-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="670d9-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="670d9-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="670d9-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="670d9-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="670d9-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="670d9-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="670d9-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="670d9-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="670d9-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="670d9-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="670d9-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="670d9-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="670d9-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="670d9-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="670d9-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="670d9-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="670d9-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-272">- BindingEvents</span></span><br><span data-ttu-id="670d9-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="670d9-273">
        - CompressedFile</span></span><br><span data-ttu-id="670d9-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-274">
        - DocumentEvents</span></span><br><span data-ttu-id="670d9-275">
        - File</span><span class="sxs-lookup"><span data-stu-id="670d9-275">
        - File</span></span><br><span data-ttu-id="670d9-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-276">
        - ImageCoercion</span></span><br><span data-ttu-id="670d9-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-277">
        - MatrixBindings</span></span><br><span data-ttu-id="670d9-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-278">
        - MatrixCoercion</span></span><br><span data-ttu-id="670d9-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="670d9-279">
        - PdfFile</span></span><br><span data-ttu-id="670d9-280">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-280">
        - Selection</span></span><br><span data-ttu-id="670d9-281">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-281">
        - Settings</span></span><br><span data-ttu-id="670d9-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-282">
        - TableBindings</span></span><br><span data-ttu-id="670d9-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-283">
        - TableCoercion</span></span><br><span data-ttu-id="670d9-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-284">
        - TextBindings</span></span><br><span data-ttu-id="670d9-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-285">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="670d9-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="670d9-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="670d9-287">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="670d9-287">Platform</span></span></th>
    <th><span data-ttu-id="670d9-288">拡張点</span><span class="sxs-lookup"><span data-stu-id="670d9-288">Extension points</span></span></th>
    <th><span data-ttu-id="670d9-289">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="670d9-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="670d9-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="670d9-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="670d9-291">Office Online</span></span></td>
    <td> <span data-ttu-id="670d9-292">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="670d9-292">- Mail Read</span></span><br><span data-ttu-id="670d9-293">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="670d9-293">
      - Mail Compose</span></span><br><span data-ttu-id="670d9-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="670d9-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="670d9-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="670d9-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="670d9-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="670d9-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="670d9-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="670d9-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="670d9-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="670d9-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="670d9-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="670d9-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="670d9-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="670d9-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="670d9-302">利用不可</span><span class="sxs-lookup"><span data-stu-id="670d9-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-303">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="670d9-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="670d9-304">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="670d9-304">- Mail Read</span></span><br><span data-ttu-id="670d9-305">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="670d9-305">
      - Mail Compose</span></span><br><span data-ttu-id="670d9-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="670d9-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="670d9-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="670d9-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="670d9-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="670d9-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="670d9-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="670d9-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="670d9-311">利用不可</span><span class="sxs-lookup"><span data-stu-id="670d9-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-312">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="670d9-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="670d9-313">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="670d9-313">- Mail Read</span></span><br><span data-ttu-id="670d9-314">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="670d9-314">
      - Mail Compose</span></span><br><span data-ttu-id="670d9-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="670d9-316">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="670d9-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="670d9-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="670d9-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="670d9-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="670d9-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="670d9-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="670d9-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="670d9-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="670d9-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="670d9-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="670d9-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="670d9-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="670d9-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="670d9-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="670d9-324">利用不可</span><span class="sxs-lookup"><span data-stu-id="670d9-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-325">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="670d9-325">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="670d9-326">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="670d9-326">- Mail Read</span></span><br><span data-ttu-id="670d9-327">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="670d9-327">
      - Mail Compose</span></span><br><span data-ttu-id="670d9-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="670d9-329">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="670d9-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="670d9-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="670d9-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="670d9-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="670d9-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="670d9-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="670d9-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="670d9-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="670d9-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="670d9-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="670d9-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="670d9-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="670d9-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="670d9-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="670d9-337">利用不可</span><span class="sxs-lookup"><span data-stu-id="670d9-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-338">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="670d9-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="670d9-339">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="670d9-339">- Mail Read</span></span><br><span data-ttu-id="670d9-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="670d9-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="670d9-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="670d9-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="670d9-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="670d9-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="670d9-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="670d9-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="670d9-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="670d9-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="670d9-346">利用不可</span><span class="sxs-lookup"><span data-stu-id="670d9-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-347">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="670d9-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="670d9-348">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="670d9-348">- Mail Read</span></span><br><span data-ttu-id="670d9-349">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="670d9-349">
      - Mail Compose</span></span><br><span data-ttu-id="670d9-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="670d9-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="670d9-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="670d9-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="670d9-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="670d9-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="670d9-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="670d9-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="670d9-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="670d9-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="670d9-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="670d9-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="670d9-357">利用不可</span><span class="sxs-lookup"><span data-stu-id="670d9-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-358">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="670d9-358">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="670d9-359">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="670d9-359">- Mail Read</span></span><br><span data-ttu-id="670d9-360">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="670d9-360">
      - Mail Compose</span></span><br><span data-ttu-id="670d9-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="670d9-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="670d9-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="670d9-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="670d9-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="670d9-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="670d9-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="670d9-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="670d9-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="670d9-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="670d9-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="670d9-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="670d9-368">利用不可</span><span class="sxs-lookup"><span data-stu-id="670d9-368">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-369">Office for Android</span><span class="sxs-lookup"><span data-stu-id="670d9-369">Office for Android</span></span></td>
    <td> <span data-ttu-id="670d9-370">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="670d9-370">- Mail Read</span></span><br><span data-ttu-id="670d9-371">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-371">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="670d9-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="670d9-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="670d9-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="670d9-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="670d9-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="670d9-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="670d9-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="670d9-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="670d9-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="670d9-377">利用不可</span><span class="sxs-lookup"><span data-stu-id="670d9-377">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="670d9-378">Word</span><span class="sxs-lookup"><span data-stu-id="670d9-378">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="670d9-379">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="670d9-379">Platform</span></span></th>
    <th><span data-ttu-id="670d9-380">拡張点</span><span class="sxs-lookup"><span data-stu-id="670d9-380">Extension points</span></span></th>
    <th><span data-ttu-id="670d9-381">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="670d9-381">API requirement sets</span></span></th>
    <th><span data-ttu-id="670d9-382"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="670d9-382"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-383">Office Online</span><span class="sxs-lookup"><span data-stu-id="670d9-383">Office Online</span></span></td>
    <td> <span data-ttu-id="670d9-384">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-384">- TaskPane</span></span><br><span data-ttu-id="670d9-385">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-385">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="670d9-386">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-386">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="670d9-387">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="670d9-387">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="670d9-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="670d9-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="670d9-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="670d9-390">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-390">- BindingEvents</span></span><br><span data-ttu-id="670d9-391">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="670d9-391">
         - CustomXmlParts</span></span><br><span data-ttu-id="670d9-392">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-392">
         - DocumentEvents</span></span><br><span data-ttu-id="670d9-393">
         - File</span><span class="sxs-lookup"><span data-stu-id="670d9-393">
         - File</span></span><br><span data-ttu-id="670d9-394">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-394">
         - HtmlCoercion</span></span><br><span data-ttu-id="670d9-395">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-395">
         - ImageCoercion</span></span><br><span data-ttu-id="670d9-396">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-396">
         - MatrixBindings</span></span><br><span data-ttu-id="670d9-397">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-397">
         - MatrixCoercion</span></span><br><span data-ttu-id="670d9-398">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-398">
         - OoxmlCoercion</span></span><br><span data-ttu-id="670d9-399">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="670d9-399">
         - PdfFile</span></span><br><span data-ttu-id="670d9-400">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-400">
         - Selection</span></span><br><span data-ttu-id="670d9-401">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-401">
         - Settings</span></span><br><span data-ttu-id="670d9-402">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-402">
         - TableBindings</span></span><br><span data-ttu-id="670d9-403">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-403">
         - TableCoercion</span></span><br><span data-ttu-id="670d9-404">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-404">
         - TextBindings</span></span><br><span data-ttu-id="670d9-405">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-405">
         - TextCoercion</span></span><br><span data-ttu-id="670d9-406">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="670d9-406">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-407">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="670d9-407">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="670d9-408">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-408">- TaskPane</span></span></td>
    <td> <span data-ttu-id="670d9-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="670d9-410">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-410">- BindingEvents</span></span><br><span data-ttu-id="670d9-411">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="670d9-411">
         - CompressedFile</span></span><br><span data-ttu-id="670d9-412">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="670d9-412">
         - CustomXmlParts</span></span><br><span data-ttu-id="670d9-413">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-413">
         - DocumentEvents</span></span><br><span data-ttu-id="670d9-414">
         - File</span><span class="sxs-lookup"><span data-stu-id="670d9-414">
         - File</span></span><br><span data-ttu-id="670d9-415">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-415">
         - HtmlCoercion</span></span><br><span data-ttu-id="670d9-416">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-416">
         - ImageCoercion</span></span><br><span data-ttu-id="670d9-417">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-417">
         - MatrixBindings</span></span><br><span data-ttu-id="670d9-418">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-418">
         - MatrixCoercion</span></span><br><span data-ttu-id="670d9-419">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-419">
         - OoxmlCoercion</span></span><br><span data-ttu-id="670d9-420">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="670d9-420">
         - PdfFile</span></span><br><span data-ttu-id="670d9-421">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-421">
         - Selection</span></span><br><span data-ttu-id="670d9-422">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-422">
         - Settings</span></span><br><span data-ttu-id="670d9-423">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-423">
         - TableBindings</span></span><br><span data-ttu-id="670d9-424">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-424">
         - TableCoercion</span></span><br><span data-ttu-id="670d9-425">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-425">
         - TextBindings</span></span><br><span data-ttu-id="670d9-426">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-426">
         - TextCoercion</span></span><br><span data-ttu-id="670d9-427">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="670d9-427">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-428">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="670d9-428">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="670d9-429">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-429">- TaskPane</span></span><br><span data-ttu-id="670d9-430">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-430">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="670d9-431">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-431">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="670d9-432">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="670d9-432">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="670d9-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="670d9-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="670d9-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="670d9-435">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-435">- BindingEvents</span></span><br><span data-ttu-id="670d9-436">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="670d9-436">
         - CompressedFile</span></span><br><span data-ttu-id="670d9-437">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="670d9-437">
         - CustomXmlParts</span></span><br><span data-ttu-id="670d9-438">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-438">
         - DocumentEvents</span></span><br><span data-ttu-id="670d9-439">
         - File</span><span class="sxs-lookup"><span data-stu-id="670d9-439">
         - File</span></span><br><span data-ttu-id="670d9-440">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-440">
         - HtmlCoercion</span></span><br><span data-ttu-id="670d9-441">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-441">
         - ImageCoercion</span></span><br><span data-ttu-id="670d9-442">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-442">
         - MatrixBindings</span></span><br><span data-ttu-id="670d9-443">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-443">
         - MatrixCoercion</span></span><br><span data-ttu-id="670d9-444">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-444">
         - OoxmlCoercion</span></span><br><span data-ttu-id="670d9-445">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="670d9-445">
         - PdfFile</span></span><br><span data-ttu-id="670d9-446">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-446">
         - Selection</span></span><br><span data-ttu-id="670d9-447">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-447">
         - Settings</span></span><br><span data-ttu-id="670d9-448">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-448">
         - TableBindings</span></span><br><span data-ttu-id="670d9-449">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-449">
         - TableCoercion</span></span><br><span data-ttu-id="670d9-450">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-450">
         - TextBindings</span></span><br><span data-ttu-id="670d9-451">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-451">
         - TextCoercion</span></span><br><span data-ttu-id="670d9-452">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="670d9-452">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-453">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="670d9-453">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="670d9-454">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-454">- TaskPane</span></span><br><span data-ttu-id="670d9-455">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-455">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="670d9-456">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-456">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="670d9-457">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="670d9-457">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="670d9-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="670d9-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="670d9-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="670d9-460">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-460">- BindingEvents</span></span><br><span data-ttu-id="670d9-461">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="670d9-461">
         - CompressedFile</span></span><br><span data-ttu-id="670d9-462">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="670d9-462">
         - CustomXmlParts</span></span><br><span data-ttu-id="670d9-463">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-463">
         - DocumentEvents</span></span><br><span data-ttu-id="670d9-464">
         - File</span><span class="sxs-lookup"><span data-stu-id="670d9-464">
         - File</span></span><br><span data-ttu-id="670d9-465">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-465">
         - HtmlCoercion</span></span><br><span data-ttu-id="670d9-466">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-466">
         - ImageCoercion</span></span><br><span data-ttu-id="670d9-467">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-467">
         - MatrixBindings</span></span><br><span data-ttu-id="670d9-468">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-468">
         - MatrixCoercion</span></span><br><span data-ttu-id="670d9-469">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-469">
         - OoxmlCoercion</span></span><br><span data-ttu-id="670d9-470">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="670d9-470">
         - PdfFile</span></span><br><span data-ttu-id="670d9-471">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-471">
         - Selection</span></span><br><span data-ttu-id="670d9-472">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-472">
         - Settings</span></span><br><span data-ttu-id="670d9-473">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-473">
         - TableBindings</span></span><br><span data-ttu-id="670d9-474">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-474">
         - TableCoercion</span></span><br><span data-ttu-id="670d9-475">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-475">
         - TextBindings</span></span><br><span data-ttu-id="670d9-476">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-476">
         - TextCoercion</span></span><br><span data-ttu-id="670d9-477">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="670d9-477">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-478">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="670d9-478">Office for iPad</span></span></td>
    <td> <span data-ttu-id="670d9-479">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-479">- TaskPane</span></span></td>
    <td> <span data-ttu-id="670d9-480">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-480">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="670d9-481">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="670d9-481">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="670d9-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="670d9-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="670d9-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="670d9-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="670d9-484">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-484">- BindingEvents</span></span><br><span data-ttu-id="670d9-485">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="670d9-485">
         - CompressedFile</span></span><br><span data-ttu-id="670d9-486">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="670d9-486">
         - CustomXmlParts</span></span><br><span data-ttu-id="670d9-487">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-487">
         - DocumentEvents</span></span><br><span data-ttu-id="670d9-488">
         - File</span><span class="sxs-lookup"><span data-stu-id="670d9-488">
         - File</span></span><br><span data-ttu-id="670d9-489">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-489">
         - HtmlCoercion</span></span><br><span data-ttu-id="670d9-490">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-490">
         - ImageCoercion</span></span><br><span data-ttu-id="670d9-491">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-491">
         - MatrixBindings</span></span><br><span data-ttu-id="670d9-492">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-492">
         - MatrixCoercion</span></span><br><span data-ttu-id="670d9-493">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-493">
         - OoxmlCoercion</span></span><br><span data-ttu-id="670d9-494">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="670d9-494">
         - PdfFile</span></span><br><span data-ttu-id="670d9-495">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-495">
         - Selection</span></span><br><span data-ttu-id="670d9-496">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-496">
         - Settings</span></span><br><span data-ttu-id="670d9-497">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-497">
         - TableBindings</span></span><br><span data-ttu-id="670d9-498">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-498">
         - TableCoercion</span></span><br><span data-ttu-id="670d9-499">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-499">
         - TextBindings</span></span><br><span data-ttu-id="670d9-500">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-500">
         - TextCoercion</span></span><br><span data-ttu-id="670d9-501">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="670d9-501">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-502">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="670d9-502">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="670d9-503">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-503">- TaskPane</span></span><br><span data-ttu-id="670d9-504">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-504">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="670d9-505">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-505">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="670d9-506">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="670d9-506">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="670d9-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="670d9-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="670d9-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="670d9-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="670d9-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-509">- BindingEvents</span></span><br><span data-ttu-id="670d9-510">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="670d9-510">
         - CompressedFile</span></span><br><span data-ttu-id="670d9-511">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="670d9-511">
         - CustomXmlParts</span></span><br><span data-ttu-id="670d9-512">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-512">
         - DocumentEvents</span></span><br><span data-ttu-id="670d9-513">
         - File</span><span class="sxs-lookup"><span data-stu-id="670d9-513">
         - File</span></span><br><span data-ttu-id="670d9-514">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-514">
         - HtmlCoercion</span></span><br><span data-ttu-id="670d9-515">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-515">
         - ImageCoercion</span></span><br><span data-ttu-id="670d9-516">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-516">
         - MatrixBindings</span></span><br><span data-ttu-id="670d9-517">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-517">
         - MatrixCoercion</span></span><br><span data-ttu-id="670d9-518">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-518">
         - OoxmlCoercion</span></span><br><span data-ttu-id="670d9-519">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="670d9-519">
         - PdfFile</span></span><br><span data-ttu-id="670d9-520">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-520">
         - Selection</span></span><br><span data-ttu-id="670d9-521">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-521">
         - Settings</span></span><br><span data-ttu-id="670d9-522">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-522">
         - TableBindings</span></span><br><span data-ttu-id="670d9-523">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-523">
         - TableCoercion</span></span><br><span data-ttu-id="670d9-524">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-524">
         - TextBindings</span></span><br><span data-ttu-id="670d9-525">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-525">
         - TextCoercion</span></span><br><span data-ttu-id="670d9-526">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="670d9-526">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-527">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="670d9-527">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="670d9-528">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-528">- TaskPane</span></span><br><span data-ttu-id="670d9-529">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-529">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="670d9-530">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-530">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="670d9-531">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="670d9-531">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="670d9-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="670d9-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="670d9-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="670d9-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="670d9-534">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-534">- BindingEvents</span></span><br><span data-ttu-id="670d9-535">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="670d9-535">
         - CompressedFile</span></span><br><span data-ttu-id="670d9-536">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="670d9-536">
         - CustomXmlParts</span></span><br><span data-ttu-id="670d9-537">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-537">
         - DocumentEvents</span></span><br><span data-ttu-id="670d9-538">
         - File</span><span class="sxs-lookup"><span data-stu-id="670d9-538">
         - File</span></span><br><span data-ttu-id="670d9-539">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-539">
         - HtmlCoercion</span></span><br><span data-ttu-id="670d9-540">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-540">
         - ImageCoercion</span></span><br><span data-ttu-id="670d9-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-541">
         - MatrixBindings</span></span><br><span data-ttu-id="670d9-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="670d9-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="670d9-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="670d9-544">
         - PdfFile</span></span><br><span data-ttu-id="670d9-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-545">
         - Selection</span></span><br><span data-ttu-id="670d9-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-546">
         - Settings</span></span><br><span data-ttu-id="670d9-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-547">
         - TableBindings</span></span><br><span data-ttu-id="670d9-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-548">
         - TableCoercion</span></span><br><span data-ttu-id="670d9-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="670d9-549">
         - TextBindings</span></span><br><span data-ttu-id="670d9-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-550">
         - TextCoercion</span></span><br><span data-ttu-id="670d9-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="670d9-551">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="670d9-552">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="670d9-552">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="670d9-553">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="670d9-553">Platform</span></span></th>
    <th><span data-ttu-id="670d9-554">拡張点</span><span class="sxs-lookup"><span data-stu-id="670d9-554">Extension points</span></span></th>
    <th><span data-ttu-id="670d9-555">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="670d9-555">API requirement sets</span></span></th>
    <th><span data-ttu-id="670d9-556"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="670d9-556"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-557">Office Online</span><span class="sxs-lookup"><span data-stu-id="670d9-557">Office Online</span></span></td>
    <td> <span data-ttu-id="670d9-558">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="670d9-558">- Content</span></span><br><span data-ttu-id="670d9-559">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-559">
         - TaskPane</span></span><br><span data-ttu-id="670d9-560">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-560">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="670d9-561">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-561">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="670d9-562">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="670d9-562">- ActiveView</span></span><br><span data-ttu-id="670d9-563">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="670d9-563">
         - CompressedFile</span></span><br><span data-ttu-id="670d9-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-564">
         - DocumentEvents</span></span><br><span data-ttu-id="670d9-565">
         - File</span><span class="sxs-lookup"><span data-stu-id="670d9-565">
         - File</span></span><br><span data-ttu-id="670d9-566">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-566">
         - ImageCoercion</span></span><br><span data-ttu-id="670d9-567">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="670d9-567">
         - PdfFile</span></span><br><span data-ttu-id="670d9-568">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-568">
         - Selection</span></span><br><span data-ttu-id="670d9-569">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-569">
         - Settings</span></span><br><span data-ttu-id="670d9-570">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-570">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-571">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="670d9-571">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="670d9-572">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="670d9-572">- Content</span></span><br><span data-ttu-id="670d9-573">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-573">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="670d9-574">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="670d9-574">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="670d9-575">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="670d9-575">- ActiveView</span></span><br><span data-ttu-id="670d9-576">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="670d9-576">
         - CompressedFile</span></span><br><span data-ttu-id="670d9-577">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-577">
         - DocumentEvents</span></span><br><span data-ttu-id="670d9-578">
         - File</span><span class="sxs-lookup"><span data-stu-id="670d9-578">
         - File</span></span><br><span data-ttu-id="670d9-579">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-579">
         - ImageCoercion</span></span><br><span data-ttu-id="670d9-580">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="670d9-580">
         - PdfFile</span></span><br><span data-ttu-id="670d9-581">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-581">
         - Selection</span></span><br><span data-ttu-id="670d9-582">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-582">
         - Settings</span></span><br><span data-ttu-id="670d9-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-583">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-584">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="670d9-584">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="670d9-585">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="670d9-585">- Content</span></span><br><span data-ttu-id="670d9-586">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-586">
         - TaskPane</span></span><br><span data-ttu-id="670d9-587">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-587">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="670d9-588">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-588">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="670d9-589">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="670d9-589">- ActiveView</span></span><br><span data-ttu-id="670d9-590">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="670d9-590">
         - CompressedFile</span></span><br><span data-ttu-id="670d9-591">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-591">
         - DocumentEvents</span></span><br><span data-ttu-id="670d9-592">
         - File</span><span class="sxs-lookup"><span data-stu-id="670d9-592">
         - File</span></span><br><span data-ttu-id="670d9-593">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-593">
         - ImageCoercion</span></span><br><span data-ttu-id="670d9-594">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="670d9-594">
         - PdfFile</span></span><br><span data-ttu-id="670d9-595">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-595">
         - Selection</span></span><br><span data-ttu-id="670d9-596">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-596">
         - Settings</span></span><br><span data-ttu-id="670d9-597">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-597">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-598">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="670d9-598">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="670d9-599">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="670d9-599">- Content</span></span><br><span data-ttu-id="670d9-600">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-600">
         - TaskPane</span></span><br><span data-ttu-id="670d9-601">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-601">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="670d9-602">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-602">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="670d9-603">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="670d9-603">- ActiveView</span></span><br><span data-ttu-id="670d9-604">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="670d9-604">
         - CompressedFile</span></span><br><span data-ttu-id="670d9-605">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-605">
         - DocumentEvents</span></span><br><span data-ttu-id="670d9-606">
         - File</span><span class="sxs-lookup"><span data-stu-id="670d9-606">
         - File</span></span><br><span data-ttu-id="670d9-607">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-607">
         - ImageCoercion</span></span><br><span data-ttu-id="670d9-608">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="670d9-608">
         - PdfFile</span></span><br><span data-ttu-id="670d9-609">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-609">
         - Selection</span></span><br><span data-ttu-id="670d9-610">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-610">
         - Settings</span></span><br><span data-ttu-id="670d9-611">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-611">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-612">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="670d9-612">Office for iPad</span></span></td>
    <td> <span data-ttu-id="670d9-613">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="670d9-613">- Content</span></span><br><span data-ttu-id="670d9-614">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-614">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="670d9-615">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-615">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="670d9-616">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="670d9-616">- ActiveView</span></span><br><span data-ttu-id="670d9-617">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="670d9-617">
         - CompressedFile</span></span><br><span data-ttu-id="670d9-618">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-618">
         - DocumentEvents</span></span><br><span data-ttu-id="670d9-619">
         - File</span><span class="sxs-lookup"><span data-stu-id="670d9-619">
         - File</span></span><br><span data-ttu-id="670d9-620">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="670d9-620">
         - PdfFile</span></span><br><span data-ttu-id="670d9-621">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-621">
         - Selection</span></span><br><span data-ttu-id="670d9-622">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-622">
         - Settings</span></span><br><span data-ttu-id="670d9-623">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-623">
         - TextCoercion</span></span><br><span data-ttu-id="670d9-624">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-624">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-625">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="670d9-625">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="670d9-626">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="670d9-626">- Content</span></span><br><span data-ttu-id="670d9-627">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-627">
         - TaskPane</span></span><br><span data-ttu-id="670d9-628">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-628">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="670d9-629">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-629">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="670d9-630">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="670d9-630">- ActiveView</span></span><br><span data-ttu-id="670d9-631">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="670d9-631">
         - CompressedFile</span></span><br><span data-ttu-id="670d9-632">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-632">
         - DocumentEvents</span></span><br><span data-ttu-id="670d9-633">
         - File</span><span class="sxs-lookup"><span data-stu-id="670d9-633">
         - File</span></span><br><span data-ttu-id="670d9-634">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-634">
         - ImageCoercion</span></span><br><span data-ttu-id="670d9-635">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="670d9-635">
         - PdfFile</span></span><br><span data-ttu-id="670d9-636">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-636">
         - Selection</span></span><br><span data-ttu-id="670d9-637">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-637">
         - Settings</span></span><br><span data-ttu-id="670d9-638">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-638">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-639">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="670d9-639">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="670d9-640">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="670d9-640">- Content</span></span><br><span data-ttu-id="670d9-641">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-641">
         - TaskPane</span></span><br><span data-ttu-id="670d9-642">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-642">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="670d9-643">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-643">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="670d9-644">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="670d9-644">- ActiveView</span></span><br><span data-ttu-id="670d9-645">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="670d9-645">
         - CompressedFile</span></span><br><span data-ttu-id="670d9-646">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-646">
         - DocumentEvents</span></span><br><span data-ttu-id="670d9-647">
         - File</span><span class="sxs-lookup"><span data-stu-id="670d9-647">
         - File</span></span><br><span data-ttu-id="670d9-648">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-648">
         - ImageCoercion</span></span><br><span data-ttu-id="670d9-649">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="670d9-649">
         - PdfFile</span></span><br><span data-ttu-id="670d9-650">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-650">
         - Selection</span></span><br><span data-ttu-id="670d9-651">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-651">
         - Settings</span></span><br><span data-ttu-id="670d9-652">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-652">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="670d9-653">OneNote</span><span class="sxs-lookup"><span data-stu-id="670d9-653">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="670d9-654">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="670d9-654">Platform</span></span></th>
    <th><span data-ttu-id="670d9-655">拡張点</span><span class="sxs-lookup"><span data-stu-id="670d9-655">Extension points</span></span></th>
    <th><span data-ttu-id="670d9-656">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="670d9-656">API requirement sets</span></span></th>
    <th><span data-ttu-id="670d9-657"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="670d9-657"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-658">Office Online</span><span class="sxs-lookup"><span data-stu-id="670d9-658">Office Online</span></span></td>
    <td> <span data-ttu-id="670d9-659">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="670d9-659">- Content</span></span><br><span data-ttu-id="670d9-660">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-660">
         - TaskPane</span></span><br><span data-ttu-id="670d9-661">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="670d9-661">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="670d9-662">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-662">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="670d9-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="670d9-664">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="670d9-664">- DocumentEvents</span></span><br><span data-ttu-id="670d9-665">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-665">
         - HtmlCoercion</span></span><br><span data-ttu-id="670d9-666">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-666">
         - ImageCoercion</span></span><br><span data-ttu-id="670d9-667">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="670d9-667">
         - Settings</span></span><br><span data-ttu-id="670d9-668">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-668">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="670d9-669">Project</span><span class="sxs-lookup"><span data-stu-id="670d9-669">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="670d9-670">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="670d9-670">Platform</span></span></th>
    <th><span data-ttu-id="670d9-671">拡張点</span><span class="sxs-lookup"><span data-stu-id="670d9-671">Extension points</span></span></th>
    <th><span data-ttu-id="670d9-672">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="670d9-672">API requirement sets</span></span></th>
    <th><span data-ttu-id="670d9-673"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="670d9-673"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-674">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="670d9-674">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="670d9-675">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-675">- TaskPane</span></span></td>
    <td> <span data-ttu-id="670d9-676">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-676">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="670d9-677">- Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-677">- Selection</span></span><br><span data-ttu-id="670d9-678">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-678">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-679">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="670d9-679">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="670d9-680">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-680">- TaskPane</span></span></td>
    <td> <span data-ttu-id="670d9-681">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-681">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="670d9-682">- Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-682">- Selection</span></span><br><span data-ttu-id="670d9-683">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-683">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="670d9-684">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="670d9-684">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="670d9-685">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="670d9-685">- TaskPane</span></span></td>
    <td> <span data-ttu-id="670d9-686">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="670d9-686">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="670d9-687">- Selection</span><span class="sxs-lookup"><span data-stu-id="670d9-687">- Selection</span></span><br><span data-ttu-id="670d9-688">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="670d9-688">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="670d9-689">関連項目</span><span class="sxs-lookup"><span data-stu-id="670d9-689">See also</span></span>

- [<span data-ttu-id="670d9-690">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="670d9-690">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="670d9-691">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="670d9-691">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="670d9-692">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="670d9-692">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="670d9-693">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="670d9-693">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
